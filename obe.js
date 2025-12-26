
/* =========================
   Config & utility
   ========================= */
const $ = id => document.getElementById(id);
const FIELD_IDS = [
  'deptname','coordinator','department','acyear','ugpg',
  'semester','section','coursecode','coursename','numcos','numstudents'
];
// localStorage keys for parsed data
const KEY_CIE = 'obeExcelDataCIE';
const KEY_SEE = 'obeExcelDataSEE';

/* =========================
   Restore form inputs & targets on load
   ========================= */
window.addEventListener('DOMContentLoaded', () => {
  // restore form values (localStorage)
  FIELD_IDS.forEach(id => {
    const v = localStorage.getItem(id);
    if (v !== null && $(id)) $(id).value = v;
  });

  // ensure numcos has a valid default
  if (!$('numcos').value) $('numcos').value = 4;

  // build target inputs for both sections
  rebuildAllTargetInputs();

  // restore previously saved targets into created inputs
  restoreTargets();

  // restore parsed uploads (from localStorage) and render if present
  restoreParsedData(KEY_CIE, 'CIE');
  restoreParsedData(KEY_SEE, 'SEE');
});

/* =========================
   Persist form inputs as user types (localStorage)
   ========================= */
FIELD_IDS.forEach(id => {
  const el = $(id);
  if (!el) return;
  el.addEventListener('input', () => {
    localStorage.setItem(id, el.value);
    if (id === 'numcos') {
      // rebuild target inputs when CO count changes
      rebuildAllTargetInputs();
      restoreTargets(); // restore any saved values for new inputs
    }
  });
});

/* =========================
   Target inputs creation / restoration
   ========================= */
function rebuildAllTargetInputs() {
  const n = Math.max(1, Math.min(12, parseInt($('numcos').value || 4, 10)));
  buildTargets('targetsContainerCIE', 'targetCIE', n);
  buildTargets('targetsContainerSEE', 'targetSEE', n);
}

// create n inputs inside container (ids prefix_co1..prefix_coN)
function buildTargets(containerId, prefix, n) {
  const cont = $(containerId);
  cont.innerHTML = '';
  for (let i = 1; i <= n; i++) {
    const wrapper = document.createElement('div');
    wrapper.innerHTML =
      `<label>CO${i} Target</label><br>` +
      `<input type="number" min="0" step="any" id="${prefix}_co${i}" />`;
    cont.appendChild(wrapper);

    // save on change
    const input = wrapper.querySelector('input');
    input.addEventListener('input', () => {
      localStorage.setItem(`${prefix}_co${i}`, input.value);
    });
  }
}

// restore saved target values (localStorage)
function restoreTargets() {
  const n = Math.max(1, Math.min(12, parseInt($('numcos').value || 4, 10)));
  ['targetCIE', 'targetSEE'].forEach(prefix => {
    for (let i = 1; i <= n; i++) {
      const key = `${prefix}_co${i}`;
      const saved = localStorage.getItem(key);
      if (saved !== null) {
        const el = $(key);
        if (el) el.value = saved;
      }
    }
  });
}

/* =========================
   Template download (dynamic)
   ========================= */
$('downloadTemplate').addEventListener('click', () => {
  const n = Math.max(1, Math.min(12, parseInt($('numcos').value || 4, 10)));
  const students = Math.max(1, parseInt($('numstudents').value || 18, 10));
  const headers = ['SLNO', 'USN', 'Student Name'];
  for (let i = 1; i <= n; i++) headers.push('CO' + i);

  const rows = [headers];
  for (let r = 1; r <= students; r++) {
    const row = [r, '', ''];
    for (let i = 1; i <= n; i++) row.push('');
    rows.push(row);
  }

  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Template');
  const fname = `${$('coursecode').value || 'OBE_Template'}_CO_Template.xlsx`;
  XLSX.writeFile(wb, fname);
});

/* =========================
   Excel parsing + processing
   (shared between CIE & SEE)
   ========================= */
function setupParser(prefix) {
  // prefix 'CIE' or 'SEE'
  $(`parseFile${prefix}`).addEventListener('click', () => {
    const input = $(`fileInput${prefix}`);
    const f = input.files[0];
    if (!f) { alert('Please choose an Excel file first.'); return; }

    const reader = new FileReader();
    reader.onload = (e) => {
      let wb;
      try {
        wb = XLSX.read(e.target.result, { type: 'binary' });
      } catch (err) {
        alert('Error reading file: ' + err);
        return;
      }
      const ws = wb.Sheets[wb.SheetNames[0]];
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (!aoa || aoa.length < 2) { alert('Uploaded sheet looks empty (need header + data).'); return; }
      processAOA(prefix, aoa);
    };
    reader.readAsBinaryString(f);
  });
}
['CIE', 'SEE'].forEach(setupParser);

// process AOA -> students array and compute/save/results
function processAOA(prefix, aoa) {
  const headers = aoa[0].map(h => String(h).trim());
  // detect CO columns by headers like CO1, CO2...
  const numCOs = headers.filter(h => /^CO\s*\d+/i.test(h)).length;
  if (numCOs === 0) { alert('No CO columns detected in header (expect CO1, CO2,…).'); return; }

  // collect targets from session inputs
  const targets = [];
  for (let i = 1; i <= numCOs; i++) {
    const t = parseFloat(($(`target${prefix}_co${i}`) || {}).value || '0');
    targets.push(isNaN(t) ? 0 : t);
  }

  const students = [];
  for (let r = 1; r < aoa.length; r++) {
    const row = aoa[r];
    if (!row || row.length === 0) continue;
    const slno = row[0] || r;
    const usn = row[1] || '';
    const name = row[2] || '';
    const cos = [];
    for (let ci = 0; ci < numCOs; ci++) {
      const val = parseFloat(row[3 + ci]);
      cos.push(isNaN(val) ? 0 : val);
    }
    students.push({ slno, usn, name, cos });
  }

  if (students.length === 0) { alert('No student data rows found.'); return; }

  // save to localStorage
  const key = (prefix === 'CIE') ? KEY_CIE : KEY_SEE;
  localStorage.setItem(key, JSON.stringify({ headers, students, targets }));

  // --------- compute summary ---------
const summary = computeSummary(students, targets);

// --------- STORE CIE/SEE percentages for auto-fill in final section ---------
if (prefix === 'CIE') {
    localStorage.setItem("ciePercentages", JSON.stringify(summary.percent));
}
if (prefix === 'SEE') {
    localStorage.setItem("seePercentages", JSON.stringify(summary.percent));
}

// --------- RENDER UPLOADED TABLE + SUMMARY ---------
renderUploaded(prefix, headers, students, numCOs);
renderSummary(prefix, summary, targets);
}

/* =========================
   Render helpers
   ========================= */
function renderUploaded(prefix, headers, students, numCOs) {
  const div = $(`uploadedTable${prefix}`);
  let html = '<table><thead><tr>';
  html += headers.slice(0, 3 + numCOs).map(h => `<th>${h}</th>`).join('');
  html += '</tr></thead><tbody>';
  students.forEach(s => {
    html += `<tr><td>${s.slno}</td><td>${s.usn}</td><td style="text-align:left">${s.name}</td>`;
    s.cos.forEach(v => html += `<td>${v}</td>`);
    html += '</tr>';
  });
  html += '</tbody></table>';
  div.innerHTML = html;


}

function renderSummary(prefix, summary, targets) {
  const div = $(`summaryTable${prefix}`);
  let html = '<table><thead><tr><th></th>';
  for (let i = 1; i <= summary.nCOs; i++) html += `<th>CO${i}</th>`;
  html += '</tr></thead><tbody>';
  html += `<tr><td><b>Target</b></td>${targets.map(t => `<td>${t}</td>`).join('')}</tr>`;
  html += `<tr><td><b>Number of students scored more than target value</b></td>${summary.counts.map(c => `<td>${c}</td>`).join('')}</tr>`;
  html += `<tr><td><b>Percentage of attainment</b></td>${summary.percent.map(p => `<td>${p}%</td>`).join('')}</tr>`;
  html += `<tr><td><b>CO Attainment Level</b></td>${summary.level.map(l => `<td>${l}</td>`).join('')}</tr>`;
  html += '</tbody></table>';
  div.innerHTML = html;
}

/* =========================
   Compute summary
   ========================= */
function computeSummary(students, targets) {
  const nStudents = students.length;
  const nCOs = targets.length;
  const counts = Array(nCOs).fill(0);
  students.forEach(s => {
    s.cos.forEach((v, i) => {
      const val = Number(v) || 0;
      if (val >= targets[i]) counts[i]++;
    });
  });
  const percent = counts.map(c => +((c / nStudents) * 100).toFixed(2));
  //const level = percent.map(p => (p >= 80 ? 3 : (p >= 70 ? 2 : 1)));
const level = percent.map(p =>
  p < 60 ? 0 : p >= 80 ? 3 : p >= 70 ? 2 : 1);

  return { nStudents, nCOs, counts, percent, level };
}

/* =========================
   Restore parsed data on load (localStorage)
   ========================= */
function restoreParsedData(storageKey, prefix) {
  const raw = localStorage.getItem(storageKey);
  if (!raw) return;
  try {
    const { headers, students, targets } = JSON.parse(raw);
    const numCOs = targets.length;
    // render
    renderUploaded(prefix, headers, students, numCOs);
    renderSummary(prefix, computeSummary(students, targets), targets);
  } catch (e) {
    console.error('Error restoring parsed data', e);
  }
}

/* =========================
   Print helpers (embed header image as base64)
   ========================= */
async function embedHeaderImageBase64() {
  const img = $('uniHeader');
  if (!img || !img.src || img.style.display === 'none') return '';
  try {
    const resp = await fetch(img.src);
    const blob = await resp.blob();
    return await new Promise(resolve => {
      const r = new FileReader();
      r.onloadend = () => resolve(`<div style="text-align:center;margin-bottom:8px;"><img src="${r.result}" style="width:100%;max-height:110px;object-fit:contain"/></div>`);
      r.readAsDataURL(blob);
    });
  } catch (e) {
    console.warn('Could not load header image for print', e);
    return '';
  }
}

function buildMetaHTML() {
  // show selected course metadata (from session)
  const rows = FIELD_IDS.map(k => {
    const v = localStorage.getItem(k) || '';
    const label = ({deptname:'Department Name',coordinator:'Course Coordinator',department:'Department',acyear:'Academic Year',
      ugpg:'UG/PG',semester:'Semester',section:'Section',coursecode:'Course Code',coursename:'Course Name',numcos:'Number of COs',numstudents:'Number of Students'})[k] || k;
    return `<tr><td style="font-weight:600;padding:6px;border:1px solid #ddd">${label}</td><td style="padding:6px;border:1px solid #ddd">${v}</td></tr>`;
  }).join('');
  return `<table style="width:100%;border-collapse:collapse;margin-bottom:8px">${rows}</table>`;
}

async function printSection(prefix) {
  // ensure there is data to print
  const uploadedDiv = $(`uploadedTable${prefix}`);
  const summaryDiv = $(`summaryTable${prefix}`);
  if (!uploadedDiv || !summaryDiv || !uploadedDiv.innerHTML) {
    if (!confirm('No uploaded table found for '+prefix+'. Still open print page?')) return;
  }

  const imgHTML = await embedHeaderImageBase64();
  const metaHTML = buildMetaHTML();
  const tableHTML = uploadedDiv ? uploadedDiv.innerHTML : '';
  const summaryHTML = summaryDiv ? summaryDiv.innerHTML : '';

  const w = window.open('', '_blank');
  w.document.write(`
    <html><head><title>${prefix} Report</title>
      <style>body{font-family:Arial;margin:30px}table{border-collapse:collapse;width:100%;}th,td{border:1px solid #000;padding:6px;text-align:center}</style>
    </head>
    <body>${imgHTML}<h2 style="text-align:center">${prefix} OBE Report</h2>${metaHTML}${tableHTML}${summaryHTML}</body></html>
  `);
  w.document.close();
  w.onload = () => setTimeout(()=>w.print(), 600);
}

/* Attach print buttons */
//$('printReportCIE').addEventListener('click', ()=>printSection('CIE'));
//$('printReportSEE').addEventListener('click', ()=>printSection('SEE'));

/* =========================
   Restore localStorage parsed data on load
   ========================= */
window.addEventListener('DOMContentLoaded', () => {
  restoreParsedData(KEY_CIE, 'CIE');
  restoreParsedData(KEY_SEE, 'SEE');
});

/* =========================
   Reset All (clear both storages and UI)
   ========================= */
$('resetAll').addEventListener('click', ()=>{
  if(!confirm('Clear all stored data ?')) return;
  localStorage.clear();
  localStorage.removeItem(KEY_CIE);
  localStorage.removeItem(KEY_SEE);
  localStorage.removeItem("CO_RESULTS");

  // clear inputs
  FIELD_IDS.forEach(id => { if($(id)) $(id).value = ''; });
  // clear target containers + results
  ['targetsContainerCIE','targetsContainerSEE','uploadedTableCIE','summaryTableCIE','uploadedTableSEE','summaryTableSEE']
    .forEach(id => { if($(id)) $(id).innerHTML = ''; });
  // recreate targets with default value
  $('numcos').value = 4;
  rebuildAllTargetInputs();
  alert('Cleared.');
});

/* =========================
   Utility: rebuild targets (small helper)
   ========================= */
function rebuildAllTargetInputs() {
  const n = Math.max(1, Math.min(12, parseInt($('numcos').value || 4, 10)));
  buildTargets('targetsContainerCIE','targetCIE', n);
  buildTargets('targetsContainerSEE','targetSEE', n);
}

// also rebuild targets when numcos input changed by user and store the new value
$('numcos').addEventListener('change', ()=>{
  const n = Math.max(1, Math.min(12, parseInt($('numcos').value || 4, 10)));
  $('numcos').value = n;
  localStorage.setItem('numcos', String(n));
  rebuildAllTargetInputs();
  restoreTargets();
});







/*
  Standalone code:
  - Reads numCOs from localStorage.numcos (if exists), else defaults to 4.
  - Generates inputs: CIE%, SEE%, IndirectScore(1-5) per CO.
  - Persists all inputs in localStorage (so refresh keeps values).
  - Computes DirectPercent = 0.5*CIE + 0.5*SEE
  - Final = DirectPercent*(directWeight/100) + IndirectScore*(indirectWeight/100)
  - Draws bar chart: DirectPercent and Final for each CO.
*/

// ---------- configuration ----------
const storedNum = parseInt(localStorage.getItem('numcos') || '4', 10);
const numCOs = (Number.isInteger(storedNum) && storedNum >= 1) ? storedNum : 4;

// prefixes used in localStorage keys:
// cie_CO1, see_CO1, indirect_CO1
const cieKey = i => `cie_CO${i}`;
const seeKey = i => `see_CO${i}`;
const indKey = i => `indirect_CO${i}`;
const directWeightKey = 'directWeight';

// Chart instance
let coChart = null;

// ---------- DOM refs ----------
const coGrid = document.getElementById('coGrid');
const directInput = document.getElementById('directWeight');
const indirectInput = document.getElementById('indirectWeight');
const computeBtn = document.getElementById('computeBtn');
const clearBtn = document.getElementById('clearBtn');
const resultsSection = document.getElementById('resultsSection');
const resultTableWrap = document.getElementById('resultTableWrap');

// ---------- initialize weights from session (if any) ----------
if (localStorage.getItem(directWeightKey)) {
  directInput.value = localStorage.getItem(directWeightKey);
}
updateIndirectAuto(); // set indirect = 100 - direct

// keep direct weight in session when changed, and update indirect auto
directInput.addEventListener('input', () => {
  let v = Number(directInput.value) || 0;
  if (v < 0) v = 0;
  if (v > 100) v = 100;
  directInput.value = v;
  localStorage.setItem(directWeightKey, String(v));
  updateIndirectAuto();
});

// recompute indirectWeight display
function updateIndirectAuto(){
  const d = Number(directInput.value) || 0;
  const ind = Math.max(0, 100 - d);
  indirectInput.value = ind;
  localStorage.setItem('indirectWeight', String(ind));
}

// ---------- build CO inputs grid (with auto-fill from uploaded CIE & SEE results) ----------
function buildCOInputs() {

  coGrid.innerHTML = '';

  // ---- FETCH STORED CIE & SEE percentages from previous uploads ----
  const cieArray = JSON.parse(localStorage.getItem("ciePercentages") || "[]");
  const seeArray = JSON.parse(localStorage.getItem("seePercentages") || "[]");

  // headings
  const headings = ['CO','CIE (%)','SEE (%)','Indirect (1-5)',''];
  for (let h of headings) {
    const div = document.createElement('div');
    div.className = 'heading';
    div.textContent = h;
    coGrid.appendChild(div);
  }

  // rows
  for (let i = 1; i <= numCOs; i++) {

    // CO label
    const coLabel = document.createElement('div');
    coLabel.style.textAlign = 'center';
    coLabel.textContent = 'CO' + i;
    coGrid.appendChild(coLabel);

    // ====== CIE INPUT ======
    const cieDiv = document.createElement('div');
    const cieInput = document.createElement('input');
    cieInput.type = 'number';
    cieInput.min = 0; 
    cieInput.max = 100; 
    cieInput.step = 'any';
    cieInput.id = cieKey(i);
    cieInput.placeholder = 'e.g. 88.89';

    // AUTO-FILL FROM UPLOADED DATA → overrides stored user-typed value
    if (cieArray[i - 1] !== undefined) {
      cieInput.value = cieArray[i - 1];
      localStorage.setItem(cieKey(i), cieArray[i - 1]);
    }
    // else restore typed value from localStorage
    else if (localStorage.getItem(cieKey(i)) !== null) {
      cieInput.value = localStorage.getItem(cieKey(i));
    }

    // save on input
    cieInput.addEventListener('input', () =>
      localStorage.setItem(cieKey(i), cieInput.value || '0')
    );
    cieDiv.appendChild(cieInput);
    coGrid.appendChild(cieDiv);


    // ====== SEE INPUT ======
    const seeDiv = document.createElement('div');
    const seeInput = document.createElement('input');
    seeInput.type = 'number';
    seeInput.min = 0; 
    seeInput.max = 100; 
    seeInput.step = 'any';
    seeInput.id = seeKey(i);
    seeInput.placeholder = 'e.g. 77.78';

    // AUTO-FILL FROM UPLOADED DATA → overrides stored user value
    if (seeArray[i - 1] !== undefined) {
      seeInput.value = seeArray[i - 1];
      localStorage.setItem(seeKey(i), seeArray[i - 1]);
    }
    // else restore from localStorage
    else if (localStorage.getItem(seeKey(i)) !== null) {
      seeInput.value = localStorage.getItem(seeKey(i));
    }

    // save on input
    seeInput.addEventListener('input', () =>
      localStorage.setItem(seeKey(i), seeInput.value || '0')
    );
    seeDiv.appendChild(seeInput);
    coGrid.appendChild(seeDiv);


    // ====== INDIRECT INPUT (1–5) ======
    const indDiv = document.createElement('div');
    const indInput = document.createElement('input');
    indInput.type = 'number';
    indInput.min = 1; 
    indInput.max = 5; 
    indInput.step = '1';
    indInput.id = indKey(i);
    indInput.placeholder = '1-5';

    // restore indirect from localStorage
    if (localStorage.getItem(indKey(i)) !== null)
      indInput.value = localStorage.getItem(indKey(i));

    indInput.addEventListener('input', () => {
      let v = Number(indInput.value) || 1;
      if (v < 1) v = 1;
      if (v > 5) v = 5;
      indInput.value = v;
      localStorage.setItem(indKey(i), String(v));
    });

    indDiv.appendChild(indInput);
    coGrid.appendChild(indDiv);

    // empty note column
    const note = document.createElement('div');
    note.className = 'small';
    note.style.textAlign = 'center';
    note.textContent = '';
    coGrid.appendChild(note);
  }
}



function computeAndRender() {

  const cieWt = (Number(document.getElementById("cieWeight").value) || 0) / 100;
  const seeWt = (Number(document.getElementById("seeWeight").value) || 0) / 100;

  const directWt = (Number(document.getElementById("directWeight").value) || 0) / 100;
  const indirectWt = (Number(document.getElementById("indirectWeight").value) || 0) / 100;

  const labels = [];
  const cieArr = [], seeArr = [], directPercentArr = [],
        indirectScoreArr = [], indirectPercentArr = [], finalArr = [];

  for (let i = 1; i <= numCOs; i++) {

    const cie = parseFloat(localStorage.getItem(cieKey(i)) ||
                (document.getElementById(cieKey(i)).value || 0)) || 0;

    const see = parseFloat(localStorage.getItem(seeKey(i)) ||
                (document.getElementById(seeKey(i)).value || 0)) || 0;

    const indirectScore = parseFloat(localStorage.getItem(indKey(i)) ||
                (document.getElementById(indKey(i)).value || 0)) || 0;

    // NEW LOGIC
    const directPercent = (cie * cieWt) + (see * seeWt);
    const indirectPercent = (indirectScore / 5) * 100;

    const final = (directPercent * directWt) + (indirectPercent * indirectWt);

    labels.push("CO" + i);
    cieArr.push(+cie.toFixed(2));
    seeArr.push(+see.toFixed(2));
    directPercentArr.push(+directPercent.toFixed(2));
    indirectScoreArr.push(indirectScore);
    indirectPercentArr.push(+indirectPercent.toFixed(2));
    finalArr.push(+final.toFixed(2));
  }

  renderTable(labels, cieArr, seeArr, directPercentArr, indirectScoreArr, indirectPercentArr, finalArr);
  renderChart(labels, directPercentArr, finalArr);

  resultsSection.style.display = "block";

  storeFinalArr(finalArr);
  buildCOPOTable(finalArr.length);
}




function renderTable(labels, cie, see, directPercent, indirectScore, indirectPercent, finalVals) {
  let html = `<table><thead><tr>
    <th>CO</th>
    <th>CIE (%)</th>
    <th>SEE (%)</th>
    <th>Direct (CIE/SEE)</th>
    <th>Indirect (1–5)</th>
    <th>Indirect (%)</th>
    <th>Final (%)</th>
  </tr></thead><tbody>`;

  for (let i = 0; i < labels.length; i++) {
    html += `<tr>
      <td>${labels[i]}</td>
      <td>${cie[i]}</td>
      <td>${see[i]}</td>
      <td>${directPercent[i]}</td>
      <td>${indirectScore[i]}</td>
      <td>${indirectPercent[i]}</td>
      <td><b>${finalVals[i]}</b></td>
    </tr>`;
  }

  html += `</tbody></table>`;
  resultTableWrap.innerHTML = html;
}








// ---------- chart ----------
function renderChart(labels, directPercentArr, finalArr) {
  const ctx = document.getElementById('coChart').getContext('2d');
  if (coChart) coChart.destroy();
  coChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        { label: 'Direct (50%CIE+50%SEE)', data: directPercentArr, backgroundColor: 'rgba(54,162,235,0.6)' },
        { label: 'Final (Direct*Dw + Indirect*Iw)', data: finalArr, backgroundColor: 'rgba(75,192,192,0.6)' }
      ]
    },
    options: {
      responsive: true,
      scales: { y: { beginAtZero: true } },
      plugins: { legend: { position: 'top' } }
    }
  });
}

// ---------- helpers ----------
function formatNum(n){ return (Number.isFinite(n) ? n : 0).toString(); }

// ---------- load initial UI ----------
buildCOInputs();

// ---------- attach handlers ----------
computeBtn.addEventListener('click', computeAndRender);
clearBtn.addEventListener('click', ()=>{
  // clear only the per-CO persisted inputs and results (keep numcos & direct weight)
  for (let i=1;i<=numCOs;i++){
    localStorage.removeItem(cieKey(i));
    localStorage.removeItem(seeKey(i));
    localStorage.removeItem(indKey(i));
    const el1 = document.getElementById(cieKey(i)); if(el1) el1.value='';
    const el2 = document.getElementById(seeKey(i)); if(el2) el2.value='';
    const el3 = document.getElementById(indKey(i)); if(el3) el3.value='';
  }
  localStorage.removeItem('indirectWeight');
  resultTableWrap.innerHTML='';
  if(coChart) { coChart.destroy(); coChart = null; }
  resultsSection.style.display='none';
});

// optional: auto compute on Enter in weight or inputs
['directWeight'].forEach(id=>{
  const el=document.getElementById(id);
  el.addEventListener('keydown', (e)=>{ if(e.key==='Enter') computeAndRender(); });
});


let GLOBAL_finalArr = [];

// PO / PSO configuration
const PO_COUNT = 12;
const PSO_COUNT = 3;
const TOTAL_PO = PO_COUNT + PSO_COUNT;

const PO_LIST = [
  "PO1","PO2","PO3","PO4","PO5","PO6",
  "PO7","PO8","PO9","PO10","PO11","PO12",
  "PSO1","PSO2","PSO3"
];


// ---------------------------------------------------------
// STORE FINAL CO PERCENTAGES
// ---------------------------------------------------------
function storeFinalArr(finalArr) {
    GLOBAL_finalArr = finalArr.slice(); 
}



// ---------------------------------------------------------
// BUILD INPUT TABLE FOR MAPPING
// ---------------------------------------------------------
function buildCOPOTable(numCOs) {
    const container = document.getElementById("copotable");
    container.innerHTML = "";

    let html = "<table id='coPoTable'><thead><tr><th>CO</th>";

    PO_LIST.forEach(po => html += `<th>${po}</th>`);
    html += "</tr></thead><tbody>";

    // CO rows
    for (let co = 1; co <= numCOs; co++) {
        html += `<tr><td>CO${co}</td>`;
        PO_LIST.forEach(po => {
            const key = `map_${po}_co${co}`;
            const val = localStorage.getItem(key) || "";
            html += `
                <td>
                    <input type="number" min="0" max="3" step="1"
                        id="${key}" value="${val}"
                        oninput="localStorage.setItem('${key}',this.value); updatePOColumnAverages();">
                </td>`;
        });
        html += "</tr>";
    }

    // Average Row
    html += `<tr id="avgRow"><td><b>Average</b></td>`;
    PO_LIST.forEach(() => html += `<td class="avgCell">0</td>`);
    html += "</tr>";

    html += "</tbody></table>";

    container.innerHTML = html;

    updatePOColumnAverages(); 
}



// ---------------------------------------------------------
// UPDATE AVERAGE ROW IN INPUT TABLE
// ---------------------------------------------------------
/* average of all cells 
function updatePOColumnAverages() {
    const numCOs = GLOBAL_finalArr.length;
    let avgCells = document.querySelectorAll("#avgRow .avgCell");

    PO_LIST.forEach((po, index) => {
        let sum = 0;

        for (let co = 1; co <= numCOs; co++) {
            const key = `map_${po}_co${co}`;
            sum += Number(localStorage.getItem(key) || 0);
        }

        const avg = numCOs ? (sum / numCOs).toFixed(2) : "0.00";
        avgCells[index].textContent = avg;
    });
}
*/


//average of entered cells only


function updatePOColumnAverages() {
    const numCOs = GLOBAL_finalArr.length;
    let avgCells = document.querySelectorAll("#avgRow .avgCell");

    PO_LIST.forEach((po, index) => {
        let sum = 0;
        let count = 0; // number of entered CO values

        for (let co = 1; co <= numCOs; co++) {
            const key = `map_${po}_co${co}`;
            const value = Number(localStorage.getItem(key));

            if (value && value > 0) {   // consider only entered cells
                sum += value;
                count++;
            }
        }

        const avg = count > 0 ? (sum / count).toFixed(2) : "0.00";
        avgCells[index].textContent = avg;
    });
}


// ---------------------------------------------------------
// RENDER FINAL OUTPUT TABLE WITH PERCENTAGES
// ---------------------------------------------------------
function renderCOPOTableWithFinal(finalArr, poFinal, poAvg) {
    const numCOs = finalArr.length;
    const table = document.getElementById("COPO_Output_Table");
    table.innerHTML = "";

    // HEADER
    let header = "<tr><th>CO</th>";
    PO_LIST.forEach(po => header += `<th>${po}</th>`);
    header += "</tr>";
    table.innerHTML += header;

    // CO ROWS: each CO × PO multiplied value
    for (let co = 1; co <= numCOs; co++) {
        const finalCO = Number(finalArr[co - 1]) || 0;
        let row = `<tr><td>CO${co} (${finalCO.toFixed(3)}%)</td>`;

        PO_LIST.forEach(po => {
            const key = `map_${po}_co${co}`;
            const level = Number(localStorage.getItem(key) || 0);
            const multiplied = ((finalCO / 100) * level).toFixed(2); // scaled to 0–3
            row += `<td>${multiplied}</td>`;
        });

        row += "</tr>";
        table.innerHTML += row;
    }

    // AVG MAPPING ROW
    // AVG MAPPING ROW (average of multiplied CO×PO values, 2 decimal places)
    let avgRow = "<tr><td><b>Avg Mapping</b></td>";

    PO_LIST.forEach((po, index) => {

        /*
        // ❌ OLD LOGIC (dividing by total COs)
        let sum = 0;

        for (let co = 1; co <= numCOs; co++) {
            const finalCO = Number(finalArr[co - 1]) || 0;
            const key = `map_${po}_co${co}`;
            const level = Number(localStorage.getItem(key) || 0);

            sum += (finalCO / 100) * level; // multiplied value
        }

        const avg = (sum / numCOs).toFixed(2); // round to 2 decimals
        */

        // ✅ NEW LOGIC (divide only by entered COs)
        let sum = 0;
        let count = 0;

        for (let co = 1; co <= numCOs; co++) {
            const finalCO = Number(finalArr[co - 1]) || 0;
            const key = `map_${po}_co${co}`;
            const level = Number(localStorage.getItem(key)) || 0;

            if (level > 0) {
                sum += (finalCO / 100) * level;
                count++;
            }
        }

        const avg = count > 0 ? (sum / count).toFixed(2) : "0.00";
        avgRow += `<td>${avg}</td>`;
    });

    avgRow += "</tr>";
    table.innerHTML += avgRow;

    // EXPECTED PO ROW
    let expectedRow = "<tr><td><b>Expected PO (%)</b></td>";
    poAvg.forEach(v => {
        const percent = ((v / 3) * 100).toFixed(2);
        expectedRow += `<td>${percent}%</td>`;
    });
    expectedRow += "</tr>";
    table.innerHTML += expectedRow;

    // PO FINAL (ATTAINED PO) ROW
    let finalRow = "<tr><td><b>PO Final (Attained PO) (%)</b></td>";
    poFinal.forEach(v => {
        const percent = v.toFixed(2);
        finalRow += `<td>${percent}%</td>`;
    });
    finalRow += "</tr>";
    table.innerHTML += finalRow;
}
// ---------------------------------------------------------
// MAIN CALCULATION — PO ATTAINMENT (percent version)
// ---------------------------------------------------------
function computePOAttainment() {
    if (!GLOBAL_finalArr || GLOBAL_finalArr.length === 0) {
        alert("Final CO percentages (finalArr) not stored yet.");
        return;
    }

    const numCOs = GLOBAL_finalArr.length;

    let poWeightedSum = Array(TOTAL_PO).fill(0);  
    let poWeight = Array(TOTAL_PO).fill(0);       
    let poAvg = Array(TOTAL_PO).fill(0);          
    let poFinal = Array(TOTAL_PO).fill(0);        

    // 1️⃣ Accumulate weighted values
    for (let co = 1; co <= numCOs; co++) {
        const finalCO = Number(GLOBAL_finalArr[co - 1]) || 0;

        PO_LIST.forEach((po, index) => {
            const key = `map_${po}_co${co}`;
            const level = Number(localStorage.getItem(key) || 0);

            poWeightedSum[index] += (finalCO / 100) * level; // scale final % to 0–1
            poWeight[index] += level;
            poAvg[index] += level;
        });
    }

    // 2️⃣ Compute Avg Mapping and store in session
    poAvg = poAvg.map(sum => (sum / numCOs));
    PO_LIST.forEach((po, i) => localStorage.setItem(`avg_${po}`, poAvg[i]));

    // 3️⃣ Compute PO Final (%) as weighted average
    for (let i = 0; i < TOTAL_PO; i++) {
       // poFinal[i] = poWeight[i] > 0 ? poWeightedSum[i] / poWeight[i] : 0;
poFinal[i]=(poWeightedSum[i] /numCOs).toFixed(2)/3*100;
//console.log('po avg',poAvg[i]);
    }

    // 4️⃣ Store expected and attained in session for chart
    storePOForChart(poFinal);

    // 5️⃣ Render final table
    renderCOPOTableWithFinal(GLOBAL_finalArr, poFinal, poAvg);
}

// ---------------------------------------------------------
// STORE EXPECTED AND ATTAINED PO FOR CHART
// ---------------------------------------------------------
function storePOForChart(poFinal) {
    // Expected PO (%) based on mapping averages
  //  const expectedPO = PO_LIST.map(po => {
   //     const avgMapping = Number(localStorage.getItem(`avg_${po}`) || 0);
  //      return avgMapping > 0 ? 100 : 0;
  //  });


const expectedPO = PO_LIST.map(po => {
    const avgMapping = Number(localStorage.getItem(`avg_${po}`) || 0);
    return ((avgMapping / 3) * 100).toFixed(2);
});



    // Attained PO (%) from poFinal (0–1 → 0–100)
    //const attainedPO = poFinal.map(v => (v * 100));
const attainedPO = poFinal.map(v => (v));

    // Store in session
    localStorage.setItem("EXPECTED_PO", JSON.stringify(expectedPO));
    localStorage.setItem("ATTAINED_PO", JSON.stringify(attainedPO));
}

let poChartInstance = null; // global variable

function plotPOChartFromSession() {
    const expectedPO = JSON.parse(localStorage.getItem("EXPECTED_PO") || "[]");
    const attainedPO = JSON.parse(localStorage.getItem("ATTAINED_PO") || "[]");

    if (!expectedPO.length || !attainedPO.length) {
        console.warn("Expected/Attained PO not found in localStorage.");
        return;
    }

    const ctx = document.getElementById('poChart').getContext('2d');

    // Destroy existing chart if it exists
    if (poChartInstance) {
        poChartInstance.destroy();
    }

    poChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: PO_LIST,
            datasets: [
                {
                    label: 'Expected PO (%)',
                    data: expectedPO,
                    backgroundColor: 'rgba(54, 162, 235, 0.6)',
                    borderColor: 'rgba(54, 162, 235, 1)',
                    borderWidth: 1
                },
                {
                    label: 'PO Final (Attained PO) (%)',
                    data: attainedPO,
                    backgroundColor: 'rgba(255, 99, 132, 0.6)',
                    borderColor: 'rgba(255, 99, 132, 1)',
                    borderWidth: 1
                }
            ]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { position: 'top' },
                title: { display: true, text: 'Expected vs Attained PO (%)' }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    max: 100,
                    title: { display: true, text: 'Percentage (%)' }
                }
            }
        }
    });
}

// ---------------------------------------------------------
// USAGE: Call this after finalArr is stored and mapping input
// ---------------------------------------------------------
function computePOAttainmentAndPlot() {
    computePOAttainment(); // calculates poFinal, poAvg, stores session
    plotPOChartFromSession(); // plots chart using session data
}






// Convert grid values into clean printable layout (two columns)
function buildGridPrintLayout() {
  const fields = [
    { label: "Department Name", id: "deptname" },
    { label: "Course Coordinator", id: "coordinator" },
    { label: "Department", id: "department" },
    { label: "Academic Year", id: "acyear" },
    { label: "UG / PG", id: "ugpg" },
    { label: "Semester", id: "semester" },
    { label: "Section", id: "section" },
    { label: "Course Code", id: "coursecode" },
    { label: "Course Name", id: "coursename" }
  ];

  let html = `
    <h2 class="section-title">Course Information</h2>
    <div class="info-grid">`;

  fields.forEach(f => {
    const value = document.getElementById(f.id)?.value || "";
    html += `
      <div class="info-item">
        <strong>${f.label}:</strong> ${value}
      </div>`;
  });

  html += `</div>`;
  return html;
}




function printAll() {
  const chartCanvas = document.getElementById("poChart");
  const chartImage = chartCanvas ? chartCanvas.toDataURL("image/png") : "";

const chartCanvas1 = document.getElementById("coChart");
  const chartImage1 = chartCanvas1 ? chartCanvas1.toDataURL("image/png") : "";


  const elementsToPrint = [
    `<img src="jssstu_header.jpg" style="width:100%; display:block; margin:auto;" />`,
    buildGridPrintLayout(),
    `<h2 class="section-title">CIE</h2>`,
    cloneElementHTML(document.getElementById("uploadedTableCIE")),
    cloneElementHTML(document.getElementById("summaryTableCIE")),
    "PAGEBREAK",
    buildFooterCourseInfo(),
    `<h2 class="section-title">SEE</h2>`,
    cloneElementHTML(document.getElementById("uploadedTableSEE")),
    cloneElementHTML(document.getElementById("summaryTableSEE")),
    "PAGEBREAK",
    buildFooterCourseInfo(),
    cloneElementHTML(document.getElementById("resultsSection")),
`<img src="${chartImage1}" style="width:100%; max-width:900px; display:block; margin:auto;" />`
 ,
    "PAGEBREAK",
    buildFooterCourseInfo(),
    `<h2 class="section-title">CO-PO Expected</h2>`,
    cloneElementHTML(document.getElementById("coPoTable"), true),
    `<h2 class="section-title">CO-PO Attained</h2>`,
    cloneElementHTML(document.getElementById("COPO_Output_Table")),
    "PAGEBREAK",
    buildFooterCourseInfo(),
    `<h2 class="section-title">CO-PO Attainment Chart</h2>`,
    `<img src="${chartImage}" style="width:100%; max-width:900px; display:block; margin:auto;" />`
  ];

  let printContents = "";
  elementsToPrint.forEach(el => {
    if (el === "PAGEBREAK") printContents += '<div class="page-break"></div>';
    else printContents += el;
  });

  const printWindow = window.open("", "_blank", "width=1000,height=800");

  printWindow.document.write(`
    <html>
        <head>
          <title>Print Document</title>
         <style>
body { font-family: 'Times New Roman', serif; font-size: 14px; }

table { border-collapse: collapse; width: 100%; margin-top: 10px; }
table, th, td { border: 1px solid black; padding: 6px; }

.section-title {
  text-align: center;
  font-size: 22px;
  font-weight: bold;
  margin-top: 25px;
}

.info-grid {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 8px;
  margin-top: 10px;
}
.info-item {
  border: 1px solid black;
  padding: 6px;
}

#uniHeader { margin-bottom: 20px; }

@media print {
  .print-footer {
    font-size: 10px;
    text-align: center;
    border-top: 1px solid black;
    padding: 6px 0;
    margin-top: 20px; /* space from content above */
  }

  .page-break {
    page-break-before: always;
  }

  .print-content {
    margin-bottom: 0; /* prevent extra bottom padding that causes overlap */
  }
}
</style>
        </head>
        <body>
  
  <div class="print-content">
    ${printContents}
  </div>
</body>
      </html>
    `);


  printWindow.document.close();

  // Wait for all images to load
  const images = printWindow.document.images;
  let loadedCount = 0;
  if (images.length === 0) {
    printWindow.focus();
    printWindow.print();
    return;
  }

  for (let img of images) {
    if (img.complete) {
      loadedCount++;
    } else {
      img.onload = img.onerror = () => {
        loadedCount++;
        if (loadedCount === images.length) {
          printWindow.focus();
          printWindow.print();
        }
      };
    }
  }

  // If all images already loaded
  if (loadedCount === images.length) {
    printWindow.focus();
    printWindow.print();
  }
}

// Helper to clone element and optionally replace inputs/selects with values
function cloneElementHTML(el, replaceInputs = false) {
  if (!el) return "";
  const clone = el.cloneNode(true);
  if (replaceInputs) {
    clone.querySelectorAll("input, select, textarea").forEach(input => {
      const span = document.createElement("span");
      span.textContent = input.value || "";
      input.parentNode.replaceChild(span, input);
    });
  }
  return clone.outerHTML;
}


function buildFooterCourseInfo() {
  const fields = [
   { label: "Department Name", id: "deptname" },
    { label: "Course Coordinator", id: "coordinator" },
    { label: "Department", id: "department" },
    { label: "Academic Year", id: "acyear" },
    { label: "UG / PG", id: "ugpg" },
    { label: "Semester", id: "semester" },
    { label: "Section", id: "section" },
    { label: "Course Code", id: "coursecode" },
    { label: "Course Name", id: "coursename" }


  ];

let html = `
  <table style="width:100%; margin-top:20px; border:1px solid black; font-size:10px;">
    <tr>${fields.map(f => `
      <td style="border:1px solid black; padding:6px; font-size:10px;">
        <strong>${f.label}:</strong> ${document.getElementById(f.id)?.value || ""}
      </td>`).join("")}
    </tr>
  </table>`;

  return html;
}


function matchSEEwithCIE() {
    const cieTable = document.querySelector("#uploadedTableCIE table");
    const seeTable = document.querySelector("#uploadedTableSEE table");

    if (!cieTable || !seeTable) {
        console.error("Tables not found!");
        return;
    }

    // Remove previous highlights
    [...cieTable.querySelectorAll("tbody tr")].forEach(row => row.classList.remove("sno-error", "missing-in-see"));
    [...seeTable.querySelectorAll("tbody tr")].forEach(row => row.classList.remove("to-delete", "sno-error"));

    // --- 1. Collect USNs ---
    const cieUSNs = new Set(
        [...cieTable.querySelectorAll("tbody tr")].map(row =>
            row.children[1]?.textContent.trim().toUpperCase()
        )
    );

    const seeUSNs = new Set(
        [...seeTable.querySelectorAll("tbody tr")].map(row =>
            row.children[1]?.textContent.trim().toUpperCase()
        )
    );

    // --- 2. Highlight SEE rows not in CIE ---
    let usnMismatchSEE = false;
    [...seeTable.querySelectorAll("tbody tr")].forEach(row => {
        const usn = row.children[1]?.textContent.trim().toUpperCase();
        if (!cieUSNs.has(usn)) {
            row.classList.add("to-delete");
            usnMismatchSEE = true;
        }
    });

    // --- 3. Highlight CIE rows not in SEE ---
    let usnMismatchCIE = false;
    [...cieTable.querySelectorAll("tbody tr")].forEach(row => {
        const usn = row.children[1]?.textContent.trim().toUpperCase();
        if (!seeUSNs.has(usn)) {
            row.classList.add("missing-in-see"); // highlight missing in SEE
            usnMismatchCIE = true;
        }
    });

    // --- 4. Check Serial Number continuity in CIE ---
    let prevSno = 0;
    let snoMismatchFound = false;
    [...cieTable.querySelectorAll("tbody tr")].forEach(row => {
        const sno = parseInt(row.children[0]?.textContent.trim());
        if (isNaN(sno) || sno !== prevSno + 1) {
            row.classList.add("sno-error");
            snoMismatchFound = true;
        }
        prevSno = isNaN(sno) ? prevSno : sno;
    });

    // --- 5. Check Serial Number continuity in SEE ---
    prevSno = 0;
    [...seeTable.querySelectorAll("tbody tr")].forEach(row => {
        const sno = parseInt(row.children[0]?.textContent.trim());
        if (isNaN(sno) || sno !== prevSno + 1) {
            row.classList.add("sno-error");
            snoMismatchFound = true;
        }
        prevSno = isNaN(sno) ? prevSno : sno;
    });

    // --- 6. Alert user ---
    let alertMsg = "";
    if (usnMismatchSEE) alertMsg += "USN(s) in SEE not found in CIE.\n";
    if (usnMismatchCIE) alertMsg += "USN(s) in CIE not found in SEE.\n";
    if (snoMismatchFound) alertMsg += "Serial number mismatch detected.\n";

    if (alertMsg) {
        alert(alertMsg + "Please correct the document(s) and re-upload.");
    }
}

// Run after tables appear
setTimeout(matchSEEwithCIE, 500);



function loadCOResults() {

    const saved = localStorage.getItem("CO_RESULTS");
    if (!saved) {
        alert("No SEE results found!");
        return;
    }

    document.getElementById("uploadedTableSEE").innerHTML = saved;

    insertNamesIntoSEE();     // 1. Insert names
    filterSEEByName();        // 2. Remove name-not-found rows
    reorderSEE_SLNO();        // ⭐ 3. Reorder serial numbers
    processFinalSEE();        // 4. Convert to AOA & compute summary
window.location.reload();

//window.history.replaceState({}, document.title, "obe.html");
}

function filterSEEByName() {
    const seeTable = document.querySelector("#uploadedTableSEE table");
    if (!seeTable) return;

    // Start from row 1 → skip header
    for (let i = seeTable.rows.length - 1; i >= 1; i--) {
        const name = seeTable.rows[i].cells[2].innerText.trim();
        if (name === "" || name === "NOT FOUND") {
            seeTable.deleteRow(i);
        }
    }

   // alert("Removed SEE rows without matching names.");
}
function processFinalSEE() {
    const seeTable = document.querySelector("#uploadedTableSEE table");
    if (!seeTable) return;

    let aoa = [];

    // Header row
    const headerRow = [];
    for (let c = 0; c < seeTable.rows[0].cells.length; c++) {
        headerRow.push(seeTable.rows[0].cells[c].innerText.trim());
    }
    aoa.push(headerRow);

    // Data rows
    for (let r = 1; r < seeTable.rows.length; r++) {
        const row = [];
        for (let c = 0; c < seeTable.rows[r].cells.length; c++) {
            row.push(seeTable.rows[r].cells[c].innerText.trim());
        }
        aoa.push(row);
    }

    // Now run your existing SEE processing
    processAOA("SEE", aoa);

    //alert("SEE summary computed and stored successfully!");
}
function reorderSEE_SLNO() {
    const seeTable = document.querySelector("#uploadedTableSEE table");
    if (!seeTable) return;

    // Start from row 1 (skip header)
    let slno = 1;
    for (let r = 1; r < seeTable.rows.length; r++) {
        seeTable.rows[r].cells[0].innerText = slno;
        slno++;
    }
}


function insertNamesIntoSEE() {

    const cieTable = document.querySelector("#uploadedTableCIE table");
    const seeTable = document.querySelector("#uploadedTableSEE table");

    if (!cieTable || !seeTable) {
        alert("Tables not found on page!");
        return;
    }

    // ----- BUILD USN → NAME MAPPING FROM CIE -----
    const usnToName = {};

    for (let i = 1; i < cieTable.rows.length; i++) {
        const row = cieTable.rows[i];

        const cieUSN = row.cells[1]?.innerText.trim();     // USN
        const cieName = row.cells[2]?.innerText.trim();    // NAME OF STUDENT

        if (cieUSN) {
            usnToName[cieUSN.toLowerCase()] = cieName;
        }
    }

    // ----- INSERT NAMES INTO SEE -----
    for (let i = 1; i < seeTable.rows.length; i++) {
        const row = seeTable.rows[i];

        const seeUSN = row.cells[1]?.innerText.trim().toLowerCase(); // student_usno
        const nameCell = row.cells[2];                               // NAME

        if (seeUSN && usnToName[seeUSN]) {
            nameCell.innerText = usnToName[seeUSN];
        } else {
            nameCell.innerText = "";  // you can put NOT FOUND if you want
        }
    }

    alert("Names inserted successfully!");
}





