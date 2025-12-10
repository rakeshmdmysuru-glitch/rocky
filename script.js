let excelData = [];
let studentDetails = [];

// ================= Upload main Excel =================
document.getElementById("excelFile").addEventListener("change", function(e){
  const reader = new FileReader();
  reader.onload = evt => {
    const workbook = XLSX.read(evt.target.result, { type:"array" });
    excelData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
    alert("Excel loaded: " + excelData.length + " students found.");
    populateSectionDropdown();
  };
  reader.readAsArrayBuffer(e.target.files[0]);
});

// ================= Upload details Excel =================
document.getElementById("detailsFile").addEventListener("change", function(e){
  const reader = new FileReader();
  reader.onload = evt => {
    const workbook = XLSX.read(evt.target.result, { type: "array" });
    studentDetails = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
    alert("Details Excel loaded!");
  };
  reader.readAsArrayBuffer(e.target.files[0]);
});

// ================= Populate SECTION Dropdown =================
function populateSectionDropdown() {
  const outputDiv = document.getElementById("output");
  let dropdown = document.getElementById("sectionFilter");

  if (!dropdown) {
    // Create dropdown above output
    const label = document.createElement("label");
    label.innerText = "Filter by SECTION:";
    label.setAttribute("for", "sectionFilter");
    label.classList.add("form-label", "fw-bold");

    dropdown = document.createElement("select");
    dropdown.id = "sectionFilter";
    dropdown.className = "form-select mb-3";
    dropdown.style.width = "200px";

    const allOption = document.createElement("option");
    allOption.value = "ALL";
    allOption.innerText = "ALL";
    dropdown.appendChild(allOption);

    outputDiv.parentNode.insertBefore(label, outputDiv);
    outputDiv.parentNode.insertBefore(dropdown, outputDiv);
  }

  // Clear old options except ALL
  Array.from(dropdown.options).forEach((opt, i) => { if(i>0) dropdown.removeChild(opt); });

  const sections = [...new Set(excelData.map(s => s.SECTION).filter(s => s))];
  sections.forEach(sec => {
    const option = document.createElement("option");
    option.value = sec;
    option.innerText = sec;
    dropdown.appendChild(option);
  });

  dropdown.addEventListener("change", filterBySection);
}

function filterBySection() {
  const selected = this.value;
  const table = document.querySelector("#output table");
  if (!table) return;

  const rows = Array.from(table.getElementsByTagName("tr")).slice(1); // exclude header
  let visibleCount = 1;

  rows.forEach(row => {
    const sectionCell = row.cells[3]; // SLNO=0, USN=1, NAME=2, SECTION=3
    if(!sectionCell) return;
    const section = sectionCell.innerText;
    
    if (selected === "ALL" || section === selected) {
      row.style.display = "";
      row.cells[0].innerText = visibleCount++; // renumber SLNO
    } else {
      row.style.display = "none";
    }
  });
}

// ================= Q Marks =================
function getQMarks(student,qNo){
  let total = 0;
  Object.keys(student).forEach(col=>{
    const match = col.match(/\d+/);
    if(match && parseInt(match[0])===qNo){
      const val = parseFloat(student[col]);
      if(!isNaN(val)) total += val;
    }
  });
  return total;
}

// ================= Compute COs =================
function computeCOs(){
  if(excelData.length===0){ alert("Please upload Excel first!"); return; }

  let coMapping = {};
  for(let q=1;q<=15;q++){
    const el = document.getElementById(`co_Q${q}`);
    if(el && el.value) coMapping[q] = parseInt(el.value);
  }

  const orPairs = [[6,7],[8,9],[10,11],[12,13],[14,15]];
  let results = [];

  excelData.forEach(student=>{
    let COscores = {};
    Object.values(coMapping).forEach(co=>{
      if(!COscores[`CO${co}`]) COscores[`CO${co}`]=0;
    });

    for(let q=1;q<=5;q++){
      if(coMapping[q]) COscores[`CO${coMapping[q]}`] += getQMarks(student,q);
    }

    orPairs.forEach(pair=>{
      const [q1,q2] = pair;
      if(!coMapping[q1]) return;
      const co = coMapping[q1];
      const m1 = getQMarks(student,q1);
      const m2 = getQMarks(student,q2);
      COscores[`CO${co}`] += Math.max(m1,m2);
    });

    results.push({student_usno: student.roll_num, NAME: student.NAME || "", SECTION: student.Section, ...COscores});
  });

  displayResult(results);
}

// ================= Display Results =================
function displayResult(data){
  if(data.length === 0){ 
    alert("No data to display"); 
    return; 
  }

  // Build table HTML
  let html = "<table class='table table-bordered table-hover table-sm mt-3'><thead><tr>";
  
  // Ensure SLNO is first, then student_usno, NAME, SECTION, then COs
  const keys = Object.keys(data[0]).filter(k => k !== "SLNO");
  html += "<th>SLNO</th>";
  keys.forEach(k => html += `<th>${k}</th>`);
  html += "</tr></thead><tbody>";

  data.forEach((row, i) => {
    html += "<tr>";
    html += `<td>${i + 1}</td>`; // SLNO
    keys.forEach(k => html += `<td>${row[k]}</td>`);
    html += "</tr>";
  });

  html += "</tbody></table>";
  document.getElementById("output").innerHTML = html;

  // After table is rendered, populate SECTION dropdown
  populateSectionDropdown();
}

// ================= Export to Excel =================
function exportToExcel() {
  const table = document.querySelector("#output table");
  if (!table) {
    alert("Please compute results before exporting!");
    return;
  }

  // Clone table so we don't modify original
  const tableClone = table.cloneNode(true);

  // Find SECTION column index
  const headerCells = tableClone.rows[0].cells;
  let sectionIndex = -1;
  for (let i = 0; i < headerCells.length; i++) {
    if (headerCells[i].innerText.trim().toUpperCase() === "SECTION") {
      sectionIndex = i;
      break;
    }
  }

  // Remove SECTION column from all rows
  if (sectionIndex !== -1) {
    for (let i = 0; i < tableClone.rows.length; i++) {
      tableClone.rows[i].deleteCell(sectionIndex);
    }
  }

  // Export
  const workbook = XLSX.utils.table_to_book(tableClone, { sheet: "CO Results" });
  XLSX.writeFile(workbook, "CO_Results.xlsx");
}

function populateSectionDropdown() {
  const dropdown = document.getElementById("sectionFilter");
  if (!dropdown) return;

  // Remove old options except ALL
  Array.from(dropdown.options).forEach((opt, i) => { if(i>0) dropdown.removeChild(opt); });

  // Get unique sections (case-insensitive)
  const sections = [...new Set(excelData.map(s => s.SECTION || s.Section || "").filter(s => s))];

  sections.forEach(sec => {
    const option = document.createElement("option");
    option.value = sec;
    option.innerText = sec;
    dropdown.appendChild(option);
  });

  // Attach listener only once
  dropdown.removeEventListener("change", filterBySection);
  dropdown.addEventListener("change", filterBySection);
}

function addNameColumn() { 
  const table = document.querySelector("#output table");
  if (!table) {
    alert("Compute CO results first!");
    return;
  }

  const headerRow = table.rows[0];
  const getText = (c) => c.innerText.trim().toLowerCase();

  // Find indexes (case-insensitive)
  const usnIndex = Array.from(headerRow.cells).findIndex(c => getText(c) === "student_usno" || getText(c) === "usn");
  let nameIndex = Array.from(headerRow.cells).findIndex(c => getText(c) === "name");
  const sectionIndex = Array.from(headerRow.cells).findIndex(c => getText(c) === "section");

  // ------------ ADD SLNO IF NOT PRESENT ------------
  if (!Array.from(headerRow.cells).some(c => getText(c) === "slno")) {
    headerRow.insertCell(0).outerHTML = "<th>SLNO</th>";
    for (let i = 1; i < table.rows.length; i++) {
      table.rows[i].insertCell(0).innerText = i;
    }
  }

  // Update indexes because SLNO insertion shifts columns
  nameIndex = Array.from(headerRow.cells).findIndex(c => getText(c) === "name");

  // ------------ ADD/UPDATE NAME COLUMN ------------
  if (usnIndex !== -1 && nameIndex === -1) {
    headerRow.insertCell(usnIndex + 1).outerHTML = "<th>NAME</th>";
    nameIndex = usnIndex + 1;

    for (let i = 1; i < table.rows.length; i++) {
      const usn = table.rows[i].cells[usnIndex].innerText.trim();
      const match = studentDetails.find(s => (s.USN || "").trim().toLowerCase() === usn.toLowerCase());
      const nameValue = match ? (match.NAME || match["NAME OF STUDENT"] || "") : "";
      table.rows[i].insertCell(nameIndex).innerText = nameValue;
    }
  } 
  else if (nameIndex !== -1) {
    for (let i = 1; i < table.rows.length; i++) {
      const usn = table.rows[i].cells[usnIndex].innerText.trim();
      const match = studentDetails.find(s => (s.USN || "").trim().toLowerCase() === usn.toLowerCase());
      table.rows[i].cells[nameIndex].innerText = match ? (match.NAME || match["NAME OF STUDENT"] || "") : "";
    }
  }

  // ------------ MOVE SECTION NEXT TO NAME ------------
  if (sectionIndex > nameIndex && sectionIndex !== -1) {
    for (let i = 0; i < table.rows.length; i++) {
      const row = table.rows[i];
      const sectionCell = row.cells[sectionIndex];
      row.insertBefore(sectionCell, row.cells[nameIndex + 1]);
    }
  }

  // ------------ SORT ROWS: NAME first, blanks last ------------
  let rows = Array.from(table.rows).slice(1); // skip header

  rows.sort((a, b) => {
    const nameA = a.cells[nameIndex].innerText.trim();
    const nameB = b.cells[nameIndex].innerText.trim();

    if (nameA && !nameB) return -1; // name first
    if (!nameA && nameB) return 1;  // blanks last

    return nameA.localeCompare(nameB);
  });

  // Reattach sorted rows
  const tbody = table.tBodies[0];
  rows.forEach((r, i) => {
    tbody.appendChild(r);
    r.cells[0].innerText = i + 1; // refresh SLNO
  });

  alert("NAME-filled rows arranged first successfully!");
}

