/* ========= CONFIG ========= */
const COL = {
  YEAR: 0,       // A
  SECTION: 2,    // C
  BRANCH: 3,     // D
  CODE: 4,       // E
  FACULTY: 5,    // F
  CIE_START: 6,  // G
  SEE_START: 11, // L
  PO_START: 17,  // R
  PSO_START: 29  // AD
};

const CIE_COUNT = 5;
const SEE_COUNT = 5;
const PO_COUNT  = 12;
const PSO_COUNT = 3;

/* ========= HELPERS ========= */
const isNumber = v => typeof v === "number" && !isNaN(v);

const average = arr => {
  const nums = arr.filter(isNumber);
  return nums.length
    ? (nums.reduce((a, b) => a + b, 0) / nums.length).toFixed(2)
    : "";
};

const uniqueList = (rows, idx) =>
  [...new Set(rows.map(r => r[idx]).filter(v => v))];

const verticalList = arr =>
  arr.map(v => `<div class="nowrap">${v}</div>`).join("");

/* ========= FILE LOAD ========= */
document.getElementById("file").addEventListener("change", e => {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = evt => parseExcel(evt.target.result);
  reader.readAsBinaryString(file);
});

/* ========= EXCEL PARSE ========= */
function parseExcel(binary) {
  const wb = XLSX.read(binary, { type: "binary" });
  const sheet = wb.Sheets[wb.SheetNames[0]];

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    range: "A9:AF2000"
  });

  generateStats(rows.slice(1)); // remove header
}

/* ========= GROUP BY YEAR + BRANCH ========= */
/* ========= GROUP BY YEAR + BRANCH + CODE ========= */
function generateStats(data) {
  const groups = {};

  data.forEach(row => {
    const year   = row[COL.YEAR];
    const branch = row[COL.BRANCH];
    const code   = row[COL.CODE];

    if (!year || !branch || !code) return;

    const key = `${year}__${branch}__${code}`;

    groups[key] ??= {
      year,
      branch,
      code,
      rows: []
    };

    groups[key].rows.push(row);
  });

  render(groups);
}

/* ========= CALCULATIONS ========= */
function calculate(group) {
  const rows = group.rows;

  const calcRange = (start, count) =>
    Array.from({ length: count }, (_, i) =>
      average(rows.map(r => r[start + i]))
    );

  return {
    cie: calcRange(COL.CIE_START, CIE_COUNT),
    see: calcRange(COL.SEE_START, SEE_COUNT),
    po:  calcRange(COL.PO_START,  PO_COUNT),
    pso: calcRange(COL.PSO_START, PSO_COUNT)
  };
}

/* ========= RENDER ========= */
function render(groups) {
  let html = `
    <table>
      <tr>
        <th>Year</th>
        <th>Branch</th>
        <th>Section(s)</th>
        <th>Code</th>
        <th>Faculty</th>
  `;

  for (let i = 1; i <= 5; i++) html += `<th>CIE CO${i}</th>`;
  for (let i = 1; i <= 5; i++) html += `<th>SEE CO${i}</th>`;
  for (let i = 1; i <= 12; i++) html += `<th>PO${i}</th>`;
  for (let i = 1; i <= 3; i++) html += `<th>PSO${i}</th>`;

  html += "</tr>";

  Object.values(groups).forEach(g => {
    const s = calculate(g);

    const sections = verticalList(uniqueList(g.rows, COL.SECTION));
    const code = g.code;

    const faculty  = verticalList(uniqueList(g.rows, COL.FACULTY));

    html += `
      <tr>
        <td>${g.year}</td>
        <td>${g.branch}</td>
        <td>${sections}</td>
       <td>${code}</td>
        <td>${faculty}</td>
    `;

    [...s.cie, ...s.see, ...s.po, ...s.pso]
      .forEach(v => html += `<td>${v}</td>`);

    html += "</tr>";
  });

  html += "</table>";
  document.getElementById("report").innerHTML = html;
}
