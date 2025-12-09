let excelData = [];

document.getElementById("excelFile").addEventListener("change", function(e){
  const reader = new FileReader();
  reader.onload = evt => {
    const workbook = XLSX.read(evt.target.result, { type:"array" });
    excelData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
    alert("Excel loaded: " + excelData.length + " students found.");
  };
  reader.readAsArrayBuffer(e.target.files[0]);
});

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

    results.push({student_usno: student.student_usno, ...COscores});
  });

  displayResult(results);
}

function displayResult(data){
  if(data.length===0){ alert("No data to display"); return; }
  let html = "<h3>Computed CO Results</h3><table border='1'><tr>";
  Object.keys(data[0]).forEach(h=>html+=`<th>${h}</th>`);
  html+="</tr>";
  data.forEach(row=>{
    html+="<tr>";
    Object.values(row).forEach(v=>html+=`<td>${v}</td>`);
    html+="</tr>";
  });
  html+="</table>";
  document.getElementById("output").innerHTML = html;
}
function exportToExcel() {
  const table = document.querySelector("#output table");
  if (!table) {
    alert("Please compute results before exporting!");
    return;
  }

  const workbook = XLSX.utils.table_to_book(table, { sheet: "CO Results" });
  XLSX.writeFile(workbook, "CO_Results.xlsx");
}

let studentDetails = [];

document.getElementById("detailsFile").addEventListener("change", function(e){
  const reader = new FileReader();
  reader.onload = evt => {
    const workbook = XLSX.read(evt.target.result, { type: "array" });
    studentDetails = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
    alert("Details Excel loaded!");
  };
  reader.readAsArrayBuffer(e.target.files[0]);
});
function addNameColumn(){
  const table = document.querySelector("#output table");
  if (!table) {
    alert("Compute CO results first!");
    return;
  }

  // ================= Add SLNO column at first position =================
  const headerRow = table.rows[0];
  headerRow.insertCell(0).outerHTML = "<th>SLNO</th>";

  for (let i = 1; i < table.rows.length; i++) {
    table.rows[i].insertCell(0).innerText = i;  // Auto numbering
  }

  // ================= Insert NAME after USN column =================
  const updatedHeaderRow = table.rows[0];
  // Recalculate USN column index AFTER adding SLNO
  const usnIndex = Array.from(updatedHeaderRow.cells)
                        .findIndex(c => c.innerText === "student_usno");

  updatedHeaderRow.insertCell(usnIndex + 1).outerHTML = "<th>NAME</th>";

  for (let i = 1; i < table.rows.length; i++) {
    const usn = table.rows[i].cells[usnIndex].innerText;  // Correct position
    const match = studentDetails.find(s => s.USN === usn);
    const nameValue = match ? match["NAME OF THE STUDENT"] : "";
    table.rows[i].insertCell(usnIndex + 1).innerText = nameValue;
  }

  alert("SLNO & Names added successfully!");
}
