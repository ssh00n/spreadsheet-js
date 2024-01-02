document.addEventListener("DOMContentLoaded", function () {
  const exportBtn = document.getElementById("exportBtn");
  exportBtn.addEventListener("click", exportSpreadsheet);
});

function generateSpreadsheet() {
  const rows = document.getElementById("rows").value;
  const columns = document.getElementById("columns").value;
  const spreadsheet = document.getElementById("spreadsheet");
  spreadsheet.innerHTML = "";
  spreadsheet.addEventListener("click", (e) => {
    if (e.target.tagName === "TD") {
      highlightHeaders(e.target);
    }
  });

  for (let i = 0; i <= rows; i++) {
    let row = spreadsheet.insertRow();
    for (let j = 0; j <= columns; j++) {
      if (i === 0 && j === 0) {
        row.insertCell().outerHTML = "<th></th>";
      } else if (i === 0) {
        row.insertCell().outerHTML = `<th>${columnName(j)}</th>`;
      } else if (j === 0) {
        row.insertCell().outerHTML = `<th>${i}</th>`;
      } else {
        let cell = row.insertCell();
        cell.contentEditable = true;
      }
    }
  }
}

function highlightHeaders(cell) {
  removeHighlights();

  let colIndex = cell.cellIndex;
  let rowIndex = cell.parentElement.rowIndex;

  document
    .querySelector(`#spreadsheet th:nth-child(${colIndex + 1})`)
    .classList.add("highlight");
  document
    .querySelector(`#spreadsheet tr:nth-child(${rowIndex + 1}) th`)
    .classList.add("highlight");
}

function removeHighlights() {
  document.querySelectorAll("#spreadsheet .highlight").forEach((cell) => {
    cell.classList.remove("highlight");
  });
}

function columnName(index) {
  let name = "";
  while (index > 0) {
    index--;
    name = String.fromCharCode("A".charCodeAt(0) + (index % 26)) + name;
    index = Math.floor(index / 26);
  }
  return name;
}

function exportSpreadsheet() {
  let table = document.getElementById("spreadsheet");
  let wb = XLSX.utils.book_new();

  let aoa = [];
  for (let i = 1; i < table.rows.length; i++) {
    let row = [];
    for (let j = 1; j < table.rows[i].cells.length; j++) {
      row.push(table.rows[i].cells[j].innerText);
    }
    aoa.push(row);
  }

  let ws = XLSX.utils.aoa_to_sheet(aoa);
  XLSX.utils.book_append_sheet(wb, ws, "Sheet 1");

  let fileName = document.getElementById("fileName").value.trim();
  if (!fileName) {
    fileName = "spreadsheet"; // default
  }

  fileName += ".xlsx";

  XLSX.writeFile(wb, fileName);
}
