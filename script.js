let table = document.getElementsByClassName("sheet-body")[0],
  rows = document.getElementsByClassName("rows")[0],
  columns = document.getElementsByClassName("columns")[0];
tableExists = false;
const generateBtn = document.getElementById("generate");
const exportBtn = document.getElementById("export");
let items;

generateBtn.addEventListener("click", () => {
  if (rows.value === "" || columns.value === "") {
    Swal.fire({
      title: "It is Empty.",
      text: "Number of Rows and Columns are Empty.",
      icon: "error",
      showCancelButton: true,
      confirmButtonText: "Fill IT",
      cancelButtonText: "No",
    }).then((result) => {
      if (!result.isConfirmed) {
        Swal.fire({
          title: "You Should Fill it",
          text: "You Should Give me Number of Rows and Columns",
          icon: "warning",
          confirmButtonText: "OK",
        });
      }
    });
  } else {
    generateTable();
  }
});

exportBtn.addEventListener("click", () => {
  if (rows.value === "" || columns.value === "") {
    Swal.fire({
      title: "It is Empty.",
      text: "There is No Table to be Exported. You Should Generate Table First.",
      icon: "error",
      showCancelButton: true,
      confirmButtonText: "Generate one Now",
      cancelButtonText: "No",
    }).then((result) => {
      if (!result.isConfirmed) {
        Swal.fire({
          title: "You Should Generate it",
          text: "Generate Table Please",
          icon: "warning",
          confirmButtonText: "OK",
        });
      }
    });
  }
});

const generateTable = () => {
  let rowsNumber = parseInt(rows.value),
    columnsNumber = parseInt(columns.value);
  table.innerHTML = "";
  for (let i = 0; i < rowsNumber; i++) {
    var tableRow = "";
    for (let j = 0; j < columnsNumber; j++) {
      tableRow += `<td contenteditable class="table-item"></td>`;
    }
    table.innerHTML += tableRow;
  }
  if (rowsNumber > 0 && columnsNumber > 0) {
    tableExists = true;
  }
};

const ExportToExcel = (type, fn, dl) => {
  if (!tableExists) {
    return;
  }
  var elt = table;
  var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
  return dl
    ? XLSX.write(wb, { bookType: type, bookSST: true, type: "base64" })
    : XLSX.writeFile(wb, fn || "MyNewSheet." + (type || "xlsx"));
};
