const excelJs = require("exceljs");

async function excelTest() {
  let output = { row: -1, col: -1 };
  const workbook = new excelJs.Workbook();
  await workbook.xlsx.readFile(
    "/Users/dhrumilsoni/Downloads/SampleExcelTest.xlsx"
  );
  const worksheet = workbook.getWorksheet("Sheet1");
  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      console.log(cell.value);
      if (cell.value === "Apple") {
        output.row = rowNumber;
        output.col = colNumber;
      }
    });
  });
  const cell = worksheet.getCell(output.row, output.col);
  cell.value = "iPhone";
  await workbook.xlsx.writeFile(
    "/Users/dhrumilsoni/Downloads/SampleExcelTest.xlsx"
  );
  console.log("Updated cell value to iPhone");
  const cell2 = worksheet.getCell(3, 2);
  console.log("Updated cell value: " + cell2.value);
}

excelTest()
  .then(() => {
    console.log("Excel file read successfully");
    console.log("Excel file updated successfully");
  })
  .catch((error) => {
    console.error("Error reading Excel file:", error);
  });
