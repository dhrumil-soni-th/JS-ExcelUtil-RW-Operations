const excelJs = require("exceljs");

async function writeExcelTest(searchText, replaceText, filePath) {
  const workbook = new excelJs.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.getWorksheet("Sheet1");

  const output = await readExcelFile(worksheet, searchText);

  const cell = worksheet.getCell(output.row, output.col);
  cell.value = replaceText;
  await workbook.xlsx.writeFile(filePath);
}

async function readExcelFile(worksheet, searchText) {
  let output = { row: -1, col: -1 };
  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      console.log(cell.value);
      if (cell.value === searchText) {
        output.row = rowNumber;
        output.col = colNumber;
      }
    });
  });
  return output;
}

writeExcelTest(
  "Banana",
  "Republic",
  "/Users/dhrumilsoni/Downloads/SampleExcelTest.xlsx"
);
