import fs from "fs-extra";
import ExcelJS from "exceljs";
import { Workbook } from "exceljs";
import { ensureTmpDirectory, copyToTmp } from "./utils/fileHandler";
import { selectedFilesAndMoveToTmp } from "./utils/cli";
(async () => {
  await ensureTmpDirectory();

  const sourceExcelPath = "./src/excel";
  const workfiles = await selectedFilesAndMoveToTmp(sourceExcelPath);
  const jsonData: { [key: string]: any } = {};
  for (const filename of workfiles) {
    await convertExcelToJSON(jsonData, filename);

    if (jsonData) {
      fs.writeFileSync(`${filename}.json`, JSON.stringify(jsonData, null, 2));
      console.log(`Converted ${filename} to ${filename}.json`);
    }
  }
  // For demonstration, you can call manipulateExcelData with the path to your copied Excel file in tmp
  // await manipulateExcelData(tmpExcelPath);
})();
async function convertExcelToJSON(jsonData: any, filename: string): Promise {
  const workbook = new Workbook();
  await workbook.xlsx.readFile(filename);

  const worksheet = workbook.worksheets[0]; // Assuming data is in the first worksheet

  // Check if "Bezeichnung" column exists
  const hasBezeichnungColumn = worksheet.columns.some(
    (col) => col.values && col.values.includes("Bezeichnung")
  );

  if (!hasBezeichnungColumn) {
    console.warn(`File ${filename} does not have a "Bezeichnung" column.`);
    return null;
  }

  // Use row.values for easy access, starting from row 2 to skip the header
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber !== 1) {
      // Skipping header
      const rowJSON: { [key: string]: any } = {};
      row.eachCell((cell, colNumber) => {
        const header = worksheet.getRow(1).getCell(colNumber).text;
        rowJSON[header] = cell.text;
      });
      jsonData[rowJSON.Bezeichnung] = rowJSON;
    }
  });
}
