import path from "path";
import fs from "fs-extra";
import ExcelJS from "exceljs";
import { Workbook } from "exceljs";
import {
  ensureTmpDirectory,
  copyToTmp,
  loadJSONConfig,
} from "./utils/fileHandler";
import { selectedFilesAndMoveToTmp } from "./utils/cli";

const CONFIG_DIR = path.join(process.cwd(), "config");

(async () => {
  await ensureTmpDirectory();

  const sourceExcelPath = "./src/excel";
  const workfiles = await selectedFilesAndMoveToTmp(sourceExcelPath);
  const jsonData: { [key: string]: any } = {};

  for (const filename of workfiles) {
    await mergeExcelToJSON(jsonData, filename);
  }

  if (Object.keys(jsonData).length) {
    // Save the merged data to a single JSON file, or separate if you need
    const mergedFileName = "merged_data.json";
    fs.writeFileSync(mergedFileName, JSON.stringify(jsonData, null, 2));
    console.log(`Merged data saved to ${mergedFileName}`);
  }

  const config = loadJSONConfig(path.join(CONFIG_DIR, "config.json"));
  console.log(config);
})();

async function mergeExcelToJSON(
  jsonData: any,
  filename: string
): Promise<void> {
  const workbook = new Workbook();
  await workbook.xlsx.readFile(filename);

  const worksheet = workbook.worksheets[0]; // Assuming data is in the first worksheet

  // Check if "Bezeichnung" column exists
  const hasBezeichnungColumn = worksheet.columns.some(
    (col) => col.values && col.values.includes("Bezeichnung")
  );

  if (!hasBezeichnungColumn) {
    console.warn(`File ${filename} does not have a "Bezeichnung" column.`);
    return;
  }

  // Use row.values for easy access, starting from row 2 to skip the header
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber !== 1) {
      // Skipping header
      const rowJSON: { [key: string]: any } = {};
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const header = worksheet.getRow(1).getCell(colNumber).text;
        // custom import logic

        rowJSON[header] = cell.text;
      });

      const key = rowJSON.Bezeichnung;

      if (jsonData[key]) {
        // Check for mismatched data or new columns to add
        for (const column in rowJSON) {
          if (
            jsonData[key][column] &&
            jsonData[key][column] !== rowJSON[column]
          ) {
            throw new Error(
              `Mismatch found in column "${column}" for key "${key}" while processing file "${filename}". Value are "${jsonData[key][column]}" and "${rowJSON[column]}".`
            );
          } else if (!jsonData[key][column]) {
            // New column encountered, add to the existing data
            jsonData[key][column] = rowJSON[column];
          }
        }
      } else {
        jsonData[key] = rowJSON;
      }
    }
  });
}
