import path from "path";
import fs from "fs-extra";
import { Workbook } from "exceljs";
import { ensureTmpDirectory, loadJSONConfig } from "./utils/fileHandler";
import { selectedFilesAndMoveToTmp } from "./utils/cli";
import { exportFilteredExcel } from "./utils/utils";

const CONFIG_DIR = path.join(process.cwd(), "config");

(async () => {
  await ensureTmpDirectory();

  const sourceExcelPath = "./src/excel";
  const workfiles = await selectedFilesAndMoveToTmp(sourceExcelPath);
  const jsonData: { [key: string]: any } = {};
  const config = loadJSONConfig(path.join(CONFIG_DIR, "config.json"));
  for (const filename of workfiles) {
    await mergeExcelToJSON(jsonData, filename, config.adjustmentkeys);
  }

  if (Object.keys(jsonData).length) {
    // Save the merged data to a single JSON file, or separate if you need
    const mergedFileName = "merged_data.json";
    fs.writeFileSync(mergedFileName, JSON.stringify(jsonData, null, 2));
    console.log(`Merged data saved to ${mergedFileName}`);
  }

  exportFilteredExcel(jsonData, config.mapping, "output.xlsx");
})();

async function mergeExcelToJSON(
  jsonData: any,
  filename: string,
  adjustmentkeys: any
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
        if (adjustmentkeys.includes(header)) {
          const value = cell.text;
          if (value === "" || value === null || parseFloat(value) > 999) {
            rowJSON[header] = "free";
          } else {
            rowJSON[header] = cell.text;
          }
        } else {
          rowJSON[header] = cell.text;
        }
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
