import path from "path";
import fs from "fs-extra";
import { Workbook, FillPattern } from "exceljs";
import { ensureTmpDirectory, loadJSONConfig } from "./utils/fileHandler";
import { selectedFilesAndMoveToTmp, ask, selectExcelFile } from "./utils/cli";
import { exportFilteredExcel, getKeyForValue } from "./utils/utils";
const CONFIG_DIR = path.join(process.cwd(), "config");
const Excel_DIR = path.join(process.cwd(), "./");
(async () => {
  await ensureTmpDirectory();

  const workfiles = await selectedFilesAndMoveToTmp(Excel_DIR);
  const jsonData: { [key: string]: any } = {};
  const config = loadJSONConfig(path.join(CONFIG_DIR, "config.json"));
  for (const filename of workfiles) {
    await mergeExcelToJSON(jsonData, filename, config);
  }

  if (Object.keys(jsonData).length) {
    // Save the merged data to a single JSON file, or separate if you need
    const mergedFileName = "merged_data.json";
    fs.writeFileSync(mergedFileName, JSON.stringify(jsonData, null, 2));
    console.log(`Merged data saved to ${mergedFileName}`);
  }
  const exportExcel = await ask("Excel export?");
  if (exportExcel) {
    await exportFilteredExcel(jsonData, config.mapping, "output.xlsx");
  }
  const mergedWithTemplate = await ask("mit Template mergen?");
  if (mergedWithTemplate) {
    const template = await selectExcelFile(Excel_DIR);
    await writeJsonToExcel(jsonData, template, config);
  }

  await ask("Press any key to exit...");
})();

async function mergeExcelToJSON(
  jsonData: any,
  filename: string,
  config: any
): Promise<void> {
  const workbook = new Workbook();
  await workbook.xlsx.readFile(filename);

  const worksheet = workbook.worksheets[0]; // Assuming data is in the first worksheet

  // Check if "Bezeichnung" column exists
  const hasBezeichnungColumn = worksheet.columns.some(
    (col) => col.values && col.values.includes(config.keyColumn)
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
        if (config.adjustmentkeys.includes(header)) {
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

async function writeJsonToExcel(
  inputJsonData: any,
  excelPath: string,
  config: any
) {
  // Open the target Excel
  const workbook = new Workbook();
  await workbook.xlsx.readFile(excelPath);
  const worksheet = workbook.getWorksheet(config.worksheetName);

  // Find the column with the content "Bezeichnung"
  let bezeichnungColIndex = -1;
  for (let col of worksheet.columns) {
    if (col.values && getKeyForValue(config.mapping, config.keyColumn)) {
      bezeichnungColIndex = col.number as number; // Get the column number
      break;
    }
  }

  if (bezeichnungColIndex === -1) {
    console.error("Column with content 'Bezeichnung' not found.");
    return;
  }

  const yellowFill: FillPattern = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFFF00" }, // Yellow color
  };

  let rowStartIndex = config.startindexTemplate; // Start from the 5th row
  for (const key in inputJsonData) {
    const data = inputJsonData[key];

    if (data.Bezeichnung) {
      const row = worksheet.getRow(rowStartIndex);

      // Check for changes
      for (const configKey in config.mapping) {
        const colName = config.mapping[configKey];
        const colIndex = worksheet.columns.findIndex(
          (col) => col.values && col.values.includes(colName)
        );
        if (colIndex !== -1) {
          const cell = row.getCell(colIndex + 1); // +1 because columns are 1-based in ExcelJS
          if (cell.value !== data[colName]) {
            console.log(
              `Value changed in ${data.Bezeichnung} - ${colName}: ${cell.value} => ${data[colName]}`
            );
            cell.fill = yellowFill; // Highlight the cell with yellow
          }
          cell.value = data[colName]; // Update the value
        }
      }

      rowStartIndex++; // Move to next row
    }
  }

  await workbook.xlsx.writeFile(excelPath); // Save changes to Excel file
}
