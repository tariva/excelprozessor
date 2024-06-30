import { Workbook } from "exceljs";
import fs from "fs-extra";

export async function exportFilteredExcel(
  inputJsonData: any,
  mappingConfig: any,
  outputPath: string
) {
  try {
    // Prepare data for Excel export
    const dataForExcel: any[] = [];

    for (const itemKey in inputJsonData) {
      const item = inputJsonData[itemKey];
      const transformedItem: { [key: string]: any } = {};

      for (const configKey in mappingConfig) {
        if (item[mappingConfig[configKey]]) {
          transformedItem[configKey] = item[mappingConfig[configKey]];
        }
      }

      dataForExcel.push(transformedItem);
    }

    // Create an Excel file with the transformed data
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet("Sheet 1");

    // Add headers
    worksheet.columns = Object.keys(mappingConfig).map((key) => ({
      header: key,
      key: key,
    }));

    // Add rows
    dataForExcel.forEach((item) => {
      worksheet.addRow(item);
    });

    await workbook.xlsx.writeFile(outputPath);
    console.log(`Excel-Datei wurde erfolgreich exportiert nach: ${outputPath}`);
  } catch (error) {
    handleFileError(error, outputPath);
    throw error;
  }
}

export const getKeyForValue = (
  obj: { [key: string]: string },
  value: string
): string | undefined => {
  for (let key in obj) {
    if (obj[key] === value) {
      return key;
    }
  }
};

function handleFileError(err: any, filePath: string) {
  if (err.code === 'EBUSY') {
    console.error(`Fehler: Ressource ist beschäftigt oder gesperrt. Kann die Datei nicht öffnen: ${filePath}`);
  } else if (err.code === 'ENOENT') {
    console.error(`Fehler: Datei nicht gefunden: ${filePath}`);
  } else {
    console.error(`Fehler beim Zugriff auf die Datei ${filePath}:`, err);
  }
}
