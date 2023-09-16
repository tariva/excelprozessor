import { Workbook } from "exceljs";

export async function exportFilteredExcel(
  inputJsonData: any,
  mappingConfig: any,
  outputPath: string
) {
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
}
// Example usage:
// const jsonData = { ... }; // This is your loaded and processed JSON data
// const configPath = 'path_to_your_config.json';
// const outputPath = 'path_for_exported_excel.xlsx';
// await exportFilteredExcel(jsonData, configPath, outputPath);
