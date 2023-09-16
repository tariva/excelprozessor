import fs from "fs-extra";
import ExcelJS from "exceljs";
import { ensureTmpDirectory, copyToTmp } from "./utils/fileHandler";
import { selectedFilesAndMoveToTmp } from "./utils/cli";
(async () => {
  await ensureTmpDirectory();

  const sourceExcelPath = "./src/excel";
  await selectedFilesAndMoveToTmp(sourceExcelPath);

  // For demonstration, you can call manipulateExcelData with the path to your copied Excel file in tmp
  // await manipulateExcelData(tmpExcelPath);
})();
