import fs from "fs-extra";
import path from "path";
import { copyToTmp } from "./fileHandler";
import checkbox from "@inquirer/checkbox";

const selectExcelFiles = async (directory: string): Promise<string[]> => {
  const files = await fs.readdir(directory);
  const excelFiles = files.filter(
    (file) =>
      file.endsWith(".xlsx") || file.endsWith(".xls") || file.endsWith(".csv")
  );

  if (excelFiles.length === 0) {
    console.log("No Excel files found in directory:", directory);
    return [];
  }

  const choices = excelFiles.map((file) => ({ name: file, value: file }));

  const selectedFiles = await checkbox({
    message: "Select Excel files to process:",
    choices: choices,
  });

  return selectedFiles;
};

const selectedFilesAndMoveToTmp = async (
  directory: string
): Promise<string[]> => {
  const selectedFiles = await selectExcelFiles(directory);
  const resultPath = [];
  for (const file of selectedFiles) {
    const filePath = path.join(directory, file);
    const tmpPath = await copyToTmp(filePath);

    if (tmpPath) {
      resultPath.push(tmpPath);
      console.log(`File ${file} copied to tmp directory at path: ${tmpPath}`);
      // Here you can further process the Excel file in tmpPath if needed
    }
  }
  return resultPath;
};

export { selectExcelFiles, selectedFilesAndMoveToTmp };
