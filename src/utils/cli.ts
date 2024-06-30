import fs from "fs-extra";
import path from "path";
import { copyToTmp } from "./fileHandler";
import checkbox from "@inquirer/checkbox";
import confirm from "@inquirer/confirm";
import select from "@inquirer/select";

const selectExcelFiles = async (directory: string): Promise<string[]> => {
  try {
    const files = await fs.readdir(directory);
    const excelFiles = files.filter(
      (file) =>
        file.endsWith(".xlsx") || file.endsWith(".xls") || file.endsWith(".csv")
    );

    if (excelFiles.length === 0) {
      console.log(`Keine Excel-Dateien im Verzeichnis gefunden: ${directory}`);
      return [];
    }

    const choices = excelFiles.map((file) => ({ name: file, value: file }));

    const selectedFiles = await checkbox({
      message: "Wählen Sie die zu verarbeitenden Excel-Dateien aus:",
      choices: choices,
    });

    return selectedFiles;
  } catch (err) {
    handleFileError(err, directory);
    throw err;
  }
};

const selectExcelFile = async (directory: string, promt = "Wählen Sie die zu verarbeitende Excel-Datei aus:", fileName?: string): Promise<string> => {
  try {
    const files = await fs.readdir(directory);
    const excelFiles = files.filter(
      (file) =>
        file.endsWith(".xlsx") || file.endsWith(".xls") || file.endsWith(".csv")
    );

    if (excelFiles.length === 0) {
      console.log(`Keine Excel-Dateien im Verzeichnis gefunden: ${directory}`);
      return "";
    }

    const choices = excelFiles.map((file) => ({ name: file, value: file }));
    // filter by filename if provided
    if (fileName) {
      const filteredChoices = choices.filter((choice) => choice.value === fileName);
      if (filteredChoices.length) {
        return filteredChoices[0].value;
      }
    }
    const selectedFiles = await select({
      message: promt,
      choices: choices,
    });

    return selectedFiles;
  } catch (err) {
    handleFileError(err, directory);
    throw err;
  }
};

const ask = async (message: string): Promise<boolean> => {
  return await confirm({ message });
};

const selectedFilesAndMoveToTmp = async (
  directory: string
): Promise<string[]> => {
  const selectedFiles = await selectExcelFiles(directory);
  const resultPath: string[] = [];
  for (const file of selectedFiles) {
    const filePath = path.join(directory, file);
    try {
      const tmpPath = await copyToTmp(filePath);

      if (tmpPath) {
        resultPath.push(tmpPath);
        console.log(`Datei ${file} wurde in das tmp-Verzeichnis kopiert: ${tmpPath}`);
        // Here you can further process the Excel file in tmpPath if needed
      }
    } catch (err) {
      handleFileError(err, filePath);
    }
  }
  return resultPath;
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

export { selectExcelFiles, selectedFilesAndMoveToTmp, ask, selectExcelFile };
