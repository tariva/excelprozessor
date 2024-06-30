import fs from "fs-extra";
import path from "path";
import json5 from "json5";

const TMP_DIR = path.join(process.cwd(), "tmp");

const ensureTmpDirectory = async (): Promise<void> => {
  try {
    await fs.ensureDir(TMP_DIR);
  } catch (error) {
    console.error("Fehler beim Erstellen des tmp-Verzeichnisses:", error);
  }
};

const copyToTmp = async (sourcePath: string): Promise<string | null> => {
  try {
    const destinationPath = path.join(TMP_DIR, path.basename(sourcePath));
    await fs.copy(sourcePath, destinationPath);
    return destinationPath;
  } catch (error) {
    handleFileError(error, sourcePath);
    return null;
  }
};

/**
 * Load a JSON configuration file.
 *
 * @param {string} filePath - The path to the JSON configuration file.
 * @returns {any} - The parsed JSON content of the configuration file.
 */
const loadJSONConfig = (filePath: string): any => {
  try {
    // Ensure the file exists
    if (!fs.existsSync(filePath)) {
      throw new Error(`Konfigurationsdatei nicht gefunden: ${filePath}`);
    }

    // Read and parse the file
    const rawContent = fs.readFileSync(filePath, "utf-8");
    return json5.parse(rawContent);
  } catch (error) {
    console.error("Fehler beim Laden der Konfigurationsdatei:", error);
    throw error;
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

// Example usage (for testing)
// const configPath = path.join(__dirname, 'path_to_config.json');
// const config = loadJSONConfig(configPath);
// console.log(config);

export { ensureTmpDirectory, copyToTmp, loadJSONConfig };
