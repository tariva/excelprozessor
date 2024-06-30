import fs from "fs-extra";
import path from "path";
import json5 from "json5";
const TMP_DIR = path.join(process.cwd(), "tmp");

const ensureTmpDirectory = async (): Promise<void> => {
  try {
    await fs.ensureDir(TMP_DIR);
  } catch (error) {
    console.error("Error ensuring tmp directory:", error);
  }
};

const copyToTmp = async (sourcePath: string): Promise<string | null> => {
  try {
    const destinationPath = path.join(TMP_DIR, path.basename(sourcePath));
    await fs.copy(sourcePath, destinationPath);
    return destinationPath;
  } catch (error) {
    console.error("Error copying file to tmp:", error);
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
  // Ensure the file exists
  if (!fs.existsSync(filePath)) {
    throw new Error(`Config file not found at path: ${filePath}`);
  }

  // Read and parse the file
  const rawContent = fs.readFileSync(filePath, "utf-8");
  return json5.parse(rawContent);
};

// Example usage (for testing)
// const configPath = path.join(__dirname, 'path_to_config.json');
// const config = loadJSONConfig(configPath);
// console.log(config);

export { ensureTmpDirectory, copyToTmp, loadJSONConfig };
