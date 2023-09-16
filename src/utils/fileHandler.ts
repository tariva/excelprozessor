import fs from "fs-extra";
import path from "path";

const TMP_DIR = path.join(__dirname, "tmp");

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

export { ensureTmpDirectory, copyToTmp };
