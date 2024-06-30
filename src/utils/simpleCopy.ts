import { Row, Workbook } from "exceljs";
import { getKeyForValue } from "./utils";
import fs from "fs-extra";

async function convertExcelToJson(filename: string, config: any): Promise<any> {
    const workbook = new Workbook();

    try {
        await workbook.xlsx.readFile(filename);
    } catch (err) {
        handleFileError(err, filename);
        throw err;
    }

    const worksheet = workbook.getWorksheet(config.destWorksheetName);

    if (!worksheet) {
        throw new Error(`Arbeitsblatt mit dem Namen ${config.destWorksheetName} wurde in ${filename} nicht gefunden.`);
    }

    const jsonData: { [key: string]: any } = {};

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber >= config.destStartDataRow) {
            const rowJSON: { [key: string]: any } = {};
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                const header = worksheet.getRow(config.destMappingRow).getCell(colNumber).text;
                rowJSON[header] = cell.text;
            });

            const key = rowJSON[config.keyColumn];
            if (!key) {
                console.warn(`Schlüsselspaltenwert fehlt in Zeile ${rowNumber}`);
            }
            jsonData[key] = rowJSON;
        }
    });

    return jsonData;
}

async function mergeExcelFiles(
    sourceFile: string,
    destinationFile: string,
    config: any
): Promise<void> {
    const sourceWorkbook = new Workbook();
    const destWorkbook = new Workbook();

    try {
        await sourceWorkbook.xlsx.readFile(sourceFile);
    } catch (err) {
        handleFileError(err, sourceFile);
        throw err;
    }

    try {
        await destWorkbook.xlsx.readFile(destinationFile);
    } catch (err) {
        handleFileError(err, destinationFile);
        throw err;
    }

    const sourceWorksheet = sourceWorkbook.getWorksheet(config.sourceWorksheetName);
    const destWorksheet = destWorkbook.getWorksheet(config.destWorksheetName);

    if (!sourceWorksheet) {
        throw new Error(`Arbeitsblatt mit dem Namen ${config.sourceWorksheetName} wurde in ${sourceFile} nicht gefunden.`);
    }
    if (!destWorksheet) {
        throw new Error(`Arbeitsblatt mit dem Namen ${config.destWorksheetName} wurde in ${destinationFile} nicht gefunden.`);
    }

    const colMap = mapColumns(sourceWorksheet, destWorksheet, config);

    const keyColIndexSource = colMap.source[config.keyColumn];
    const keyColIndexDest = colMap.dest[config.keyColumn];

    if (keyColIndexSource === undefined) {
        throw new Error(`Schlüsselspalte ${config.keyColumn} wurde in der Quelldatei ${sourceFile} nicht gefunden.`);
    }
    if (keyColIndexDest === undefined) {
        throw new Error(`Schlüsselspalte ${config.keyColumn} wurde in der Zieldatei ${destinationFile} nicht gefunden.`);
    }

    const keyMap: { [key: string]: { sourceRow: Row | undefined; destRow: Row } } = {};

    destWorksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber >= config.destStartDataRow) {
            const key = row.getCell(keyColIndexDest).text;
            if (key) {
                keyMap[key] = { sourceRow: undefined, destRow: row };
            } else {
                console.warn(`Fehlender Schlüsselwert in der Zielzeile ${rowNumber}`);
            }
        }
    });

    sourceWorksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber >= config.sourceStartDataRow) {
            const key = row.getCell(keyColIndexSource).text;
            if (key && keyMap[key]) {
                keyMap[key].sourceRow = row;
            } else if (!key) {
                console.warn(`Fehlender Schlüsselwert in der Quellzeile ${rowNumber}`);
            }
        }
    });

    for (const key in keyMap) {
        const mapping = keyMap[key];
        if (mapping.sourceRow && mapping.destRow) {
            for (const [sourceCol, destCol] of Object.entries(config.mapping)) {
                const sourceColIndex = colMap.source[sourceCol];
                const destColIndex = colMap.dest[destCol as any];

                if (sourceColIndex !== undefined && destColIndex !== undefined) {
                    const sourceCellValue = mapping.sourceRow.getCell(sourceColIndex).text;
                    mapping.destRow.getCell(destColIndex).value = sourceCellValue;
                } else {
                    if (sourceColIndex === undefined) {
                        console.warn(`Spalte ${sourceCol} wurde in der Quelldatei nicht gefunden.`);
                    }
                    if (destColIndex === undefined) {
                        console.warn(`Spalte ${destCol} wurde in der Zieldatei nicht gefunden.`);
                    }
                }
            }
        } else if (!mapping.sourceRow) {
            console.warn(`Keine passende Quellzeile für Schlüssel ${key}`);
        } else if (!mapping.destRow) {
            console.warn(`Keine passende Zielzeile für Schlüssel ${key}`);
        }
    }

    try {
        await destWorkbook.xlsx.writeFile(destinationFile);
    } catch (err) {
        handleFileError(err, destinationFile);
        throw err;
    }
    console.log(`Daten wurden zusammengeführt und in ${destinationFile} gespeichert.`);
}

function mapColumns(sourceWorksheet: any, destWorksheet: any, config: any): { source: { [key: string]: number }, dest: { [key: string]: number } } {
    const sourceColMap: { [key: string]: number } = {};
    const destColMap: { [key: string]: number } = {};

    // Map source columns based on config
    sourceWorksheet.getRow(config.sourceMappingRow).eachCell((cell, colNumber) => {
        const colName = cell.text;
        if (config.mapping[colName]) {
            sourceColMap[colName] = colNumber;
        }
    });

    // Map destination columns based on config
    destWorksheet.getRow(config.destMappingRow).eachCell((cell, colNumber) => {
        const colName = cell.text;
        if (getKeyForValue(config.mapping, colName)) {
            destColMap[colName] = colNumber;
        }
    });

    return { source: sourceColMap, dest: destColMap };
}

function handleFileError(err: any, filePath: string) {
    if (err.code === 'EBUSY') {
        console.error(`Fehler: Ressource ist beschäftigt oder gesperrt. Kann die Datei nicht öffnen: ${filePath}`);
    } else if (err.code === 'ENOENT') {
        console.error(`Fehler: Datei nicht gefunden: ${filePath}`);
    } else {
        console.error(`Fehler beim Zugriff auf die Datei ${filePath}:`, err);
    }
}

export { convertExcelToJson, mergeExcelFiles };
