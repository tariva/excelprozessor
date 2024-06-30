import { Row, Workbook } from "exceljs";
import { getKeyForValue } from "./utils";

async function convertExcelToJson(filename: string, config: any): Promise<any> {
    const workbook = new Workbook();
    await workbook.xlsx.readFile(filename);
    const worksheet = workbook.getWorksheet(config.worksheetName);

    const jsonData: { [key: string]: any } = {};

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber >= config.destStartDataRow) {
            const rowJSON: { [key: string]: any } = {};
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                const header = worksheet.getRow(config.destStartDataRow - 1).getCell(colNumber).text;
                rowJSON[header] = cell.text;
            });

            const key = rowJSON[config.keyColumn];
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

    await sourceWorkbook.xlsx.readFile(sourceFile);
    await destWorkbook.xlsx.readFile(destinationFile);

    const sourceWorksheet = sourceWorkbook.getWorksheet(config.sourceWorksheetName);
    const destWorksheet = destWorkbook.getWorksheet(config.destWorksheetName);

    const colMap = mapColumns(sourceWorksheet, destWorksheet, config);

    const keyColIndexSource = colMap.source[config.keyColumn];
    const keyColIndexDest = colMap.dest[config.keyColumn];

    if (keyColIndexSource === undefined || keyColIndexDest === undefined) {
        console.error("Key column not found in one of the files.");
        return;
    }

    const keyMap: { [key: string]: { sourceRow: Row | undefined; destRow: Row } } = {};

    destWorksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber >= config.destStartDataRow) {
            const key = row.getCell(keyColIndexDest).text;
            if (key) {
                keyMap[key] = { sourceRow: undefined, destRow: row };
            }
        }
    });

    sourceWorksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber >= config.sourceStartDataRow) {
            const key = row.getCell(keyColIndexSource).text;
            if (key && keyMap[key]) {
                keyMap[key].sourceRow = row;
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
                }
            }
        }
    }

    await destWorkbook.xlsx.writeFile(destinationFile);
    console.log(`Data merged and saved to ${destinationFile}`);
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
export { convertExcelToJson, mergeExcelFiles }