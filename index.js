const XLSX = require('xlsx');
const path = require('path');
const ExcelJS = require('exceljs');
let dir = "D:\\Shourov\\office\\documents\\Prime-Asia\\exam";
let fileName = "Pau-Exam-Data-after142.xlsx";
let tabName = "main";
let workbook = null;
const startCellNo = "AQ5";
const endCellNo = "BN100";


async function readLargeFile() {

    let [startRow, startCol] = getRowColNum(startCellNo);
    let [endRow, endCol] = getRowColNum(endCellNo);

    // Use the stream reader
    const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(path.join(dir, fileName));

    for await (const worksheetReader of workbookReader) {
        if (worksheetReader.name !== tabName) continue;
        for await (const row of worksheetReader) {
            // 'row.values' gives you the data for the current row
            // You can filter your specific range here manually
            if (row.number >= startRow && row.number <= endRow) {
                console.log(`Row ${row.number}:`, row.values[startCol]);
                let rowData = "";
                for (let i = startCol; i <= endCol; i++) {
                    let cell = row.values[i]
                    try {
                        let temp = (cell);
                        if (temp['result']) rowData += temp.result;
                        else rowData += cell
                    } catch (error) {
                        rowData += cell ?? " ";

                    }

                    if (i < endCol) rowData += ","
                }
                console.log(rowData)
                break;
            }

            // Stop early if you only need a specific range at the top
            if (row.number > 5) break;
        }

    }
}


function columnToNumber(columnName) {
    let number = 0;
    const name = columnName.toUpperCase();

    for (let i = 0; i < name.length; i++) {
        // Get the character code (A=65, B=66, etc.)
        // Subtract 64 so that A=1, B=2...
        const charValue = name.charCodeAt(i) - 64;

        // Shift the previous total by power of 26 and add current char
        number = number * 26 + charValue;
    }

    return number;
}
function getRowColNum(cellKey) {
    let col = "";
    let row = "";
    for (let i of cellKey) {
        if (isNaN(i * 1)) col += i;
        else row += i;
    }
    return [row * 1, columnToNumber(col.toUpperCase())]
}
readLargeFile();