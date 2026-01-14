const fs = require('fs')
const path = require('path');
const ExcelJS = require('exceljs');
let dir = "D:\\Shourov\\office\\documents\\Prime-Asia\\exam";
let fileName = "Pau-Exam-Data-after142.xlsx";
let tabName = "main";
let startCellNo = "AQ4";
let endCellNo = "BN542549";
//let endCellNo = "BN2549";

let chunkSize = 25000;
let outputDir = "D:\\Shourov\\office\\documents\\Prime-Asia\\exam\\exam-parts";
let outputFileTemplate = "Pau-Result-Part-"
async function readLargeFile() {

    let [startRow, startCol] = getRowColNum(startCellNo);
    let [endRow, endCol] = getRowColNum(endCellNo);

    // Use the stream reader
    console.log("Loading Excel File!");
    const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(path.join(dir, fileName));

    for await (const worksheetReader of workbookReader) {
        if (worksheetReader.name !== tabName) continue;
        let stream = null;
        let rowNum = 0;
        let chunkNo = 1;
        let headerRow = "";
        for await (const row of worksheetReader) {
            if (row.number < startRow) continue;
            if (row.number > endRow) break;
            let values = [];
            for (let i = startCol; i <= endCol; i++) {
                values.push(row.getCell(i).value);
            }

            if (row.number == startRow) {
                headerRow = getCsvRow(values);
            }
            if (rowNum % chunkSize == 0) {
                if (stream) stream.end();
                console.log(`Processing Chunk ${chunkNo}`);
                stream = fs.createWriteStream(path.join(outputDir, `${outputFileTemplate}${chunkNo}.csv`));
                stream.write(headerRow)
                chunkNo++;
                rowNum++;
                if (row.number == startRow) continue;
            }

            let rowData = getCsvRow(values);
            if (rowData.includes('object')) {
                rowData = getCsvRow(values);
            }
            stream.write(rowData);
            rowNum++;


        }
        if (stream) stream.end();
    }
}

function getCsvRow(row) {
    let rowData = "";
    for (let i = 0; i < row.length; i++) {
        let cell = row[i]
        try {
            let temp = (cell);

            if (temp != null && typeof temp == 'object' && temp['result'] != null) rowData += temp.result;
            else rowData += cell ?? " ";
        } catch (error) {
            rowData += cell ?? " ";

        }

        if (i < row.length - 1) rowData += ","
    }
    return rowData + '\n';
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