#!/usr/bin/env node
const fs = require('fs');
const readline = require('readline');

const path = require('path');
const ExcelJS = require('exceljs');


async function getValidatedInputs() {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    async function takeInput(question) {
        question += ": "
        return new Promise((resolve, reject) => {
            rl.question(question, (data) => {
                resolve(data)
            });
        })
    }

    let existingSettingPath = await takeInput("Existing Setting Path?");

    if (existingSettingPath && fs.existsSync(existingSettingPath)) {
        rl.close();
        return JSON.parse(fs.readFileSync(existingSettingPath))
    }

    const validate = {
        pathExists: (p) => fs.existsSync(p) && fs.lstatSync(p).isDirectory(),

        isExcelFile: (p) => fs.existsSync(p) && /\.(xlsx|xls|csv)$/i.test(p),

        isValidCell: (cell) => /^[A-Z]{1,3}[1-9][0-9]*$/i.test(cell),

        isValidChunk: (size) => !isNaN(size) && parseInt(size) > 0,

        isValidTemplate: (name) => !/[<>:"/\\|?*]/.test(name) && name.length > 0
    };

    let inputs = {};

    // 1. Validate Source Directory & File
    while (true) {
        inputs.dir = await takeInput("Source File Path: ");
        inputs.fileName = await takeInput("Source File Name: ");
        const fullPath = path.join(inputs.dir, inputs.fileName);

        if (validate.isExcelFile(fullPath)) break;
        console.error("❌ Error: File not found or is not a valid Excel file (.xlsx).");
    }
    while (true) {
        inputs.tabName = await takeInput("Worksheet Name: ");
        if (inputs.tabName == "") {
            console.error("❌ Error: Worksheet Name Required!");

        }
        else break;

    }
    while (true) {
        inputs.outputFileTemplate = await takeInput("Output File Template: ");
        if (inputs.outputFileTemplate == "") {
            console.error("❌ Error: Output File Template Required!");

        }
        else break;
    }


    // 2. Validate Cell Range
    while (true) {
        inputs.startCellNo = await takeInput("Start Cell (e.g., A1): ");
        inputs.endCellNo = await takeInput("Ending Cell (e.g., Z100): ");

        if (validate.isValidCell(inputs.startCellNo) && validate.isValidCell(inputs.endCellNo)) break;
        console.error("❌ Error: Invalid cell notation. Use letters followed by numbers (e.g., B2).");
    }

    // 3. Validate Chunk Size
    while (true) {
        inputs.chunkSize = await takeInput("Chunk Size (Rows per file): ");
        if (validate.isValidChunk(inputs.chunkSize)) {
            inputs.chunkSize = parseInt(inputs.chunkSize);
            break;
        }
        console.error("❌ Error: Chunk size must be a positive number.");
    }

    // 4. Validate Output Path
    while (true) {
        inputs.outputDir = await takeInput("Output Path: ");
        if (validate.pathExists(inputs.outputDir)) break;
        console.error("❌ Error: Output directory does not exist.");
    }

    rl.close();
    return inputs;
}

async function readLargeFile() {



    let settings = await getValidatedInputs();
    let { dir, fileName, tabName, startCellNo, endCellNo, chunkSize, outputDir, outputFileTemplate } = settings;


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
    let stream = fs.createWriteStream(path.join(outputDir, "Setting.json"), 'utf-8');
    stream.write(JSON.stringify(settings));

    stream.end(() => {
        console.log("closed stream")
    });
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