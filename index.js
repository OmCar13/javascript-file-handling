const ExcelJS = require('exceljs');
const fs = require('fs');
const chokidar = require('chokidar');

const inputFile = "input.xlsx";
const outputFile = "output.xlsx";

async function dummy_data() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sheet_1');
    
    worksheet.columns = [
        { header: 'Key', key: 'key' },
        { header: 'Values', key: 'values' }
    ];
    
    const testData = [
        [1, ['apple', 'banana', 'cherry']],
        [2, ['red', 'green', 'blue']],
        [3, ['cat', 'dog', 'fish']]
    ];
    
    testData.forEach(([key, values]) => {
        worksheet.addRow([key, ...values]);
    });
    
    await workbook.xlsx.writeFile(inputFile);
    console.log('Initial test data created in', inputFile);
}

async function readExcel(filename) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filename);
    const worksheet = workbook.getWorksheet(1);
    
    const dataMap = new Map();

    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) {
            const key = row.getCell(1).value;
            const values = row.values.slice(2).filter(val => val !== null);
            dataMap.set(key, values);
        }
    });

    return dataMap;
}

async function writeMap(dataMap, filename) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sheet_1');

    worksheet.columns = [
        { header: 'Key', key: 'key' },
        { header: 'Values', key: 'values' }
    ];

    dataMap.forEach((values, key) => {
        const row = worksheet.addRow([key, ...values]);
    });
    await workbook.xlsx.writeFile(filename);
}
    
async function handleFileChange(){
    try {
        const dataMap = await readExcel(inputFile);
        await writeMap(dataMap, outputFile);
        console.log('Changes detected and saved to', outputFile);
    } catch (error) {
        console.error('Error processing file:', error);
    }
}

async function main() {
    if (!fs.existsSync(inputFile)) {
        await dummy_data();
    }

    chokidar.watch(inputFile).on('change', handleFileChange);

    console.log('Watching for changes in', inputFile);
}

main().catch(console.error);