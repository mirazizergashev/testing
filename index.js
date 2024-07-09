const fs = require('fs');
const xlsx = require('xlsx');
const createCsvWriter = require('csv-writer').createObjectCsvWriter;

const inputFilePath = "C:\\Users\\Miraziz Ergashev\\Music\\report\\aaa.xlsx";
const outputFilePath = 'aaaa.csv';

// Excel faylini o'qing
const workbook = xlsx.readFile(inputFilePath);
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

// Excel ma'lumotlarini JSON formatiga o'zgartiring
const jsonData = xlsx.utils.sheet_to_json(sheet, { header: 1 });

// Ko'rsatilgan ustun nomlari
const requiredColumns = [
    
];

// Ustun nomlarini topish va indekslarini olish
const headerRowIndex = jsonData.findIndex(row => requiredColumns.every(col => row.includes(col)));
const headerRow = jsonData[headerRowIndex];
const columnIndexes = requiredColumns.map(col => headerRow.indexOf(col));

// "Total" qatori topish
const totalIndex = jsonData.findIndex(row => row[0] === 'Total');

// Ko'rsatilgan ustunlar bo'yicha ma'lumotlarni ajratib olish
const dataToWrite = jsonData.slice(headerRowIndex + 1, totalIndex).map(row =>
    columnIndexes.reduce((acc, colIndex, idx) => {
        acc[requiredColumns[idx]] = row[colIndex];
        return acc;
    }, {})
);

// CSV yozish
const headers = requiredColumns.map(column => ({ id: column, title: column }));

const csvWriter = createCsvWriter({
    path: outputFilePath,
    header: headers
});

csvWriter.writeRecords(dataToWrite)
    .then(() => {
        console.log('CSV file has been written successfully');
    })
    .catch(err => {
        console.error('Error writing CSV file:', err);
    });
