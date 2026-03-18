const XLSX = require('xlsx');
const path = require('path');

const filePath = path.join(__dirname, 'APURAÇÃO MF SANTOS MATRIZ (1) - APAGADO.xlsx');
const workbook = XLSX.readFile(filePath);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(worksheet, { header: 'A' });

console.log('--- APAGADO File Inspection ---');
console.log('Row 0:', JSON.stringify(data[0]));
console.log('Row 5:', JSON.stringify(data[5]));
