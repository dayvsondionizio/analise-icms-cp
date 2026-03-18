const XLSX = require('xlsx');
const path = require('path');

const filePath = path.join(__dirname, 'APURAÇÃO MF SANTOS MATRIZ (1).xlsx');
const workbook = XLSX.readFile(filePath);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(worksheet, { header: 'A' }).slice(1);

console.log('--- CST Analysis ---');
const cstMap = {};

data.forEach((row, i) => {
    const cst = String(row['E']).trim();
    const g = parseFloat(row['G']);
    if (!cstMap[cst]) cstMap[cst] = { count: 0, hasIcms: 0, totalG: 0 };
    cstMap[cst].count++;
    if (g > 0) {
        cstMap[cst].hasIcms++;
        cstMap[cst].totalG += g;
    }
});

console.log(JSON.stringify(cstMap, null, 2));
