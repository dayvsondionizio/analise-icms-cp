const XLSX = require('xlsx');
const path = require('path');

const filePath = path.join(__dirname, 'APURAÇÃO MF SANTOS MATRIZ (1).xlsx');
const workbook = XLSX.readFile(filePath);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(worksheet, { header: 'A' }).slice(1); // skip header

console.log('--- Ratio Analysis (G / D) ---');
let count205 = 0;
let countOther = 0;
let totalWithIcms = 0;

data.forEach((row, i) => {
    const d = parseFloat(row['D']);
    const g = parseFloat(row['G']);
    if (g > 0 && d > 0) {
        totalWithIcms++;
        const ratio = g / d;
        if (Math.abs(ratio - 0.205) < 0.001) {
            count205++;
        } else {
            countOther++;
            if (countOther < 10) {
                console.log(`Other ratio found at row ${i+1}: ${row['C']} - Ratio: ${ratio.toFixed(4)} (G: ${g}, D: ${d})`);
            }
        }
    }
});

console.log('\nSummary:');
console.log('Total items with ICMS (>0):', totalWithIcms);
console.log('Items with ~20.5% ratio:', count205);
console.log('Items with OTHER ratio:', countOther);
