const XLSX = require('xlsx');
const path = require('path');

const filePath = path.join(__dirname, 'APURAÇÃO MF SANTOS MATRIZ (1).xlsx');
const workbook = XLSX.readFile(filePath);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(worksheet, { header: 'A' });

console.log('--- Manual File Inspection ---');
// Find where the data starts (usually there's a header)
let headerRow = -1;
for(let i=0; i<Math.min(data.length, 50); i++) {
    const row = data[i];
    if (Object.values(row).some(v => String(v).toUpperCase().includes('ITEM') || String(v).toUpperCase().includes('VALOR CONTAB'))) {
        headerRow = i;
        break;
    }
}

if (headerRow === -1) {
    console.log('Could not find header row automatically.');
    headerRow = 0;
}

console.log('Header Row Index:', headerRow);
console.log('Sample data row:', JSON.stringify(data[headerRow + 5]));

// Count items and identify potential "20,5%" markers
let count205 = 0;
const items205 = [];

// Look for a column that might indicate the 20.5% calculation or situation 2
data.slice(headerRow + 1).forEach((row, i) => {
    const vals = Object.values(row).map(v => String(v));
    // Check if any value looks like 20.5% calculation or if there's a status
    if (vals.some(v => v.includes('20,5') || v.includes('0,205'))) {
        count205++;
        items205.push({
            item: row['C'] || row['D'] || row['B'], // Need to identify the Item column
            all: row
        });
    }
});

console.log('Found 20.5% items count:', count205);
if (items205.length > 0) {
    console.log('First 5 items with 20.5%:', JSON.stringify(items205.slice(0, 5), null, 2));
}
