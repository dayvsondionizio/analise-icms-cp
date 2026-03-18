const XLSX = require('xlsx');
const path = require('path');

const filePath = path.join(__dirname, 'APURAÇÃO MF SANTOS MATRIZ (1).xlsx');
const workbook = XLSX.readFile(filePath);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(worksheet, { header: 'A' });

console.log('--- Extensive Manual File Inspection ---');
console.log('All headers (Row 0):', JSON.stringify(data[0]));

// Sample some items
console.log('\nSample items (Rows 1-5):');
for(let i=1; i<=5; i++) {
    console.log(`Row ${i}:`, JSON.stringify(data[i]));
}

// Find columns where data exists after G
const keys = Object.keys(data[1]);
console.log('\nAll available column keys:', keys.join(', '));

// Search for any value that mentions "Outros" or "20.5" or "0.205" in any row
let foundAnything = false;
for(let i=0; i<data.length; i++) {
    const row = data[i];
    for(const key in row) {
        const val = String(row[key]);
        if (val.includes('Outros') || val.includes('OUTROS') || val.includes('Situação 2') || val.includes('SITUAÇÃO 2')) {
            console.log(`\nFound marker at row ${i} [${key}]:`, val);
            console.log('Full row:', JSON.stringify(row));
            foundAnything = true;
            break;
        }
    }
    if (foundAnything) break;
}

if (!foundAnything) {
    console.log('\nCould not find literal "Outros" markers in data. Searching for the 0.205 factor in formulas or values.');
    // Check if any column beyond G has values
    const dataColumns = keys.filter(k => k > 'G');
    if (dataColumns.length > 0) {
        console.log('Columns beyond G found:', dataColumns.join(', '));
        // Sample rows with these columns
        for(let i=1; i<10; i++) {
            const extraData = dataColumns.map(k => `${k}: ${data[i][k]}`).join(' | ');
            if (extraData) console.log(`Row ${i} extra:`, extraData);
        }
    } else {
        console.log('No columns found beyond G.');
    }
}
