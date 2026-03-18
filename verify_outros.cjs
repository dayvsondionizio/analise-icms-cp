const XLSX = require('xlsx');
const path = require('path');
const Database = require('better-sqlite3');

const filePath = path.join(__dirname, 'APURAÇÃO MF SANTOS MATRIZ (1).xlsx');
const workbook = XLSX.readFile(filePath);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(worksheet, { header: 'A' }).slice(1);

const db = new Database('data/tax_rules.db');

console.log('--- Analyzing Manual Outros Débitos (Column H) ---');

const manualOutrosItems = new Set();
data.forEach(row => {
    if (row['H'] && parseFloat(row['H']) > 0) {
        manualOutrosItems.add(String(row['C']).trim().toUpperCase());
    }
});

console.log('Total unique items with manual Outros Débitos:', manualOutrosItems.size);

const sampleItems = Array.from(manualOutrosItems).slice(0, 10);
console.log('Sample Manual Outros Items:', sampleItems);

console.log('\n--- Checking these items in DB ---');
for(const item of sampleItems) {
    const rule = db.prepare('SELECT * FROM tax_rules WHERE item = ?').get(item);
    if (rule) {
        console.log(`Item: ${item} | DB Situacao: ${rule.situacao} | DB Acao: ${rule.acao}`);
    } else {
        console.log(`Item: ${item} | NOT FOUND IN DB`);
    }
}

// Find items that ARE Situation 2 in DB but NOT in manual column H
const dbSituacao2 = db.prepare('SELECT item FROM tax_rules WHERE situacao = 2').all().map(r => r.item);
console.log('\nTotal items in DB marked as Situation 2:', dbSituacao2.length);

const mismatch = dbSituacao2.filter(item => !manualOutrosItems.has(item));
console.log('Quantity of items matching DB Situation 2 but MISSING in manual column H:', mismatch.length);
if (mismatch.length > 0) {
    console.log('Sample mismatches (DB=2, Manual=Empty):', mismatch.slice(0, 5));
}
