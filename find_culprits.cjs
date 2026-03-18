const XLSX = require('xlsx');
const path = require('path');
const Database = require('better-sqlite3');

const db = new Database('data/tax_rules.db');
const rules = db.prepare('SELECT item FROM tax_rules WHERE situacao = 2').all();
const outrosItemsSet = new Set(rules.map(r => r.item.trim().toUpperCase()));

const apagadoPath = path.join(__dirname, 'APURAÇÃO MF SANTOS MATRIZ (1) - APAGADO.xlsx');
const workbook = XLSX.readFile(apagadoPath);
const data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 'A' }).slice(1);

console.log('--- Finding Situation 2 items with ICMS Highlight ---');
let sumIcms = 0;
const culprits = [];

data.forEach((row, i) => {
    const itemName = String(row['C'] || '').trim().toUpperCase();
    const valIcms = parseFloat(row['G']) || 0;
    
    if (outrosItemsSet.has(itemName) && valIcms > 0) {
        sumIcms += valIcms;
        culprits.push({
            row: i + 2,
            item: row['C'],
            valContabil: row['D'],
            valIcms: row['G']
        });
    }
});

console.log('Total Unexpected ICMS Sum:', sumIcms.toFixed(2));
console.log('Items found:');
culprits.forEach(c => {
    console.log(`Row ${c.row}: ${c.item} | Contabil: ${c.valContabil} | ICMS: ${c.valIcms}`);
});
