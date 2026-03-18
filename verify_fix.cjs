const XLSX = require('xlsx');
const path = require('path');
const Database = require('better-sqlite3');

const db = new Database('data/tax_rules.db');
const rules = db.prepare('SELECT * FROM tax_rules').all();

const apagadoPath = path.join(__dirname, 'APURAÇÃO MF SANTOS MATRIZ (1) - APAGADO.xlsx');
const workbook = XLSX.readFile(apagadoPath);
const data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 'A' }).slice(1);

// New Logic from App.tsx
function processLine(row) {
    const itemName = String(row['C'] || '').trim().toUpperCase();
    const valorContabil = parseFloat(row['D']) || 0;
    const valorIcms = parseFloat(row['G']) || 0;
    const rowHasIcms = valorIcms > 0;

    const matchingRule = rules.find(r => 
        r.item && r.item.trim().toUpperCase() === itemName && 
        (r.situacao !== 1 || r.hasIcms === (rowHasIcms ? 1 : 0))
    );

    if (matchingRule && matchingRule.situacao === 2) {
        return valorContabil * 0.205;
    }
    return 0;
}

let totalOutros = 0;
data.forEach(row => {
    totalOutros += processLine(row);
});

console.log('Result with New Logic:');
console.log('Total Outros Débitos:', totalOutros.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }));
console.log('Target (20,5% of manual base): ~R$ 10.139,11');
