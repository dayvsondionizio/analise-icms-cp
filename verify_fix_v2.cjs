const XLSX = require('xlsx');
const path = require('path');
const Database = require('better-sqlite3');

const db = new Database('data/tax_rules.db');
const rules = db.prepare('SELECT * FROM tax_rules').all();

const apagadoPath = path.join(__dirname, 'APURAÇÃO MF SANTOS MATRIZ (1) - APAGADO.xlsx');
const workbook = XLSX.readFile(apagadoPath);
const data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 'A' }).slice(1);

// New Logic with ICMS zeroed for Sit 2
function processLine(row) {
    const itemName = String(row['C'] || '').trim().toUpperCase();
    const valorContabil = parseFloat(row['D']) || 0;
    const valorIcms = parseFloat(row['G']) || 0;
    const rowHasIcms = valorIcms > 0;

    const matchingRule = rules.find(r => 
        r.item && r.item.trim().toUpperCase() === itemName && 
        (r.situacao !== 1 || r.hasIcms === (rowHasIcms ? 1 : 0))
    );

    let finalIcms = valorIcms;
    let finalOutros = 0;

    if (matchingRule && matchingRule.situacao === 2) {
        finalOutros = valorContabil * 0.205;
        finalIcms = 0; // The fix
    }

    return { finalIcms, finalOutros };
}

let totalIcms = 0;
let totalOutros = 0;
data.forEach(row => {
    const res = processLine(row);
    totalIcms += res.finalIcms;
    totalOutros += res.finalOutros;
});

console.log('--- Simulation with Fix ---');
console.log('Total ICMS in Outros Débitos items:', totalIcms.toFixed(2));
console.log('Total Outros Débitos (Base * 20,5%):', totalOutros.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }));
