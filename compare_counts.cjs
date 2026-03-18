const XLSX = require('xlsx');
const path = require('path');

const manualPath = path.join(__dirname, 'APURAÇÃO MF SANTOS MATRIZ (1).xlsx');
const apagadoPath = path.join(__dirname, 'APURAÇÃO MF SANTOS MATRIZ (1) - APAGADO.xlsx');

const manualWorkbook = XLSX.readFile(manualPath);
const apagadoWorkbook = XLSX.readFile(apagadoPath);

const manualData = XLSX.utils.sheet_to_json(manualWorkbook.Sheets[manualWorkbook.SheetNames[0]], { header: 'A' }).slice(1);
const apagadoData = XLSX.utils.sheet_to_json(apagadoWorkbook.Sheets[apagadoWorkbook.SheetNames[0]], { header: 'A' }).slice(1);

console.log('--- Row Count & Value Comparison ---');

let manualTotalRows = 0;
let manualTotalValue = 0;
dataMap = {};

manualData.forEach(row => {
    if (row['H'] && parseFloat(row['H']) > 0) {
        manualTotalRows++;
        manualTotalValue += parseFloat(row['H']);
    }
});

console.log('MANUAL FILE (Column H):');
console.log('Total Rows with Outros Debitos:', manualTotalRows);
console.log('Total Value of Outros Debitos:', manualTotalValue.toFixed(2));

// The 45 items we identified
const outrosItemsSet = new Set();
manualData.forEach(row => {
    if (row['H'] && parseFloat(row['H']) > 0) {
        outrosItemsSet.add(String(row['C']).trim().toUpperCase());
    }
});

// Calculate what the system SHOULD find in the APAGADO file
let apagadoTotalPossibleRows = 0;
let apagadoTotalPossibleValue = 0;
let apagadoItemsFound = new Set();

apagadoData.forEach(row => {
    const itemName = String(row['C']).trim().toUpperCase();
    if (outrosItemsSet.has(itemName)) {
        apagadoTotalPossibleRows++;
        apagadoTotalPossibleValue += (parseFloat(row['D']) * 0.205);
        apagadoItemsFound.add(itemName);
    }
});

console.log('\nAPAGADO FILE (Potential matches for Situation 2):');
console.log('Total Rows matching Situation 2 items:', apagadoTotalPossibleRows);
console.log('Total Potential Value (Contabil * 0.205):', apagadoTotalPossibleValue.toFixed(2));
console.log('Unique Outros items found in APAGADO file:', apagadoItemsFound.size);

if (manualTotalRows !== apagadoTotalPossibleRows) {
    console.log('\nDISCREPANCY DETECTED in row counts!');
}
if (Math.abs(manualTotalValue - apagadoTotalPossibleValue) > 1) {
    console.log('\nDISCREPANCY DETECTED in total values!');
}

// Check for "hasIcms" condition mismatch
let apagadoNoIcmsRows = 0;
apagadoData.forEach(row => {
    const itemName = String(row['C']).trim().toUpperCase();
    const g = parseFloat(row['G']);
    if (outrosItemsSet.has(itemName) && (!g || g <= 0)) {
        apagadoNoIcmsRows++;
    }
});
console.log('\nRows in APAGADO file that match Outros items but HAVE NO ICMS (this could break the rule):', apagadoNoIcmsRows);
