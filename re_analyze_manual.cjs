const XLSX = require('xlsx');
const path = require('path');

const manualPath = path.join(__dirname, 'APURAÇÃO MF SANTOS MATRIZ (1).xlsx');
const manualWorkbook = XLSX.readFile(manualPath);
const manualData = XLSX.utils.sheet_to_json(manualWorkbook.Sheets[manualWorkbook.SheetNames[0]], { header: 'A' }).slice(1);

console.log('--- RE-ANALYSIS: Outros Débitos (Column H) ---');

let totalManualBase = 0;
let rowCountH = 0;

manualData.forEach(row => {
    const valH = parseFloat(row['H']);
    if (valH > 0) {
        totalManualBase += valH;
        rowCountH++;
    }
});

const calculatedTax = totalManualBase * 0.205;

console.log('Total manual rows (Col H > 0):', rowCountH);
console.log('Total manual base value (100%):', totalManualBase.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }));
console.log('Calculated Tax (Base * 20,5%):', calculatedTax.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }));

console.log('\n--- Comparing with CST 0 (Column G) ---');
// Some items might have CST 0 and be "H", some might be just CST 0.
let totalCST0Value = 0;
let rowCountG = 0;
manualData.forEach(row => {
    const valG = parseFloat(row['G']);
    if (valG > 0) {
        totalCST0Value += valG;
        rowCountG++;
    }
});
console.log('Total rows with ICMS Highlight (Col G > 0):', rowCountG);
console.log('Total ICMS highlighted value (Col G sum):', totalCST0Value.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }));
