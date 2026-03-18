const XLSX = require('xlsx');
const path = require('path');

const manualPath = path.join(__dirname, 'APURAÇÃO MF SANTOS MATRIZ (1).xlsx');
const manualWorkbook = XLSX.readFile(manualPath);
const manualData = XLSX.utils.sheet_to_json(manualWorkbook.Sheets[manualWorkbook.SheetNames[0]], { header: 'A' }).slice(1);

console.log('--- Case Study: SONHO. ---');

manualData.forEach((row, i) => {
    const itemName = String(row['C']).trim().toUpperCase();
    if (itemName === 'SONHO.') {
        console.log(`Row ${i+1}: D(Contabil): ${row['D']}, G(ICMS): ${row['G']}, H(Outros): ${row['H']}`);
        if (row['H']) {
            const ratioH_D = parseFloat(row['H']) / parseFloat(row['D']);
            console.log(`Ratio H/D: ${ratioH_D.toFixed(4)}`);
        }
    }
});
