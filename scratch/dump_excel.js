const XLSX = require('xlsx');
const filename = 'MPS2605-1.xlsx';
const wb = XLSX.readFile(filename);

wb.SheetNames.forEach((name, idx) => {
    console.log(`\n--- Sheet ${idx}: ${name} ---`);
    const ws = wb.Sheets[name];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });
    raw.slice(0, 10).forEach((row, rIdx) => {
        console.log(`Row ${rIdx}:`, JSON.stringify(row.slice(0, 15)));
    });
});
