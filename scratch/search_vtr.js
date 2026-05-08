const XLSX = require('xlsx');
const wb = XLSX.readFile('MPS2604-1.xlsx');
const ws = wb.Sheets[wb.SheetNames.find(n => n.toUpperCase() === 'MPS')];
const raw = XLSX.utils.sheet_to_json(ws, {header:1});

console.log('--- VTR Search in Master Plan ---');
raw.forEach((row, i) => {
    const s = JSON.stringify(row);
    if (s.includes('VTR162') || s.includes('VTR121')) {
        console.log(`Row ${i} Site[${row[1]}]:`, row.slice(0, 40)); 
    }
});
