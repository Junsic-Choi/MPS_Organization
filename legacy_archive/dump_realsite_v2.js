const XLSX = require('xlsx');
const fs = require('fs');

try {
    const filePath = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\Real site.xlsx';
    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
    fs.writeFileSync('realsite_full_dump.json', JSON.stringify(data, null, 2));
    console.log('Real site dumped.');
} catch (e) {
    fs.writeFileSync('realsite_error.txt', e.stack);
}
