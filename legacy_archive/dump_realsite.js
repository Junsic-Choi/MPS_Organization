const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

try {
    const filePath = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\Real site.xlsx';
    if (!fs.existsSync(filePath)) {
        console.log('Real site.xlsx not found at: ' + filePath);
        process.exit(1);
    }
    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
    
    fs.writeFileSync('realsite_dump.json', JSON.stringify(data.slice(0, 50), null, 2));
    console.log('Real site dump completed.');
} catch (e) {
    console.error(e);
}
