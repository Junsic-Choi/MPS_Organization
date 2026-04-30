const XLSX = require('xlsx');
const fs = require('fs');
try {
    const wb = XLSX.readFile('site.xlsx');
    const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
    fs.writeFileSync('site_map.json', JSON.stringify(data, null, 2));
    console.log('Successfully saved site_map.json');
} catch (e) {
    console.error('Failed to parse site.xlsx:', e.message);
}
