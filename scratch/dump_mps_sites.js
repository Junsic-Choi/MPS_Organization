const XLSX = require('xlsx');

function dumpMpsSites(filename) {
    console.log(`\n--- Sites in ${filename} [Sheet: MPS] ---`);
    const wb = XLSX.readFile(filename);
    const ws = wb.Sheets['MPS'];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
    const sites = new Set();
    for (let i = 5; i < data.length; i++) {
        if (data[i] && data[i][6]) sites.add(data[i][6]);
    }
    console.log('Unique Sites (Col 6):', Array.from(sites));
}

dumpMpsSites('MPS2605-1.xlsx');
