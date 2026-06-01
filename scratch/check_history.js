const XLSX = require('xlsx');

['MPS2603-1.xlsx', 'MPS2604-1.xlsx', 'MPS2605-1.xlsx'].forEach(filename => {
    try {
        const wb = XLSX.readFile('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\' + filename);
        console.log(`=== ${filename} ===`);
        const masterWs = wb.Sheets[wb.SheetNames[1]];
        const masterRaw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });
        masterRaw.forEach((row, idx) => {
            const product = row[4];
            const site = row[6];
            if (site == '1842' && product && (product.startsWith('DC') || product.startsWith('DCM'))) {
                console.log(`Row ${idx+1}: PL="${row[1]}", Group="${row[2]}", Model="${row[3]}", Product="${product}", Site="${site}"`);
            }
        });
    } catch (e) {
        console.log(`Error reading ${filename}:`, e.message);
    }
});
