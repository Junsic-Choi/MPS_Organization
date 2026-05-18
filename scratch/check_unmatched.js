const XLSX = require('xlsx');

function searchModel(filename, modelName) {
    console.log(`\nSearching for "${modelName}" in ${filename} [Sheet: MPS]...`);
    const wb = XLSX.readFile(filename);
    const ws = wb.Sheets['MPS'];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
    let found = false;
    data.forEach((row, idx) => {
        const rowStr = JSON.stringify(row);
        if (rowStr.toUpperCase().includes(modelName.toUpperCase())) {
            console.log(`Row ${idx + 1}: ${rowStr}`);
            found = true;
        }
    });
    if (!found) console.log('Not found in MPS sheet.');
}

searchModel('MPS2605-1.xlsx', 'DBM');
