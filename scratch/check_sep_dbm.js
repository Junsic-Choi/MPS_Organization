const XLSX = require('xlsx');

function checkAllPlanning(filename, modelPattern) {
    console.log(`\nChecking All Planning for "${modelPattern}" in ${filename}...`);
    const wb = XLSX.readFile(filename);
    const ws = wb.Sheets['MPS'];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
    
    const cols = [8, 12, 17, 22, 28, 34, 40];
    const colNames = ['4월', '5월', '6월', '7월', '8월', '9월', '10월'];
    
    data.forEach((row, idx) => {
        const rowStr = JSON.stringify(row);
        if (rowStr.toUpperCase().includes(modelPattern.toUpperCase())) {
            let planStr = '';
            cols.forEach((c, i) => {
                const q = parseInt(row[c]) || 0;
                if (q > 0) planStr += `${colNames[i]}: ${q}, `;
            });
            if (planStr) {
                console.log(`Row ${idx + 1} (Site ${row[6]}, PL ${row[1]}): ${planStr}`);
                console.log(`  Model: ${row[4]}`);
            }
        }
    });
}

checkAllPlanning('MPS2605-1.xlsx', 'DNM');

