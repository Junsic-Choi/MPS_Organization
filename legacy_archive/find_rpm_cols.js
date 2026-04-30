const XLSX = require('xlsx');
const fs = require('fs');

try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    let report = "";

    workbook.SheetNames.forEach(name => {
        const sheet = workbook.Sheets[name];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        report += `--- Sheet: ${name} ---\n`;
        // Search first 50 rows for "RPM" string
        for(let r=0; r<50; r++) {
            if(data[r]) {
                data[r].forEach((cell, c) => {
                    if(cell && cell.toString().toUpperCase().includes('RPM')) {
                        report += `FOUND 'RPM' at Row ${r}, Col ${c}: "${cell}"\n`;
                    }
                });
            }
        }
    });

    fs.writeFileSync('rpm_search_results.txt', report);
} catch (e) {
    fs.writeFileSync('rpm_search_results.txt', 'ERROR: ' + e.message);
}
