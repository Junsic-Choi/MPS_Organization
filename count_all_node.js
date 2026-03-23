const XLSX = require('xlsx');
const fs = require('fs');
try {
    const buffer = fs.readFileSync('data_working.xlsx');
    const wb = XLSX.read(buffer, { type: 'buffer' });
    let res = "Sheet Row Counts:\n";
    wb.SheetNames.forEach(name => {
        const ws = wb.Sheets[name];
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
        const rows = range.e.r - range.s.r + 1;
        res += `- ${name}: ${rows} rows\n`;
    });
    fs.writeFileSync('all_sheet_counts_node.txt', res);
} catch (e) {
    fs.writeFileSync('all_sheet_counts_node.txt', 'ERROR: ' + e.message);
}
