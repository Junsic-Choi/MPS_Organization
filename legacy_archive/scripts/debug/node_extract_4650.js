const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

try {
    console.log('Loading XLSX via Buffer...');
    const buffer = fs.readFileSync('data_working.xlsx');
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[1]; // 생산배포용
    const ws = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

    const row3 = data[2]; // Month row
    const row4 = data[3]; // "생산" row

    const targetCols = [];
    row4.forEach((v, idx) => {
        if (v && v.toString().includes('생산')) {
            targetCols.push({ idx, month: row3[idx] });
        }
    });

    console.log(`Found ${targetCols.length} production columns.`);

    const output = [['Site', 'Group', 'Model', 'RPM', 'Month', 'Code', 'Product']];
    let totalRows = 0;

    for (let r = 6; r < data.length; r++) {
        const row = data[r];
        if (!row[0] && !row[2]) continue; // Skip empty rows

        const site = row[0] || '';
        const group = row[1] || '';
        const model = row[2] || '';
        const rpm = row[3] || '';

        targetCols.forEach(col => {
            const qty = row[col.idx];
            if (typeof qty === 'number' && qty > 0) {
                for (let i = 0; i < qty; i++) {
                    output.push([site, group, model, rpm, col.month, '', '']);
                    totalRows++;
                }
            }
        });
    }

    const csvContent = output.map(row => row.map(v => `"${v}"`).join(',')).join('\n');
    fs.writeFileSync('_FinalList_4650.csv', csvContent);
    fs.writeFileSync('extraction_4650_stat.txt', `TOTAL ROWS: ${totalRows}\nColumns: ${targetCols.length}`);

    console.log('Extraction complete. Total rows:', totalRows);
} catch (e) {
    fs.writeFileSync('extraction_4650_error.txt', e.stack);
}
