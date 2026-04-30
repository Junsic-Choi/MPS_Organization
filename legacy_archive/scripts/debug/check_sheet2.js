const XLSX = require('xlsx');
const fs = require('fs');

try {
    console.log('Reading workbook...');
    const workbook = XLSX.readFile('data_working.xlsx');
    const sheetName = workbook.SheetNames[1]; // 2nd sheet: 생산배포용
    const ws = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

    const row3 = data[2]; // Row 3 (index 2) - Month?
    const row4 = data[3]; // Row 4 (index 3) - "생산"?

    let stats = `Sheet: ${sheetName}\nRow 4 (Production Scan):\n`;
    let totalSum = 0;
    const targetCols = [];

    row4.forEach((v, idx) => {
        if (v && (v.toString().includes('생산') || v.toString().includes('Production'))) {
            targetCols.push(idx);
            stats += `Col ${idx}: R3=[${row3[idx]}] R4=[${v}]\n`;
        }
    });

    // Sum data rows (from index 6 to the end)
    for (let r = 6; r < data.length; r++) {
        const row = data[r];
        targetCols.forEach(c => {
            const val = row[c];
            if (typeof val === 'number') totalSum += val;
        });
    }

    stats += `\nTOTAL SUM: ${totalSum}\n`;
    fs.writeFileSync('node_sheet2_stat.txt', stats);
    console.log('Total Sum:', totalSum);
} catch (e) {
    fs.writeFileSync('node_sheet2_stat.txt', 'ERROR: ' + e.stack);
}
