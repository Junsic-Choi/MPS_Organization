const xlsx = require('xlsx');

try {
    const wb = xlsx.readFile('일반비_MPS2603-1(생산배포용).xlsx');
    const ws = wb.Sheets['MPS'];
    const data = xlsx.utils.sheet_to_json(ws, { header: 1 });

    let totalSum = 0;
    const targetCols = [8, 12, 17, 22, 28, 34]; // I, M, R, W, AC, AI (0-indexed)

    // Check Row 5 (index 4) and Row 6 onwards (index 5+)
    for (let r = 4; r < data.length; r++) {
        const row = data[r];
        if (!row) continue;
        targetCols.forEach((colIdx) => {
            const val = row[colIdx];
            if (typeof val === 'number') totalSum += val;
            else if (typeof val === 'string' && !isNaN(val) && val.trim() !== '') totalSum += parseInt(val);
        });
    }

    console.log('Total Quantity Sum (from Row 5):', totalSum);
} catch (e) {
    console.error(e);
}
