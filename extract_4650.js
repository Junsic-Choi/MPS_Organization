const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const EXCEL_FILE = '일반비_MPS2603-1(생산배포용).xlsx';
const OUTPUT_FILE = '_FinalList.csv';

try {
    console.log(`Reading ${EXCEL_FILE}...`);
    const workbook = XLSX.readFile(EXCEL_FILE);
    const sheetName = workbook.SheetNames[1]; // Sheet 2 (생산배포용)
    const ws = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

    const targetCols = [
        { idx: 4, m: "2월" },
        { idx: 7, m: "3월" },
        { idx: 8, m: "4월" },
        { idx: 9, m: "5월" },
        { idx: 10, m: "6월" },
        { idx: 12, m: "7월" }
    ];

    let results = [];
    let total = 0;

    // Start from Row 7 (index 6)
    for (let r = 6; r < data.length; r++) {
        const row = data[r];
        if (!row || !row[2]) { // If Model(Col 3, index 2) is empty
            if (r > 2000) break;
            continue;
        }

        const site = row[0] || "";
        const group = row[1] || "";
        const model = (row[2] || "").toString().trim();
        const rpm = row[3] || "";

        targetCols.forEach(col => {
            const val = row[col.idx];
            if (val && typeof val === 'number' && val > 0) {
                const qty = Math.floor(val);
                for (let q = 0; q < qty; q++) {
                    // Site, Group, Model, RPM, Month, Code, Product
                    results.push(`"${site}","${group}","${model}","${rpm}","${col.m}","",""`);
                    total++;
                }
            }
        });
    }

    console.log(`Extraction complete. Total rows: ${total}`);
    fs.writeFileSync(OUTPUT_FILE, results.join('\n'), 'utf8');

    // Save stats for verification
    fs.writeFileSync('extraction_stats.txt', `TOTAL_ROWS: ${total}`);

} catch (e) {
    console.error('ERROR:', e.message);
    fs.writeFileSync('extraction_error.txt', e.stack);
    process.exit(1);
}
