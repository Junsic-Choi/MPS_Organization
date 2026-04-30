const xlsx = require('xlsx');
const fs = require('fs');

function deriveModelName(productName) {
    if (!productName) return "";
    const p = productName.toString().toUpperCase().trim();

    // Pattern: Prefix + Digit(-) -> Prefix with last digit replaced by 0 + II/III
    // Example: DVF5002-XXXX -> DVF5000 II
    const genMatch = p.match(/^([A-Z0-9]+(\d))-/);
    if (genMatch) {
        const prefixFull = genMatch[1];
        const digit = genMatch[2];
        const prefixBase = prefixFull.substring(0, prefixFull.length - 1);

        if (digit === "2") return prefixBase + "0 II";
        if (digit === "3") return prefixBase + "0 III";

        // Special case for HSP8000/DHF8000 -> 800
        if (prefixFull.match(/^(HSP|DHF)8000$/)) return prefixFull.substring(0, 3) + "800";

        return prefixFull;
    }

    // Fallback for models without hyphen but with digits (e.g. NHP6300)
    const fallbackMatch = p.match(/^(NHP|NHM|HSP|DHF)(\d+)0$/);
    if (fallbackMatch) return fallbackMatch[1] + fallbackMatch[2];

    return p.split("-")[0];
}

try {
    console.log('--- Start MPS Extraction (Node.js) ---');
    const files = fs.readdirSync('.').filter(f => f.includes('생산배포용') && f.endsWith('.xlsx'));
    if (files.length === 0) throw new Error('생산배포용 Excel File NOT found');
    const path = files[0];
    console.log('Selected File:', path);

    const workbook = xlsx.readFile(path);
    const ws = workbook.Sheets['생산배포용'] || workbook.Sheets['MPS'];
    if (!ws) throw new Error('생산배포용 Sheet NOT found');

    const data = xlsx.utils.sheet_to_json(ws, { header: 1 });

    // Column Config (0-indexed)
    const colSite = 1;    // B
    const colGroup = 2;   // C
    const colCode = 3;    // D
    const colName = 4;    // E
    const colRPM = 5;     // F

    // Month Columns (I, M, R, W, AC, AI)
    const targetCols = [8, 12, 17, 22, 28, 34];
    const monthNames = targetCols.map(c => (data[2][c] || `Col${c + 1}`).toString().trim());
    console.log('Months:', monthNames.join(', '));

    const results = [];
    results.push('Site,Group,Model,RPM,Month,SerialNo,ModelCode,ProductName');

    let totalRows = 0;
    // Data starts from Row 5 (Index 4)
    for (let r = 4; r < data.length; r++) {
        const row = data[r];
        if (!row) continue;

        const productName = row[colName];
        if (!productName || productName.toString().trim() === "") continue;

        const site = (row[colSite] || "").toString().trim();
        const group = (row[colGroup] || "").toString().trim();
        const partCode = (row[colCode] || "").toString().trim();
        const rpm = (row[colRPM] || "").toString().trim();
        const modelName = deriveModelName(productName);

        targetCols.forEach((colIdx, mIdx) => {
            const val = row[colIdx];
            let qty = 0;
            if (typeof val === 'number') qty = Math.floor(val);
            else if (!isNaN(val) && val.toString().trim() !== "") qty = parseInt(val);

            if (qty > 0) {
                const month = monthNames[mIdx];
                for (let i = 1; i <= qty; i++) {
                    totalRows++;
                    const entry = [
                        site, group, modelName, rpm, month, i, partCode, productName
                    ].map(v => `"${v.toString().replace(/"/g, '""')}"`).join(',');
                    results.push(entry);
                }
            }
        });
    }

    fs.writeFileSync('_FinalList.csv', results.join('\n'), 'utf8');
    fs.writeFileSync('_FinalList_utf8.csv', results.join('\n'), 'utf8');

    console.log('SUCCESS: Total Rows Extracted =', totalRows);
    if (totalRows !== 4650) {
        console.warn('WARNING: Total rows', totalRows, 'does not match expected 4650.');
    }
} catch (e) {
    console.error('CRITICAL:', e.message);
}
