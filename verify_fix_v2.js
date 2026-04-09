const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const normCache = new Map();
function norm(s) {
    if (!s) return "";
    const raw = s.toString();
    if (normCache.has(raw)) return normCache.get(raw);
    let res = raw.toUpperCase().replace(/[^A-Z0-9]/g, '');
    res = res.replace(/(\D)(\d{2})00(\D|$)/g, '$1$2$3');
    if (res.startsWith('DCM')) res = 'DC' + res.substring(3);
    if (res.startsWith('PUMA')) res = 'P' + res.substring(4);
    if (res.startsWith('LYNX')) res = res.substring(4);
    const final = res.replace(/[^A-Z0-9]/g, '');
    normCache.set(raw, final);
    return final;
}

function getBase(s) {
    if (!s) return "";
    const parts = s.toString().split('-');
    return norm(parts[0]);
}

const FILE_PATH = path.join(__dirname, 'MPS2603-1.xlsx');
const workbook = XLSX.readFile(FILE_PATH);
const mpsName = workbook.SheetNames.find(n => n.toUpperCase().includes('MPS'));
const prodName = workbook.SheetNames.find(n => n.includes('배포'));

const mpsRaw = XLSX.utils.sheet_to_json(workbook.Sheets[mpsName], { header: 1 });
const prodRaw = XLSX.utils.sheet_to_json(workbook.Sheets[prodName], { header: 1 });

const mpsPool = [];
mpsRaw.forEach((row, i) => {
    if (i === 0) return;
    const prod = (row[4] || '').toString();
    if (!prod) return;
    const mNorm = getBase(prod); // Use getBase like server.js
    mpsPool.push({ code: row[3], product: prod, mNorm });
});

let out = "--- Post-Fix Hutec LYNX XG800 Verification (with getBase) ---\n";
let lastSite = "";
prodRaw.slice(6).forEach((row, i) => {
    if (row[0]) lastSite = row[0].toString().trim();
    if (lastSite.includes('휴텍')) {
        const model = (row[2] || '').toString();
        if (model.includes('LYNX XG800')) {
            const pNorm = norm(model);
            let match = mpsPool.find(m => m.mNorm === pNorm || 'L' + pNorm === m.mNorm);
            if (match) {
                out += `FOUND: Model=${model} -> Product=${match.product} (Code=${match.code})\n`;
            } else {
                out += `NOT FOUND: Model=${model} (Normalized as ${pNorm})\n`;
            }
        }
    }
});

fs.writeFileSync('debug_verify_fix_v2.txt', out);
console.log('Verification done.');
