const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const normCache = new Map();
function norm(s) {
    if (!s) return "";
    const raw = s.toString().toUpperCase().trim();
    if (normCache.has(raw)) return normCache.get(raw);
    let res = raw;
    res = res.replace(/ III/g, '3').replace(/ II/g, '2').replace(/ I/g, '1');
    res = res.replace(/III/g, '3').replace(/II/g, '2');
    if (res.startsWith('DCM')) res = 'DC' + res.substring(3);
    if (res.startsWith('PUMA')) res = 'P' + res.substring(4);
    if (res.startsWith('LYNX')) res = res.substring(4);
    const final = res.replace(/[^A-Z0-9]/g, '');
    normCache.set(raw, final);
    return final;
}

const FILE_PATH = path.join(__dirname, 'MPS2603-1.xlsx');
const workbook = XLSX.readFile(FILE_PATH);
const prodName = workbook.SheetNames.find(n => n.includes('배포'));
const sheet = workbook.Sheets[prodName];
const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 });

let out = "--- Production Models Normalizing to XG800 ---\n";
let lastSite = "";
for (let r = 0; r < raw.length; r++) {
    const row = raw[r] || [];
    if (row[0]) lastSite = row[0].toString().trim();
    const model = (row[2] || '').toString();
    if (!model) continue;
    if (norm(model) === 'XG800') {
        out += `Row ${r+1}: Site=${lastSite}, Model=${model}, 2월=${row[4]||0}, 3월=${row[7]||0}, 6월=${row[10]||0}\n`;
    }
}

fs.writeFileSync('debug_xg800_culprits.txt', out);
console.log('Done.');
