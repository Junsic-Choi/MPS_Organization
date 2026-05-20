const XLSX = require('xlsx');

const wb = XLSX.readFile('MPS2605-1.xlsx');
const ws = wb.Sheets['MPS'];
const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

const masterKeywords = {
    Model: ['기종', 'Model'],
    Group: ['그룹', 'Group', 'Series'],
    Site: ['사업장', '공장', 'Site'],
    PL: ['PL', '제품군'],
    Ver: ['Ver', '버전'],
    Pjt: ['PJT', '프로젝트', 'Product Name', 'Product']
};

const findCols = (raw, keywordsMap, maxRows = 20) => {
    for (let r = 0; r < Math.min(maxRows, raw.length); r++) {
        const row = raw[r] || [];
        const rowResult = {};
        row.forEach((cell, idx) => {
            const val = String(cell || '').trim().toUpperCase();
            for (const [key, keywords] of Object.entries(keywordsMap)) {
                if (rowResult[key] === undefined) {
                    const hasExact = keywords.some(k => val === k.toUpperCase());
                    if (hasExact) rowResult[key] = idx;
                }
            }
        });
        row.forEach((cell, idx) => {
            const val = String(cell || '').trim().toUpperCase();
            for (const [key, keywords] of Object.entries(keywordsMap)) {
                if (rowResult[key] === undefined) {
                    const hasPartial = keywords.some(k => val.includes(k.toUpperCase()));
                    if (hasPartial) rowResult[key] = idx;
                }
            }
        });
        if (rowResult.Model !== undefined && (rowResult.Site !== undefined || rowResult.Group !== undefined)) {
            rowResult.headerRowIdx = r;
            return rowResult;
        }
    }
    return {};
};

const mCols = findCols(raw, masterKeywords, 50);

let monthRowIdx = -1;
let typeRowIdx = -1;
for (let r = 0; r < Math.min(50, raw.length); r++) {
    const rowStr = (raw[r] || []).join('|');
    if (rowStr.includes('월') && monthRowIdx === -1) monthRowIdx = r;
    if (rowStr.includes('생산') && rowStr.includes('판매') && r > monthRowIdx) {
        typeRowIdx = r;
        break;
    }
}

function extractMonth(s) {
    if (!s) return null;
    if (s instanceof Date) return s.getMonth() + 1;
    const str = s.toString().trim();
    const dotMatch = str.match(/(?:20)?26\.(\d+)/);
    if (dotMatch) {
        const n = parseInt(dotMatch[1]);
        if (n >= 1 && n <= 12) return n;
    }
    const monthWordMatch = str.match(/(\d+)\s*월/);
    if (monthWordMatch) {
        const n = parseInt(monthWordMatch[1]);
        if (n >= 1 && n <= 12) return n;
    }
    if (/^\d+$/.test(str)) {
        const n = parseInt(str);
        if (n >= 1 && n <= 12) return n;
    }
    return null;
}

const masterMonthCols = [];
const monthRow = raw[monthRowIdx];
const typeRow = raw[typeRowIdx];
typeRow.forEach((cell, idx) => {
    const type = (cell || '').toString().trim();
    if (type === '생산') {
        for (let c = idx; c >= 0; c--) {
            const mNum = extractMonth(monthRow[c]);
            if (mNum !== null) {
                masterMonthCols.push({ name: mNum + '월', col: idx });
                break;
            }
        }
    }
});

let sumTotal1842 = 0;
const destSums = {};

raw.forEach((row, idx) => {
    if (idx <= typeRowIdx) return;
    const mSiteIdx = mCols.Site !== undefined ? mCols.Site : 6;
    const originalSite = (row[mSiteIdx] || '').toString().trim();

    if (originalSite === '1842') {
        masterMonthCols.forEach(mCol => {
            const mMatch = mCol.name.match(/(\d{1,2})/);
            if (mMatch && parseInt(mMatch[1]) >= 10) return;
            const q = parseInt(row[mCol.col]) || 0;
            if (q > 0) {
                sumTotal1842 += q;
                // Standardize Site lookup like in extractor.js to see where it goes
                // Wait, let's just log and sum
            }
        });
    }
});

console.log('Total sum of raw 1842 (Seongju) in MPS (excluding 10월):', sumTotal1842);
