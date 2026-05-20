const XLSX = require('xlsx');

const findCols = (raw, keywordsMap, maxRows = 20) => {
    for (let r = 0; r < Math.min(maxRows, raw.length); r++) {
        const row = raw[r] || [];
        const rowResult = {};

        // 1. Exact match pass for this row
        row.forEach((cell, idx) => {
            const val = String(cell || '').trim().toUpperCase();
            for (const [key, keywords] of Object.entries(keywordsMap)) {
                if (rowResult[key] === undefined) {
                    const hasExact = keywords.some(k => val === k.toUpperCase());
                    if (hasExact) {
                        rowResult[key] = idx;
                    }
                }
            }
        });

        // 2. Partial match pass for this row (only for keys not yet matched)
        row.forEach((cell, idx) => {
            const val = String(cell || '').trim().toUpperCase();
            for (const [key, keywords] of Object.entries(keywordsMap)) {
                if (rowResult[key] === undefined) {
                    const hasPartial = keywords.some(k => val.includes(k.toUpperCase()));
                    if (hasPartial) {
                        rowResult[key] = idx;
                    }
                }
            }
        });

        // Check if this row is the header row
        if (rowResult.Model !== undefined && (rowResult.Site !== undefined || rowResult.Group !== undefined)) {
            rowResult.headerRowIdx = r;
            // Also try to find any missing columns in the same row just in case, but they should already be in rowResult.
            return rowResult;
        }
    }
    return {};
};

const wb = XLSX.readFile('MPS2605-1.xlsx');

console.log('--- TESTING "생산배포용" SHEET ---');
const prodWsName = '생산배포용';
const prodWs = wb.Sheets[prodWsName];
const prodRaw = XLSX.utils.sheet_to_json(prodWs, { header: 1 });

const prodKeywords = {
    Model: ['기종', 'Model'],
    Group: ['기종분류', '시리즈', 'Series', '그룹'],
    Site: ['생산처', '공장', '사업장', 'Site'],
    RPM: ['RPM', '주축', 'Spindle']
};

const pCols = findCols(prodRaw, prodKeywords, 50);
console.log('pCols:', pCols);

console.log('--- TESTING "MPS" SHEET ---');
const masterWsName = 'MPS';
const masterWs = wb.Sheets[masterWsName];
const masterRaw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });

const masterKeywords = {
    Model: ['기종', 'Model'],
    Group: ['그룹', 'Group', 'Series'],
    Site: ['사업장', '공장', 'Site'],
    PL: ['PL', '제품군'],
    Ver: ['Ver', '버전'],
    Pjt: ['PJT', '프로젝트', 'Product Name']
};

const mCols = findCols(masterRaw, masterKeywords, 50);
console.log('mCols:', mCols);
