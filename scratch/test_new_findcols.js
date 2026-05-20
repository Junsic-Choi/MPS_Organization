const XLSX = require('xlsx');

const findCols = (raw, keywordsMap, maxRows = 20) => {
    const result = {};
    // First pass: try exact match
    for (let r = 0; r < Math.min(maxRows, raw.length); r++) {
        const row = raw[r] || [];
        row.forEach((cell, idx) => {
            const val = String(cell || '').trim().toUpperCase();
            for (const [key, keywords] of Object.entries(keywordsMap)) {
                if (result[key] === undefined) {
                    const hasExact = keywords.some(k => val === k.toUpperCase());
                    if (hasExact) {
                        result[key] = idx;
                    }
                }
            }
        });
        if (result.Model !== undefined && (result.Site !== undefined || result.Group !== undefined)) {
            result.headerRowIdx = r;
            return result;
        }
    }

    // Second pass: fallback to includes if not found in first pass
    for (let r = 0; r < Math.min(maxRows, raw.length); r++) {
        const row = raw[r] || [];
        row.forEach((cell, idx) => {
            const val = String(cell || '').trim().toUpperCase();
            for (const [key, keywords] of Object.entries(keywordsMap)) {
                if (result[key] === undefined) {
                    const hasPartial = keywords.some(k => val.includes(k.toUpperCase()));
                    if (hasPartial) {
                        result[key] = idx;
                    }
                }
            }
        });
        if (result.Model !== undefined && (result.Site !== undefined || result.Group !== undefined)) {
            result.headerRowIdx = r;
            return result;
        }
    }
    return result;
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
