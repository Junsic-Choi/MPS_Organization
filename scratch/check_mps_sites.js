const XLSX = require('xlsx');
const wb = XLSX.readFile('MPS2605-1.xlsx');
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

const findCols = (raw, keywordsMap, maxRows = 20) => {
    const result = {};
    for (let r = 0; r < Math.min(maxRows, raw.length); r++) {
        const row = raw[r] || [];
        row.forEach((cell, idx) => {
            const val = String(cell || '').trim();
            for (const [key, keywords] of Object.entries(keywordsMap)) {
                if (keywords.some(k => val.includes(k)) && result[key] === undefined) {
                    result[key] = idx;
                }
            }
        });
        if (result.Model !== undefined && (result.Site !== undefined || result.Group !== undefined)) {
            result.headerRowIdx = r;
            break;
        }
    }
    return result;
};

const mCols = findCols(masterRaw, masterKeywords, 50);
console.log('mCols detected:', mCols);

const siteMapCounts = {};
for (let r = mCols.headerRowIdx + 1; r < masterRaw.length; r++) {
    const row = masterRaw[r];
    if (!row) continue;
    const siteVal = row[mCols.Site];
    siteMapCounts[siteVal] = (siteMapCounts[siteVal] || 0) + 1;
}
console.log('MPS sheet Site column raw counts:', siteMapCounts);
