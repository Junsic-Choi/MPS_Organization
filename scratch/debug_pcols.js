const XLSX = require('xlsx');
const wb = XLSX.readFile('MPS2605-1.xlsx');
const prodSheet = wb.SheetNames.find(name => ['생산배포', '배포용', 'Production'].some(k => name.includes(k))) || wb.SheetNames[0];
const ws = wb.Sheets[prodSheet];
const prodRaw = XLSX.utils.sheet_to_json(ws, { header: 1 });

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

const prodKeywords = {
    Model: ['기종', 'Model'],
    Group: ['시리즈', 'Series', '그룹', '기종분류'],
    Site: ['공장', '사업장', 'Site', '생산처'],
    RPM: ['RPM', '주축', 'Spindle']
};

const pCols = findCols(prodRaw, prodKeywords, 50);
console.log('pCols:', pCols);

// Let's print out what metaMap is generated with these columns
const prodHeaderIdx = pCols.headerRowIdx !== undefined ? pCols.headerRowIdx : -1;
console.log('prodHeaderIdx:', prodHeaderIdx);

const metaMap = {};
let lastMetaSite = '', lastMetaGroup = '';

prodRaw.forEach((row, idx) => {
    if (idx <= prodHeaderIdx) return;
    const s = (row[pCols.Site] || '').toString().trim();
    const g = (row[pCols.Group] || '').toString().trim();
    const m = (row[pCols.Model] || '').toString().trim();
    const rpm = (row[pCols.RPM] || '').toString().trim();
    
    if (s) lastMetaSite = s;
    if (g) lastMetaGroup = g;
    
    if (m) {
        // Let's log first few to see
        if (idx < 20) {
            console.log(`Row ${idx} - original s: "${s}", g: "${g}", m: "${m}". lastMetaSite: "${lastMetaSite}", lastMetaGroup: "${lastMetaGroup}"`);
        }
        const mKey = m.replace(/PUMA|LYNX/g, '').replace(/\s+/g, '').trim().toUpperCase(); // simplified key for debug
        if (!metaMap[mKey]) {
            metaMap[mKey] = { site: lastMetaSite, group: lastMetaGroup, model: m, rpm: rpm };
        }
    }
});

console.log('Sample metaMap keys:', Object.keys(metaMap).slice(0, 10));
console.log('Sample metaMap values:', Object.values(metaMap).slice(0, 10));
