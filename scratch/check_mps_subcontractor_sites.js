const fs = require('fs');
const extractor = require('../extractor.js');

const buf = fs.readFileSync('MPS2605-1.xlsx');
const result = extractor.processMpsFile(buf);

const mapping = {};

// We need to access raw rows or write a simple parser to see the original row[mSiteIdx] and final Site
const XLSX = require('xlsx');
const wb = XLSX.read(buf);
const prodWs = wb.Sheets['생산배포용'];
const masterWs = wb.Sheets['MPS'];
const prodRaw = XLSX.utils.sheet_to_json(prodWs, { header: 1 });
const masterRaw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });

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

const pCols = findCols(prodRaw, {
    Model: ['기종', 'Model'],
    Group: ['기종분류', '시리즈', 'Series', '그룹'],
    Site: ['생산처', '공장', '사업장', 'Site'],
    RPM: ['RPM', '주축', 'Spindle']
}, 50);

const mCols = findCols(masterRaw, {
    Model: ['기종', 'Model'],
    Group: ['그룹', 'Group', 'Series'],
    Site: ['사업장', '공장', 'Site'],
    PL: ['PL', '제품군'],
    Ver: ['Ver', '버전'],
    Pjt: ['PJT', '프로젝트', 'Product Name', 'Product']
}, 50);

function getMatchKey(s) {
    if (!s) return '';
    let n = s.toString().toUpperCase().trim();
    n = n.replace(/PUMA|LYNX/g, '').replace(/\s+/g, '').trim();
    if (n.startsWith('P') || (n.startsWith('L') && !n.startsWith('LEO'))) {
        n = n.substring(1);
    }
    n = n.replace(/VIII/g, '8').replace(/VII/g, '7').replace(/VI/g, '6').replace(/IV/g, '4').replace(/IX/g, '9');
    n = n.replace(/III/g, '3').replace(/II/g, '2');
    if (n.includes('SMX')) {
        n = n.replace(/SMX2(?![0-9])/g, 'SMX21');
        n = n.replace(/2100/g, '21').replace(/3100/g, '31').replace(/5100/g, '51');
        n = n.replace(/SYYB/g, 'SYY').replace(/STB/g, 'SB');
    }
    if (n.includes('DNM750L/50') || n === 'DNM755L') return 'D755L';
    if (n.includes('DNM750/50') || n === 'DNM7550' || n === 'DNM755') return 'D755';
    if (n.includes('ST38GS')) return 'ST38GS2'; 
    if (n.includes('ST10GS')) return 'ST1GS2';
    if (n.includes('DST20')) return 'DST20';
    if (n.includes('LEO16')) return 'LEO16';
    if (n.startsWith('VTR')) {
        let m = n.match(/VTR(\d{2})/);
        if (m) return 'VTR' + m[1];
    }
    if (n.startsWith('VCF') || n.startsWith('VF') || n.startsWith('DVF')) {
        let m = n.match(/(?:VCF|VF|DVF)(\d)/);
        if (m) return 'VF' + m[1];
    }
    n = n.replace(/DNM(\d+)0\/(\d+)/, 'DNM$1$2');
    if (n.startsWith('MYNX')) n = 'M' + n.substring(4);
    else if (n.startsWith('VMX')) n = 'M' + n.substring(3);
    else if (n.startsWith('VM')) n = 'V' + n.substring(2);
    else if (n.startsWith('MP')) n = 'M' + n.substring(2);
    else if (n.startsWith('DNM')) n = 'D' + n.substring(3);
    else if (n.startsWith('DBC')) n = 'DB' + n.substring(3);
    else if (n.startsWith('DCM')) n = 'DC' + n.substring(3);
    else if (n.startsWith('DVF')) n = 'V' + n.substring(3);
    else if (n.startsWith('VT') && !n.startsWith('VTR')) n = 'V' + n.substring(2);
    else if (n.startsWith('TT') && !n.startsWith('TTR')) n = 'T' + n.substring(2);
    let key = n.replace(/[^A-Z1-9]/g, '');
    if (key === 'DST2') key = 'DST20';
    key = key.replace(/0/g, '');
    if (key === 'D755' && n.includes('L')) return 'D755L';
    key = key.replace(/[A-Z]+$/, '');
    if (key.startsWith('DC')) key = key.replace(/[A-Z]$/, '');
    return key;
}

const metaMap = {};
let lastSite = '';
let lastGroup = '';
prodRaw.forEach((row, idx) => {
    if (idx <= pCols.headerRowIdx) return;
    const s = (row[pCols.Site] || '').toString().trim();
    const g = (row[pCols.Group] || '').toString().trim();
    const m = (row[pCols.Model] || '').toString().trim();
    const rpm = (row[pCols.RPM] || '').toString().trim();
    if (s) lastSite = s;
    if (g) lastGroup = g;
    if (m) {
        const mKey = getMatchKey(m);
        if (!metaMap[mKey]) {
            metaMap[mKey] = { site: lastSite, group: lastGroup, model: m, rpm: rpm };
        }
    }
});

const mSiteIdx = mCols.Site !== undefined ? mCols.Site : 6;
result.finalResults.forEach((r, idx) => {
    // Let's find corresponding row in masterRaw
    // Since r contains index or we can match by model and month, let's find the original raw row
    // Wait, result.finalResults has exact 1-to-1 match with rows in masterRaw
    // Actually, we can just run the loop on masterRaw and do standard resolution
});

const relation = {};

masterRaw.forEach((row, idx) => {
    if (idx <= mCols.headerRowIdx) return;
    const mModelIdx = mCols.Model !== undefined ? mCols.Model : 1;
    const mPlIdx = mCols.PL !== undefined ? mCols.PL : 1;
    const mVerIdx = mCols.Ver !== undefined ? mCols.Ver : 7;
    const mSiteIdx = mCols.Site !== undefined ? mCols.Site : 6;
    const mPjtIdx = mCols.Pjt !== undefined ? mCols.Pjt : 4;

    const pName = (row[mPjtIdx] || row[mModelIdx] || '').toString().trim();
    const plCode = (row[mPlIdx] || '').toString().trim();
    const verCode = (row[mVerIdx] || '').toString().trim();
    let mModel = (row[mModelIdx] || '').toString().trim();

    const pNamePrefix = pName.split('-')[0].trim();
    const mKey = getMatchKey(pNamePrefix || mModel);

    let originalSite = (row[mSiteIdx] || '').toString().trim();
    let finalSite = originalSite;
    let foundMeta = metaMap[mKey];
    if (!foundMeta) {
        const possibleKeys = Object.keys(metaMap);
        const bestMatch = possibleKeys.find(k => k.includes(mKey) || mKey.includes(k));
        if (bestMatch) foundMeta = metaMap[bestMatch];
    }

    if (foundMeta) {
        finalSite = foundMeta.site;
    }

    // Standardize Site
    if (plCode === 'I0215001') finalSite = 'LEO';
    if (plCode === 'I0169394' || verCode === '9ACE') finalSite = '지티테크';
    
    if (finalSite === '1842' || finalSite === 1842 || finalSite.includes('성주')) finalSite = '성주';
    else if (finalSite === '1840' || finalSite === 1840 || finalSite.includes('남산')) finalSite = '남산';

    finalSite = finalSite.replace(/^\d+\.\s*/, '').trim();
    if (finalSite.includes('LEO')) finalSite = 'LEO';
    if (finalSite.includes('지티')) finalSite = '지티테크';

    if (!relation[originalSite]) relation[originalSite] = new Set();
    relation[originalSite].add(finalSite);
});

console.log('Original Site Code -> Final Standardized Sites:');
for (const [orig, finals] of Object.entries(relation)) {
    console.log(`Original: "${orig}" -> Finals:`, Array.from(finals));
}
