const fs = require('fs');
const XLSX = require('xlsx');

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

const wb = XLSX.readFile('MPS2605-1.xlsx');
const prodWsName = '생산배포용';
const masterWsName = 'MPS';

const prodWs = wb.Sheets[prodWsName];
const masterWs = wb.Sheets[masterWsName];
const prodRaw = XLSX.utils.sheet_to_json(prodWs, { header: 1 });
const masterRaw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });

const prodKeywords = {
    Model: ['기종', 'Model'],
    Group: ['기종분류', '시리즈', 'Series', '그룹'],
    Site: ['생산처', '공장', '사업장', 'Site'],
    RPM: ['RPM', '주축', 'Spindle']
};
const pCols = findCols(prodRaw, prodKeywords, 50);
const prodHeaderIdx = pCols.headerRowIdx !== undefined ? pCols.headerRowIdx : -1;

console.log('--- AUDITING SEONGJU IN 생산배포용 ---');
let lastSite = '';
let prodSeongjuSum = 0;
let rowSeongjuData = [];

// Columns in Row 2 of 생산배포용:
// Column 0: 생산처
// Column 1: 기종분류
// Column 2: 기종
// Column 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14 ...: monthly data
// Let's see how much quantity is in the '합계' row or daily rows for '성주'
prodRaw.forEach((row, idx) => {
    if (idx <= prodHeaderIdx) return;
    const s = (row[pCols.Site] || '').toString().trim();
    if (s) lastSite = s;
    
    // We check if lastSite is '성주'
    if (lastSite.includes('성주')) {
        const m = (row[pCols.Model] || '').toString().trim();
        if (m) {
            // Let's sum quantities for months (columns 7, 8, 9, 11, 13) in this row
            // Row 2 headers: [..., '합계 : 생산', '합계 : 판매P', '합계 : 재고', '합계 : 생산2', '합계 : 생산3', '합계 : 생산4', '합계 : 생산5', '합계 : 자재', '합계 : 생산6', '합계 : 자재2', '합계 : 생산7']
            // Wait, let's look at the columns for monthly production in Row 2:
            // 5월: Column 7 ('합계 : 생산2')
            // 6월: Column 8 ('합계 : 생산3')
            // 7월: Column 9 ('합계 : 생산4')
            // 8월: Column 10 ('합계 : 생산5')
            // 9월: Column 12 ('합계 : 생산6')
            // 10월: Column 14 ('합계 : 생산7')
            // Let's look at row[7], row[8], row[9], row[10], row[12] (excluding 10월, which is column 14)
            const q5 = parseInt(row[7]) || 0;
            const q6 = parseInt(row[8]) || 0;
            const q7 = parseInt(row[9]) || 0;
            const q8 = parseInt(row[10]) || 0;
            const q9 = parseInt(row[12]) || 0;
            const rowSum = q5 + q6 + q7 + q8 + q9;
            prodSeongjuSum += rowSum;
            if (rowSum > 0) {
                rowSeongjuData.push({ model: m, group: row[pCols.Group], q5, q6, q7, q8, q9, sum: rowSum });
            }
        }
    }
});
console.log('Total Seongju Sum in 생산배포용 (excluding 10월):', prodSeongjuSum);
console.log('Seongju models with Qty > 0 in 생산배포용:', rowSeongjuData);

// Now let's audit MPS sheet for rows that match these keys or have finalSite = '성주'
console.log('\n--- AUDITING SEONGJU IN MPS SHEET ---');
const masterKeywords = {
    Model: ['기종', 'Model'],
    Group: ['그룹', 'Group', 'Series'],
    Site: ['사업장', '공장', 'Site'],
    PL: ['PL', '제품군'],
    Ver: ['Ver', '버전'],
    Pjt: ['PJT', '프로젝트', 'Product Name', 'Product']
};
const mCols = findCols(masterRaw, masterKeywords, 50);

let monthRowIdx = -1;
let typeRowIdx = -1;
for (let r = 0; r < Math.min(50, masterRaw.length); r++) {
    const rowStr = (masterRaw[r] || []).join('|');
    if (rowStr.includes('월') && monthRowIdx === -1) monthRowIdx = r;
    if (rowStr.includes('생산') && rowStr.includes('판매') && r > monthRowIdx) {
        typeRowIdx = r;
        break;
    }
}

const masterMonthCols = [];
if (monthRowIdx !== -1 && typeRowIdx !== -1) {
    const monthRow = masterRaw[monthRowIdx];
    const typeRow = masterRaw[typeRowIdx];
    typeRow.forEach((cell, idx) => {
        const type = (cell || '').toString().trim();
        if (type === '생산') {
            for (let c = idx; c >= 0; c--) {
                const mNum = String(monthRow[c] || '').trim();
                if (mNum) {
                    const mMatch = mNum.match(/(\d+)/);
                    if (mMatch) {
                        masterMonthCols.push({ name: mMatch[1] + '월', col: idx });
                        break;
                    }
                }
            }
        }
    });
}

// Build metaMap first
const metaMap = {};
lastSite = '';
let lastGroup = '';
prodRaw.forEach((row, idx) => {
    if (idx <= prodHeaderIdx) return;
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

let mpsSeongjuSum = 0;
let mpsSeongjuRows = [];

masterRaw.forEach((row, idx) => {
    if (idx <= typeRowIdx) return;
    const mModelIdx = mCols.Model !== undefined ? mCols.Model : 1;
    const mGroupIdx = mCols.Group !== undefined ? mCols.Group : 2;
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

    let finalSite = (row[mSiteIdx] || '').toString().trim();
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

    if (finalSite === '성주') {
        let rowSum = 0;
        masterMonthCols.forEach(mCol => {
            const mMatch = mCol.name.match(/(\d{1,2})/);
            if (mMatch && parseInt(mMatch[1]) >= 10) return;
            const q = parseInt(row[mCol.col]) || 0;
            rowSum += q;
        });
        if (rowSum > 0) {
            mpsSeongjuSum += rowSum;
            mpsSeongjuRows.push({ product: pName, model: mModel, code: verCode || plCode, mpsSite: row[mSiteIdx], finalSite, sum: rowSum });
        }
    }
});

console.log('Total Seongju Sum in extracted MPS (excluding 10월):', mpsSeongjuSum);
console.log('Seongju rows in extracted MPS (first 15):', mpsSeongjuRows.slice(0, 15));
console.log('Total Seongju rows count in MPS:', mpsSeongjuRows.length);
