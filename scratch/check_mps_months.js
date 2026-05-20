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
console.log('mCols detected:', mCols);

// Let's print out rows where months are defined
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
console.log(`monthRowIdx=${monthRowIdx}, typeRowIdx=${typeRowIdx}`);

// Build site lookup from 생산배포용
const prodWs = wb.Sheets['생산배포용'];
const prodRaw = XLSX.utils.sheet_to_json(prodWs, { header: 1 });
const prodKeywords = {
    Model: ['기종', 'Model'],
    Group: ['기종분류', '시리즈', 'Series', '그룹'],
    Site: ['생산처', '공장', '사업장', 'Site'],
    RPM: ['RPM', '주축', 'Spindle']
};
const pCols = findCols(prodRaw, prodKeywords, 50);
const prodHeaderIdx = pCols.headerRowIdx !== undefined ? pCols.headerRowIdx : -1;

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

// Let's list the month columns detected and sum Seongju for each
const masterMonthCols = [];
const monthRow = raw[monthRowIdx];
const typeRow = raw[typeRowIdx];

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

const monthSums = {};
raw.forEach((row, idx) => {
    if (idx <= typeRowIdx) return;
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

    if (plCode === 'I0215001') finalSite = 'LEO';
    if (plCode === 'I0169394' || verCode === '9ACE') finalSite = '지티테크';
    
    if (finalSite === '1842' || finalSite === 1842 || finalSite.includes('성주')) finalSite = '성주';
    else if (finalSite === '1840' || finalSite === 1840 || finalSite.includes('남산')) finalSite = '남산';

    finalSite = finalSite.replace(/^\d+\.\s*/, '').trim();
    if (finalSite.includes('LEO')) finalSite = 'LEO';
    if (finalSite.includes('지티')) finalSite = '지티테크';

    if (finalSite === '성주' && (pName || mModel)) {
        masterMonthCols.forEach(mCol => {
            const q = parseInt(row[mCol.col]) || 0;
            if (q > 0) {
                monthSums[mCol.name] = (monthSums[mCol.name] || 0) + q;
            }
        });
    }
});

console.log('Seongju Monthly sums in MPS sheet:');
console.log(monthSums);
