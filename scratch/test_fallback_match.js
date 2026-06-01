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
    if (n.startsWith('DVF')) {
        let m = n.match(/DVF(\d)/);
        if (m) return 'DVF' + m[1];
    }
    if (n.startsWith('VCF') || n.startsWith('VF')) {
        let m = n.match(/(?:VCF|VF)(\d)/);
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
    if (key === 'D755' && n.includes('L')) {
        return 'D755L';
    }
    key = key.replace(/[A-Z]+$/, '');
    if (key.startsWith('DC')) key = key.replace(/[A-Z]$/, '');
    return key;
}

const wb = XLSX.readFile('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\MPS2605-2.xlsx');
const prodWs = wb.Sheets['생산배포용'];
const prodRaw = XLSX.utils.sheet_to_json(prodWs, { header: 1 });

const pCols = { Site: 0, Group: 1, Model: 2, RPM: 3 };
const prodHeaderIdx = 2;

// Build Meta Map with relaxed condition (rpm is optional)
const metaMap = {};
let lastMetaSite = '', lastMetaGroup = '', lastMetaModel = '';
prodRaw.forEach((row, idx) => {
    if (idx <= prodHeaderIdx) return;
    const s = (row[pCols.Site] || '').toString().trim();
    const g = (row[pCols.Group] || '').toString().trim();
    const m = (row[pCols.Model] || '').toString().trim();
    const rpm = (row[pCols.RPM] || '').toString().trim();
    
    if (s) lastMetaSite = s;
    if (g) lastMetaGroup = g;
    if (m) lastMetaModel = m;
    
    if (lastMetaModel) {
        const mKey = getMatchKey(lastMetaModel);
        if (!metaMap[mKey]) {
            metaMap[mKey] = [];
        }
        const exists = metaMap[mKey].some(item => item.site === lastMetaSite && item.model === lastMetaModel && item.rpm === rpm);
        if (!exists) {
            metaMap[mKey].push({ 
                site: lastMetaSite, 
                group: lastMetaGroup, 
                model: lastMetaModel, 
                rpm: rpm
            });
        }
    }
});

// Test matching logic with fallback
const testKeys = ['DC3710F', 'DC378F2', 'DC428F', 'DC421F2', 'DCM315H', 'DCM316H'];
testKeys.forEach(tk => {
    const mKey = getMatchKey(tk);
    let foundMetaList = metaMap[mKey];
    let isFallback = false;
    if (!foundMetaList || foundMetaList.length === 0) {
        const possibleKeys = Object.keys(metaMap);
        const bestMatch = possibleKeys.find(k => {
            if (k.length <= 1 || mKey.length <= 1) return false;
            return k.startsWith(mKey) || mKey.startsWith(k);
        });
        if (bestMatch) {
            foundMetaList = metaMap[bestMatch];
            isFallback = true;
        }
    }

    console.log(`\nMatch for "${tk}" (mKey: "${mKey}"):`);
    if (foundMetaList) {
        foundMetaList.forEach(match => {
            console.log(`  -> [${isFallback ? 'FALLBACK' : 'EXACT'}] Model="${match.model}", Group="${match.group}", Site="${match.site}", RPM="${match.rpm}"`);
        });
    } else {
        console.log(`  -> No matches found`);
    }
});
