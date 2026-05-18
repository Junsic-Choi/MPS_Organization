const XLSX = require('xlsx');

function sumQuota(filename, site, key, monthIdx) {
    const wb = XLSX.readFile(filename);
    const ws = wb.Sheets['MPS'];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
    
    // I need the getMatchKey logic here too
    const { getMatchKey } = require('../extractor');
    
    let total = 0;
    data.forEach((row, idx) => {
        if (idx < 5) return;
        const plCode = (row[1] || '').toString();
        let s = (row[6] || '').toString();
        // Simple mapping for 성주
        if (s === '1842' || plCode === 'I0206954') s = '성주';
        
        if (s === site) {
            const mCode = (row[5] || row[3] || '').toString();
            const pName = (row[4] || row[2] || '').toString();
            if (pName || mCode) {
                const modelPart = (pName || mCode).split('-')[0].trim();
                const k = getMatchKey(modelPart);
                if (k === key) {
                    const q = parseInt(row[monthIdx]) || 0;
                    if (q > 0) {
                        console.log(`Row ${idx+1}: ${modelPart} Qty ${q}`);
                        total += q;
                    }
                }
            }
        }
    });
    console.log(`\nTotal Quota for ${site} ${key} in Month ${monthIdx}: ${total}`);
}

sumQuota('MPS2605-1.xlsx', '성주', 'VTR121', 17); // 17 is 6월
