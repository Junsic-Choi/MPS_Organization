const XLSX = require('xlsx');
const wb = XLSX.readFile('MPS2604-1.xlsx');
const sheetNames = wb.SheetNames;
const masterWsName = sheetNames.find(n => n.toUpperCase() === 'MPS') || 'MPS';
const masterWs = wb.Sheets[masterWsName];

if (masterWs) {
    const raw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });
    console.log(`Searching for TW Series in Master Plan [${masterWsName}]...`);
    
    // Find headers
    let colProduct = -1, colCode = -1;
    for(let r=0; r<20; r++) {
        const row = raw[r] || [];
        row.forEach((cell, idx) => {
            const s = (cell || '').toString().trim().toUpperCase();
            if (s.includes('품명') || s.includes('PRODUCT') || s === 'MODEL') colProduct = idx;
            if (s.includes('모델') || s.includes('CODE') || s === '품번') colCode = idx;
        });
        if (colProduct !== -1 && colCode !== -1) break;
    }

    let headerRow = -1;
    for(let r=0; r<10; r++) {
        if ((raw[r] || []).some(c => (c || '').toString().includes('월'))) {
            headerRow = r;
            break;
        }
    }

    raw.forEach((row, i) => {
        const pName = (row[colProduct] || '').toString().trim();
        const mCode = (row[colCode] || '').toString().trim();
        if (pName.includes('TW')) {
            const site = (row[6] || '').toString().trim(); // Plant? No, Col 7 is '9ASW'
            const siteVer = (row[7] || '').toString().trim();
            console.log(`Row ${i}: ${pName} SiteVer[${siteVer}]`);
            const months = ['26.3월 실적', '4월 예상', '5월 예상', '6월 예상', '7월 예상', '8월 예상'];
            months.forEach(m => {
                const idx = (raw[headerRow] || []).findIndex(c => (c || '').toString().includes(m.split(' ')[0]));
                if (idx !== -1) {
                    const q = row[idx] || 0;
                    if (q > 0) console.log(`  ${m}: ${q}`);
                }
            });
        }
    });
}
