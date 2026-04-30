const fs = require('fs');
const XLSX = require('xlsx');

function solve() {
    const buf = fs.readFileSync('MPS2603-1.xlsx');
    const wb = XLSX.read(buf, {type:'buffer'});
    const sheetNames = wb.SheetNames;
    const mpsWsName = sheetNames.find(n => n.toUpperCase() === 'MPS') || sheetNames[1];
    const ws = wb.Sheets[mpsWsName];
    const data = XLSX.utils.sheet_to_json(ws, {header:1});

    const mpsSiteMap = {
        "1840":"01. 남산", "1841":"01. 남산", "1":"01. 남산", "10":"01. 남산(보조)",
        "연암FA_OEM_1840": "연암FA (01. 남산)",
        "쎈텍_OEM_1840": "쎈텍 (01. 남산)",
        "씨이에스테크_OEM_1840": "씨이에스테크 (01. 남산)"
    };

    const monthCols = [];
    const headerRow = data[4] || data[5] || [];
    headerRow.forEach((c, i) => { if(/\d+월/.test((c||'').toString())) monthCols.push({name:c, col:i}); });
    
    if(monthCols.length === 0) {
        // Fallback same as extractor.js
        ['2월','3월','4월','5월','6월','7월'].forEach((m, i) => monthCols.push({name:m, col: [8,12,17,22,28,34][i]}));
    }

    let namsanDetails = {};
    let monthlyTotal = {};
    let runningSite = '';

    for(let r=5; r<data.length; r++) {
        const row = data[r] || [];
        const siteCode = (row[6] || '').toString().trim();
        if(siteCode) runningSite = mpsSiteMap[siteCode] || siteCode;

        if (runningSite && runningSite.includes('남산')) {
            monthCols.forEach(m => {
                let Q = parseInt(row[m.col]) || 0;
                if (Q > 1000) Q = 1;
                if (Q > 0) {
                    if(!monthlyTotal[m.name]) monthlyTotal[m.name] = 0;
                    monthlyTotal[m.name] += Q;
                    
                    const detailKey = `${runningSite} (${m.name})`;
                    namsanDetails[detailKey] = (namsanDetails[detailKey] || 0) + Q;
                }
            });
        }
    }

    console.log('--- Ground Truth Namsan Analysis ---');
    console.log('Breakdown:', JSON.stringify(namsanDetails, null, 2));
    console.log('Monthly Totals:', JSON.stringify(monthlyTotal, null, 2));
    let grandTotal = Object.values(monthlyTotal).reduce((a, b) => a + b, 0);
    console.log('Grand Total for Namsan:', grandTotal);
}

try { solve(); } catch(e) { console.error(e); }
