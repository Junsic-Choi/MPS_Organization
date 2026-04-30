const fs = require('fs');
const XLSX = require('xlsx');

function solve() {
    const buf = fs.readFileSync('MPS2603-1.xlsx');
    const wb = XLSX.read(buf, {type:'buffer'});
    const data = XLSX.utils.sheet_to_json(wb.Sheets['MPS'], {header:1});

    const monthCols = [];
    const headerRow = data[4] || data[5] || [];
    headerRow.forEach((c, i) => { if(/\d+월/.test((c||'').toString())) monthCols.push({name:c, col:i}); });
    
    if(monthCols.length === 0) {
        ['2월','3월','4월','5월','6월','7월'].forEach((m, i) => monthCols.push({name:m, col: [8,12,17,22,28,34][i]}));
    }

    let p41Details = {};

    for(let r=5; r<data.length; r++) {
        const row = data[r] || [];
        const prod = (row[4]||'').toString().toUpperCase();
        const siteCode = (row[6]||'').toString().trim();
        const verCode = (row[7]||'').toString().trim();
        
        if (prod.includes('P41')) {
            monthCols.forEach(m => {
                let Q = parseInt(row[m.col]) || 0;
                if (Q > 0) {
                    const key = `${siteCode} | ${verCode} | ${prod}`;
                    p41Details[key] = (p41Details[key] || 0) + Q;
                }
            });
        }
    }

    console.log('--- P41 Detailed Breakdown (Site | Version | Prod) ---');
    console.log(JSON.stringify(p41Details, null, 2));
}

solve();
