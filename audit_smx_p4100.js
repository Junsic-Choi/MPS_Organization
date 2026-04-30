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

    let smxRecords = [];
    let p4100Records = [];
    let runningSite = '';

    for(let r=5; r<data.length; r++) {
        const row = data[r] || [];
        if (row[6]) runningSite = row[6].toString().trim();
        const prod = (row[4]||'').toString();
        
        if (prod.toUpperCase().includes('SMX')) {
            monthCols.forEach(m => {
                let Q = parseInt(row[m.col]) || 0;
                if (Q > 0) smxRecords.push({ Month: m.name, SiteCode: runningSite, Prod: prod, Q: Q });
            });
        }
        if (prod.toUpperCase().includes('P41')) {
            monthCols.forEach(m => {
                let Q = parseInt(row[m.col]) || 0;
                if (Q > 0) p4100Records.push({ Month: m.name, SiteCode: runningSite, Prod: prod, Q: Q });
            });
        }
    }

    const groups = {
        "1840": "01. 남산", "1841": "01. 남산", "1": "01. 남산", "10": "01. 남산",
        "1842": "02. 성주", "2": "02. 성주",
        "1848": "03. 창원", "3": "03. 창원"
    };

    console.log('--- SMX Summary by Mapped Site ---');
    let smxBySite = {};
    smxRecords.forEach(r => {
        let name = groups[r.SiteCode] || r.SiteCode || 'Unknown';
        smxBySite[name] = (smxBySite[name] || 0) + r.Q;
    });
    console.log(JSON.stringify(smxBySite, null, 2));

    console.log('\n--- P41 Summary by Site Code ---');
    let p41BySite = {};
    p4100Records.forEach(r => {
        p41BySite[r.SiteCode] = (p41BySite[r.SiteCode] || 0) + r.Q;
    });
    console.log(JSON.stringify(p41BySite, null, 2));
}

solve();
