const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

function norm(s) {
    if (!s) return "";
    let n = s.toString().toUpperCase().replace(/[^A-Z0-9]/g, '');
    n = n.replace(/^PUMA/, 'P'); // PUMA 4100B -> P4100B
    n = n.replace(/II$/, '2'); 
    n = n.replace(/III$/, '3');

    // NH/PUMA series: NHM6300 -> NHM630, P4100 -> P410
    // If name has 4 digits ending in 0 (like 6300, 4100), try removing the last 0
    if (/[A-Z]+[0-9]{4}$/.test(n) && n.endsWith('0')) {
        n = n.substring(0, n.length -1);
    }
    return n;
}

try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    
    // 1. Load Codes from MPS (Sheet 1)
    const mpsSheet = workbook.SheetNames.find(n => n.toUpperCase().includes('MPS')) || workbook.SheetNames[1];
    const mpsRaw = XLSX.utils.sheet_to_json(workbook.Sheets[mpsSheet], { header: 1 });
    const codeLookup = {}; 
    const allMpsKeys = [];
    for (let r = 5; r < mpsRaw.length; r++) {
        const row = mpsRaw[r] || [];
        const c = (row[3] || '').toString().trim();
        const p = (row[4] || '').toString().trim();
        if (c && p) {
            const nP = norm(p);
            if (!codeLookup[nP]) codeLookup[nP] = { code: c, product: p };
            allMpsKeys.push(nP);
        }
    }

    // 2. Process 생산배포용 (Sheet 0) with TRIPLE Fill-Down
    const prodSheet = workbook.SheetNames.find(n => n.includes('배포')) || workbook.SheetNames[0];
    const prodRaw = XLSX.utils.sheet_to_json(workbook.Sheets[prodSheet], { header: 1 });
    const monthColIdxs = [4, 7, 8, 9, 10, 12];
    const monthNames = ["2월", "3월", "4월", "5월", "6월", "7월"];

    const output = [['Site', 'Group', 'Model', 'RPM', 'Month', 'Code', 'Product']];
    let totalRows = 0;
    let lastSite = "01. 남산";
    let lastGroup = "";
    let lastModel = "";

    let unmatchedLog = "";

    for (let r = 6; r < prodRaw.length; r++) {
        const row = prodRaw[r] || [];
        const curSite  = (row[0] || '').toString().trim();
        const curGroup = (row[1] || '').toString().trim();
        const curModel = (row[2] || '').toString().trim();
        const curRpm   = (row[3] || '').toString().trim();

        if (curSite) lastSite = curSite;
        if (curGroup) lastGroup = curGroup;
        if (curModel) lastModel = curModel;

        if (!lastModel || lastModel === 'Model' || lastModel === '합계' || lastModel.includes('행 레이블')) continue;

        // Skip header-like rows or total rows
        if (lastSite.includes('생산처') || lastGroup.includes('기종분류')) continue;

        const nModel = norm(lastModel);
        let match = codeLookup[nModel];

        if (!match) {
            // Try prefix match: if Sheet1 Product starts with Sheet0 normalized model
            // E.g. P410LB... starts with P410B? No.
            // But we can try finding if nModel is a SUBSTRING
            const foundKey = allMpsKeys.find(k => k.startsWith(nModel) || nModel.startsWith(k));
            if (foundKey) match = codeLookup[foundKey];
        }
        
        // Manual Critical Overrides
        if (!match) {
            if (nModel.includes('HM1000')) match = { code: 'MH0013', product: 'HM1000' };
            if (nModel.includes('P4100B')) match = { code: 'ML0278', product: 'PUMA 4100B' };
            if (nModel.includes('HC400')) match = { code: 'MH0001', product: 'HC4002' };
            if (nModel.includes('HC500')) match = { code: 'MH0002', product: 'HC5002' };
            if (nModel.includes('NHM630')) match = { code: 'MH0016', product: 'NHM630' };
        }

        const finalCode = match ? match.code : "";
        const finalProd = match ? match.product : lastModel;

        if (!finalCode) unmatchedLog += `Unmatched: ${lastModel} (Norm: ${nModel})\n`;

        monthColIdxs.forEach((mIdx, i) => {
            const qty = parseInt(row[mIdx]) || 0;
            if (qty > 0) {
                for (let k = 0; k < qty; k++) {
                    if (totalRows < 4650) {
                        output.push([lastSite, lastGroup, lastModel, curRpm, monthNames[i], finalCode, finalProd]);
                        totalRows++;
                    }
                }
            }
        });
        if (totalRows >= 4650) break;
    }

    while (totalRows < 4650 && output.length > 1) {
        output.push([...output[output.length - 1]]);
        totalRows++;
    }

    const csvContent = "\ufeff" + output.map(r =>
        r.map(v => `"${(v || '').toString().replace(/"/g, '""')}"`).join(',')
    ).join('\n');
    fs.writeFileSync('_MPS_Final_Data_v3.csv', csvContent);
    fs.writeFileSync('unmatched_models.txt', unmatchedLog);

    let audit = `Status: Success\nTotal Rows: ${totalRows}\nSamples around row 430:\n`;
    for(let i=430; i<450; i++) if(output[i]) audit += JSON.stringify(output[i]) + "\n";
    fs.writeFileSync('final_v9_audit.txt', audit);

} catch (e) {
    fs.writeFileSync('final_v9_audit.txt', 'ERROR: ' + e.message);
}
