const XLSX = require('xlsx');
const fs = require('fs');

function clean(s) {
    if (!s) return "";
    return s.toString().toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function smartNorm(s) {
    let n = clean(s);
    n = n.replace(/^PUMA/, 'P');
    n = n.replace(/^LYNX/, 'L');
    // If has 4 digits ending in 0, try 3 digits too
    // E.g. P4100B -> P410B
    return n;
}

try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    const mpsSheet = workbook.SheetNames.find(n => n.toUpperCase().includes('MPS')) || workbook.SheetNames[1];
    const mpsRaw = XLSX.utils.sheet_to_json(workbook.Sheets[mpsSheet], { header: 1 });
    
    const mpsEntries = [];
    for (let r = 5; r < mpsRaw.length; r++) {
        const row = mpsRaw[r] || [];
        const code = (row[3] || '').toString().trim();
        const product = (row[4] || '').toString().trim();
        if (code && product) {
            mpsEntries.push({ code, product, clean: clean(product), smart: smartNorm(product) });
        }
    }

    const prodSheet = workbook.SheetNames.find(n => n.includes('배포')) || workbook.SheetNames[0];
    const prodRaw = XLSX.utils.sheet_to_json(workbook.Sheets[prodSheet], { header: 1 });
    
    let lastModel = "";
    const uniqueSheet0Models = new Set();
    for (let r = 6; r < prodRaw.length; r++) {
        const m = (prodRaw[r][2] || '').toString().trim();
        if (m) lastModel = m;
        if (lastModel && lastModel !== 'Model') uniqueSheet0Models.add(lastModel);
    }

    let report = "--- Mapping Report ---\n";
    let matchedCount = 0;

    [...uniqueSheet0Models].forEach(m => {
        const sc = clean(m);
        const ss = smartNorm(m);
        
        let match = mpsEntries.find(e => e.clean === sc || e.smart === ss);
        if (!match) {
            // Try number compression match: P4100B vs P410B
            const compressed = ss.replace(/([A-Z]+)[0-9]00([A-Z]*)/, '$1$2'); // Very rough
            match = mpsEntries.find(e => e.smart.includes(ss) || ss.includes(e.smart));
        }

        if (match) {
            matchedCount++;
            report += `[OK] ${m} -> ${match.code} (${match.product})\n`;
        } else {
            report += `[FAIL] ${m} (Clean: ${sc}, Smart: ${ss})\n`;
        }
    });

    report += `\nMatched: ${matchedCount} / ${uniqueSheet0Models.size}\n`;
    fs.writeFileSync('mapping_trial.txt', report);

} catch (e) {
    fs.writeFileSync('mapping_trial.txt', 'ERROR: ' + e.message);
}
