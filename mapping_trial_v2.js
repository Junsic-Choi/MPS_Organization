const XLSX = require('xlsx');
const fs = require('fs');

function parseModel(m) {
    if (!m) return null;
    let s = m.toUpperCase().replace(/\s+/g, '');
    let prefix = s.match(/^[A-Z]+/)?.[0] || "";
    let rest = s.substring(prefix.length);
    let digits = rest.match(/^[0-9]+/)?.[0] || "";
    let suffix = rest.substring(digits.length);
    
    // Normalization rules for prefixes
    let normPrefix = prefix;
    if (prefix === "PUMA") normPrefix = "P";
    if (prefix === "LYNX") normPrefix = "L";

    // Normalization rules for digits (4-digit to 3-digit compression)
    let altDigits = digits;
    if (digits.length === 4 && digits.endsWith('0')) {
        altDigits = digits.substring(0, 3);
    }

    return { 
        raw: m, 
        clean: prefix + digits + suffix,
        norm: normPrefix + digits + suffix,
        alt: normPrefix + altDigits + suffix,
        core: digits,
        altCore: altDigits,
        prefix: normPrefix,
        suffix: suffix
    };
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
            mpsEntries.push({ code, product, p: parseModel(product) });
        }
    }

    const prodSheet = workbook.SheetNames.find(n => n.includes('배포')) || workbook.SheetNames[0];
    const prodRaw = XLSX.utils.sheet_to_json(workbook.Sheets[prodSheet], { header: 1 });
    
    const sheet0Models = new Set();
    let lastModel = "";
    for (let r = 6; r < prodRaw.length; r++) {
        const m = (prodRaw[r][2] || '').toString().trim();
        if (m) lastModel = m;
        if (lastModel && lastModel !== 'Model') sheet0Models.add(lastModel);
    }

    let report = "--- Advanced Mapping Report ---\n";
    let matchedCount = 0;

    [...sheet0Models].forEach(m => {
        const p0 = parseModel(m);
        if (!p0) return;

        // Multi-level matching
        let match = mpsEntries.find(e => 
            e.p.clean === p0.clean || 
            e.p.norm === p0.norm || 
            e.p.alt === p0.alt ||
            e.product.startsWith(p0.norm) ||
            e.product.startsWith(p0.alt)
        );

        if (!match) {
            // Suffix check: e.g. P410LB... vs PUMA 4100LB
            match = mpsEntries.find(e => 
                e.p.prefix === p0.prefix && 
                (e.p.core === p0.altCore || e.p.core === p0.core) &&
                (e.p.suffix.startsWith(p0.suffix) || p0.suffix.startsWith(e.p.suffix))
            );
        }

        if (match) {
            matchedCount++;
            report += `[OK] ${m} -> ${match.code} (${match.product})\n`;
        } else {
            report += `[FAIL] ${m} (Norm: ${p0.norm}, Alt: ${p0.alt})\n`;
        }
    });

    report += `\nMatched: ${matchedCount} / ${sheet0Models.size}\n`;
    fs.writeFileSync('mapping_trial_v2.txt', report);

} catch (e) {
    fs.writeFileSync('mapping_trial_v2.txt', 'ERROR: ' + e.message);
}
