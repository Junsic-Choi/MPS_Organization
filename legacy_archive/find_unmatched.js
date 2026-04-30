const XLSX = require('xlsx');
const fs = require('fs');

function getSeq(s) {
    if (!s) return "";
    let n = s.toString().toUpperCase().replace(/[^A-Z1-9]/g, '');
    if (n.startsWith('PUMA')) n = 'P' + n.substring(4);
    if (n.startsWith('LYNX')) n = 'L' + n.substring(4);
    return n;
}

function isSubsequence(sub, main) {
    if (!sub || !main) return false;
    let subIdx = 0;
    for (let i = 0; i < main.length && subIdx < sub.length; i++) {
        if (main[i] === sub[subIdx]) subIdx++;
    }
    return subIdx === sub.length;
}

try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    const mpsSheet = workbook.SheetNames.find(n => n.toUpperCase().includes('MPS')) || workbook.SheetNames[1];
    const mpsRaw = XLSX.utils.sheet_to_json(workbook.Sheets[mpsSheet], { header: 1 });
    const codeLookup = [];
    for (let r = 5; r < mpsRaw.length; r++) {
        const row = mpsRaw[r] || [];
        const code = (row[3] || '').toString().trim();
        const prodFull = (row[4] || '').toString().trim();
        if (code && prodFull) {
            const prodModel = prodFull.split('-')[0].trim();
            codeLookup.push({ code, product: prodFull, seq: getSeq(prodModel), clean: prodModel.toUpperCase().replace(/[^A-Z0-9]/g, '') });
        }
    }

    const prodSheet = workbook.SheetNames.find(n => n.includes('배포')) || workbook.SheetNames[0];
    const prodRaw = XLSX.utils.sheet_to_json(workbook.Sheets[prodSheet], { header: 1 });
    
    let lastModel = "";
    const unmatched = new Set();
    for (let r = 6; r < prodRaw.length; r++) {
        const m = (prodRaw[r][2] || '').toString().trim();
        if (m) lastModel = m;
        if (!lastModel || lastModel === 'Model' || lastModel === '합계') continue;

        const mySeq = getSeq(lastModel);
        const myClean = lastModel.toUpperCase().replace(/[^A-Z0-9]/g, '');
        let match = codeLookup.find(e => 
            (e.seq.length > 2 && (isSubsequence(e.seq, mySeq) || isSubsequence(mySeq, e.seq))) ||
            (e.clean.length > 2 && (e.clean.includes(myClean) || myClean.includes(e.clean)))
        );

        if (!match) unmatched.add(lastModel);
    }

    fs.writeFileSync('still_unmatched.txt', [...unmatched].join('\n'));

} catch (e) {
    fs.writeFileSync('still_unmatched.txt', 'ERROR: ' + e.message);
}
