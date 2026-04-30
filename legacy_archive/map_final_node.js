const fs = require('fs');
const path = require('path');

const refPath = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\mps_all_raw.txt';
const csvPath = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\_FinalList_4650_Complete.csv';
const outPath = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\_FinalList_4650_Complete_Verified.csv';

function norm(s) {
    if (!s) return "";
    return s.toString().toUpperCase().replace(/[^A-Z0-9]/g, '');
}

// 1. Read Reference
const refLines = fs.readFileSync(refPath, 'utf8').split('\n');
const mpsList = [];
refLines.forEach(line => {
    const parts = line.trim().split(/\t/); // Use Tab specifically if possible, or \s+
    let p = parts;
    if (p.length < 3) p = line.trim().split(/\s+/);
    
    if (p.length >= 3 && !isNaN(p[0])) {
        const code = p[1];
        const prod = p[2];
        mpsList.push({ code, prod, norm: norm(prod) });
    }
});
console.log(`Loaded ${mpsList.length} reference items.`);

// 2. Read CSV
let csvText = fs.readFileSync(csvPath, 'utf8');
if (csvText.charCodeAt(0) === 0xFEFF) csvText = csvText.slice(1);

const lines = csvText.split('\n');
const header = lines[0];
const results = [header];

let matchCount = 0;

for (let r = 1; r < lines.length; r++) {
    let line = lines[r].trim();
    if (!line) continue;
    
    // Robust CSV split for "A","B","C"
    const cols = line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/).map(c => c.trim().replace(/^"|"$/g, '').replace(/""/g, '"'));
    
    if (cols.length < 3) {
        results.push(line);
        continue;
    }
    
    const model = cols[2];
    const mNorm = norm(model);
    
    let fC = "", fP = "";
    const variants = [mNorm];
    if (mNorm.startsWith("PUMA")) variants.push(mNorm.substring(4), "P" + mNorm.substring(4));
    if (mNorm.startsWith("LYNX")) variants.push(mNorm.substring(4), "L" + mNorm.substring(4));
    if (mNorm.startsWith("VCF")) variants.push("VF" + mNorm.substring(3));

    for (const v of variants) {
        let short = v.replace(/II/g, '2');
        let short2 = short;
        if (short.length > 4 && short.endsWith('0')) short2 = short.slice(0, -1);

        const match = mpsList.find(m => m.norm === short || m.norm.startsWith(short) || m.norm.startsWith(short2));
        if (match) {
            fC = match.code;
            fP = match.prod;
            matchCount++;
            break;
        }
    }
    
    // Final mapping override for specific critical errors if still failing
    if (mNorm === "PUMA4100B" && !fC) { fC = "ML0278"; fP = "P4100B-F0TP-0-K30"; } 
    
    cols[5] = fC;
    cols[6] = fP;
    
    results.push('"' + cols.join('","') + '"');
}

fs.writeFileSync(outPath, "\ufeff" + results.join('\n'));
console.log(`REBUILT VERIFIED CSV. Matches: ${matchCount} / ${results.length-1}`);
