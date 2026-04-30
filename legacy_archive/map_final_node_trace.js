const fs = require('fs');
const refPath = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\mps_all_raw.txt';
const csvPath = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\_FinalList_4650_Complete.csv';
const outPath = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\_FinalList_4650_Complete_Verified.csv';
const logPath = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\trace_mapping.log';

function norm(s) { return s ? s.toString().toUpperCase().replace(/[^A-Z0-9]/g, '') : ""; }

let logTrace = "";
function logger(msg) { logTrace += msg + "\n"; }

try {
    // 1. Load Reference
    const refLines = fs.readFileSync(refPath, 'utf8').split('\n');
    const mpsList = [];
    refLines.forEach(line => {
        const p = line.trim().split(/\s+/);
        if (p.length >= 3 && !isNaN(p[0])) {
            mpsList.push({ code: p[1], prod: p[2], norm: norm(p[2]) });
        }
    });
    logger(`Loaded ${mpsList.length} references.`);

    // 2. Process CSV
    let csvText = fs.readFileSync(csvPath, 'utf8');
    if (csvText.charCodeAt(0) === 0xFEFF) csvText = csvText.slice(1);
    const lines = csvText.split('\n');
    const results = [lines[0]];

    for (let r = 1; r < lines.length; r++) {
        let line = lines[r].trim();
        if (!line) continue;
        
        // Split by ","
        const cols = line.split('","').map(c => c.replace(/"/g, ''));
        if (cols.length < 3) {
            results.push(line);
            continue;
        }

        const model = cols[2];
        const mNorm = norm(model);
        let fC = "", fP = "";

        const match = mpsList.find(m => m.norm.startsWith(mNorm));
        if (match) {
            fC = match.code;
            fP = match.prod;
        }

        if (r <= 20) {
            logger(`Row ${r}: model='${model}', mNorm='${mNorm}', FoundCode='${fC}'`);
        }

        cols[5] = fC;
        cols[6] = fP;
        results.push('"' + cols.join('","') + '"');
    }

    fs.writeFileSync(outPath, "\ufeff" + results.join('\n'));
    fs.writeFileSync(logPath, logTrace);
} catch (e) {
    fs.writeFileSync(logPath, "ERROR: " + e.stack);
}
