const fs = require('fs');
const refPath = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\mps_all_raw.txt';
function norm(s) { return s ? s.toString().toUpperCase().replace(/[^A-Z0-9]/g, '') : ""; }

let log = "";
function logger(msg) { log += msg + "\n"; }

try {
    const refLines = fs.readFileSync(refPath, 'utf8').split('\n');
    const mpsList = [];
    refLines.forEach(line => {
        const p = line.trim().split(/\s+/);
        if (p.length >= 3 && !isNaN(p[0])) {
            mpsList.push({ code: p[1], prod: p[2], norm: norm(p[2]) });
        }
    });

    const testModel = "HM1000";
    const mNorm = norm(testModel);
    logger(`DEBUG: testModel='${testModel}', mNorm='${mNorm}'`);

    const match = mpsList.find(m => {
        const ok = m.norm.startsWith(mNorm);
        if (m.prod.includes("HM1000")) logger(`  Checking Ref: Code=${m.code}, Prod=${m.prod}, Norm=${m.norm}, Match=${ok}`);
        return ok;
    });

    if (match) logger(`SUCCESS: Found ${match.code} for ${testModel}`);
    else logger(`FAILURE: No match for ${testModel}`);
    
    fs.writeFileSync('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\debug_mapping.log', log);
} catch (e) {
    fs.writeFileSync('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\debug_mapping.log', "CRITICAL: " + e.message);
}
