const fs = require('fs');
const path = require('path');

const csvHistory = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\일반비_MPS2603-1(생산배포용)_FinalList.csv';
const jsonSite = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\site_data.json';
const outMap = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\master_map.json';

const masterMap = {};

// 1. Extract from CSV History
try {
    if (fs.existsSync(csvHistory)) {
        const data = fs.readFileSync(csvHistory, 'utf8');
        const lines = data.split('\n');
        for (let i = 1; i < lines.length; i++) {
            const line = lines[i].trim();
            if (!line) continue;
            const columns = line.split('","').map(c => c.replace(/"/g, ''));
            if (columns.length >= 8) {
                const model = columns[2];
                const code = columns[6];
                const product = columns[7];
                if (model && code && product) {
                    masterMap[model] = { code, product };
                }
            }
        }
    }
} catch (err) {
    console.error("CSV History Error:", err.message);
}

// 2. Extract from JSON (try to find anything useful)
try {
    if (fs.existsSync(jsonSite)) {
        // Try reading with multiple encodings or just as-is and look for patterns
        const raw = fs.readFileSync(jsonSite);
        // Maybe it's UTF-16? Or EUC-KR?
        // Let's assume it's roughly readable as utf8 for patterns
        const data = raw.toString('utf8');
        // Look for things like "Model":"...", "Code":"..." if they exist
        // The previous view showed: {"Plant":"1840","Prod. Ver":"0ACE","Prod. Ver Description":"..."}
        // Let's map "Prod. Ver Description" to "Prod. Ver" if they look like models
        try {
            const json = JSON.parse(data.replace(/^\uFEFF/, '')); // Handle BOM
            json.forEach(item => {
                const desc = item["Prod. Ver Description"];
                const ver = item["Prod. Ver"];
                if (desc && ver && desc.length > 2) {
                    // If desc contains common model prefixes
                    if (/^[A-Z]/.test(desc)) {
                        // masterMap[desc] = { code: ver, product: desc };
                    }
                }
            });
        } catch (e) {
            console.error("JSON Parse Error:", e.message);
        }
    }
} catch (err) {
    console.error("JSON Error:", err.message);
}

fs.writeFileSync(outMap, JSON.stringify(masterMap, null, 2), 'utf8');
console.log(`Extracted ${Object.keys(masterMap).length} mapping pairs.`);
