const express = require('express');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const cors = require('cors');

const app = express();
const PORT = 8890;
const FILE_PATH = path.join(__dirname, 'MPS2603-1.xlsx');

app.use(cors());
app.use(express.static(__dirname));

const normCache = new Map();
function norm(s) {
    if (!s) return "";
    const raw = s.toString().toUpperCase().trim();
    if (normCache.has(raw)) return normCache.get(raw);
    let res = raw;
    res = res.replace(/ III/g, '3').replace(/ II/g, '2').replace(/ I/g, '1');
    res = res.replace(/III/g, '3').replace(/II/g, '2');
    if (res.startsWith('DCM')) res = 'DC' + res.substring(3);
    if (res.startsWith('PUMA')) res = 'P' + res.substring(4);
    if (res.startsWith('LYNX')) res = res.substring(4);
    const final = res.replace(/[^A-Z0-9]/g, '');
    normCache.set(raw, final);
    return final;
}

function getBase(prod) {
    if (!prod) return "";
    return norm(prod.split('-')[0]);
}

function isSub(sub, main) {
    if (!sub || !main) return false;
    let s = 0;
    for (let i = 0; i < main.length && s < sub.length; i++) {
        if (main[i] === sub[s]) s++;
    }
    return s === sub.length;
}

function runExtraction() {
    console.log(`Extraction Start: ${new Date().toISOString()}`);
    const tempWB = XLSX.readFile(FILE_PATH, { bookSheets: true });
    const mpsName = tempWB.SheetNames.find(n => n.toUpperCase().includes('MPS'));
    const prodName = tempWB.SheetNames.find(n => n.includes('배포'));
    const workbook = XLSX.readFile(FILE_PATH, { sheets: [mpsName, prodName] });

    const mpsRaw = XLSX.utils.sheet_to_json(workbook.Sheets[mpsName], { header: 1 });
    const prodRaw = XLSX.utils.sheet_to_json(workbook.Sheets[prodName], { header: 1 });
    
    const monthNames = ["2월", "3월", "4월", "5월", "6월", "7월"];
    const mpsMonthIdxs = [8, 12, 17, 22, 28, 34];
    const mpsPool = {};
    const mpsFlatPool = {};
    monthNames.forEach(m => { mpsPool[m] = {}; mpsFlatPool[m] = []; });

    for (let r = 5; r < mpsRaw.length; r++) {
        const row = mpsRaw[r] || [];
        const code = (row[3] || '').toString().trim();
        const prod = (row[4] || '').toString().trim();
        if (!code || !prod || code.toUpperCase().includes('TOTAL')) continue;
        const base = getBase(prod);
        
        mpsMonthIdxs.forEach((colIdx, i) => {
            const m = monthNames[i];
            const q = parseInt(row[colIdx]) || 0;
            if (q > 0) {
                if (!mpsPool[m][base]) mpsPool[m][base] = [];
                for (let k = 0; k < q; k++) {
                    const item = { code, product: prod, base };
                    mpsPool[m][base].push(item);
                    mpsFlatPool[m].push(item);
                }
            }
        });
    }

    const prodMonthIdxs = [4, 7, 8, 9, 10, 12];
    const finalResults = [];
    const stats = { exact: 0, fuzzy: 0, fallback: 0, failed: 0 };

    monthNames.forEach((month, mIdx) => {
        const colIdx = prodMonthIdxs[mIdx];
        const monthNeeds = [];
        let lastSite = "", lastGroup = "";
        
        // 1. Collect all needs for the month
        for (let r = 6; r < prodRaw.length; r++) {
            const row = prodRaw[r] || [];
            if (row[0]) lastSite = row[0].toString().trim();
            if (row[1]) lastGroup = row[1].toString().trim();
            const model = (row[2] || '').toString().trim();
            if (!model || lastSite.includes('총합계') || model === 'Model' || model === '합계') continue;
            
            const q = parseInt(row[colIdx]) || 0;
            const rpm = (row[3] || '').toString();
            for (let k = 0; k < q; k++) {
                monthNeeds.push({ site: lastSite, group: lastGroup, model, rpm, month, myBase: norm(model), match: null });
            }
        }

        // 2. Pass 1: Exact Match (PRIORITY)
        monthNeeds.forEach(need => {
            if (mpsPool[month][need.myBase] && mpsPool[month][need.myBase].length > 0) {
                need.match = mpsPool[month][need.myBase].shift();
                // Sync with Flat Pool
                const fIdx = mpsFlatPool[month].findIndex(e => e.product === need.match.product);
                if (fIdx !== -1) mpsFlatPool[month].splice(fIdx, 1);
                stats.exact++;
            } else if (mpsPool[month]['L' + need.myBase] && mpsPool[month]['L' + need.myBase].length > 0) {
                need.match = mpsPool[month]['L' + need.myBase].shift();
                const fIdx = mpsFlatPool[month].findIndex(e => e.product === need.match.product);
                if (fIdx !== -1) mpsFlatPool[month].splice(fIdx, 1);
                stats.exact++;
            }
        });

        // 3. Pass 2: Fuzzy Match
        monthNeeds.forEach(need => {
            if (need.match) return;
            const fIdx = mpsFlatPool[month].findIndex(e => isSub(need.myBase, e.base) || isSub(e.base, need.myBase));
            if (fIdx !== -1) {
                need.match = mpsFlatPool[month].splice(fIdx, 1)[0];
                // Sync with Pool
                const pList = mpsPool[month][need.match.base];
                if (pList) pList.pop();
                stats.fuzzy++;
            }
        });

        // 4. Pass 3: Fallback (Random/First available)
        monthNeeds.forEach(need => {
            if (need.match) return;
            if (mpsFlatPool[month].length > 0) {
                need.match = mpsFlatPool[month].shift();
                const pList = mpsPool[month][need.match.base];
                if (pList) pList.pop();
                stats.fallback++;
            } else {
                stats.failed++;
            }
        });

        finalResults.push(...monthNeeds);
    });

    const outputRows = [['Site', 'Group', 'Model', 'RPM', 'Month', 'Code', 'Product']];
    finalResults.forEach(r => {
        outputRows.push([
            r.site, r.group, r.model, r.rpm === "0" ? "" : r.rpm, r.month, 
            r.match ? r.match.code : "", 
            r.match ? r.match.product : "UNMAPPED"
        ]);
    });

    fs.writeFileSync('_MPS_Final_Data_v3.csv', "\ufeff" + outputRows.map(r => r.map(v => `"${(v||'').toString().replace(/"/g, '""')}"`).join(',')).join('\n'));
    fs.writeFileSync('server_startup.log', `Final Result: Total ${finalResults.length}, Exact: ${stats.exact}, Fuzzy: ${stats.fuzzy}, Fallback: ${stats.fallback}, Failed: ${stats.failed}\n`);
    console.log('Extraction Done.');
}

try {
    runExtraction();
    app.post('/api/extract', (req, res) => {
        try { runExtraction(); res.json({ success: true }); } catch (e) { res.status(500).json({ error: e.message }); }
    });
    app.listen(PORT, () => console.log(`Server LIVE Port ${PORT}`));
} catch (e) { fs.writeFileSync('extraction_error.log', e.stack); }
