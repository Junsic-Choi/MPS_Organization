const XLSX = require('xlsx');

function getMatchKey(s) {
    if (!s) return '';
    let n = s.toString().toUpperCase().trim();
    
    // [1. ROMAN TO NUMBER - 가장 먼저 표준화]
    n = n.replace(/VIII/g, '8').replace(/VII/g, '7').replace(/VI/g, '6').replace(/IV/g, '4').replace(/IX/g, '9');
    n = n.replace(/III/g, '3').replace(/II/g, '2');

    // [2. SPECIALIZED SERIES RULES]
    
    // [SMX]: SMX2100STB -> SMX21SB, SMX2STB -> SMX21SB
    if (n.includes('SMX')) {
        n = n.replace(/SMX2(?![0-9])/g, 'SMX21');
        n = n.replace(/2100/g, '21').replace(/3100/g, '31').replace(/5100/g, '51');
        n = n.replace(/SYYB/g, 'SYY').replace(/STB/g, 'SB');
    }

    // [VCF]: VCF850LSR -> VF8LSR
    if (n.startsWith('VCF')) {
        n = n.replace('VCF', 'VF');
        n = n.replace(/(\d)\d0/, '$1');
    }

    // [MYNX]: MYNX6500/40 -> M654
    if (n.startsWith('MYNX')) {
        let taperMatch = n.match(/6500\/(\d)0/);
        if (taperMatch) n = 'M65' + taperMatch[1];
        else n = 'M' + n.substring(4);
    }
    
    // [VM]: VM6500 -> VM65
    if (n.startsWith('VM')) {
        if (n.startsWith('VMX')) n = 'M' + n.substring(3);
        else n = 'V' + n.substring(2);
    }

    // [DNM / SHORTHAND]: DNM750/50 -> DNM7550
    if (n.includes('DNM')) {
        n = n.replace(/DNM(\d+)0\/(\d+)/, 'DNM$1$2');
    }

    // [3. GENERIC NORMALIZATION]
    n = n.replace(/PUMA|LYNX/g, '').replace(/^P|^L/, '').replace(/\s+/g, '').trim();

    // [VTR]: VTR1620 -> VTR162
    if (n.startsWith('VTR')) {
        let numMatch = n.match(/VTR(\d+)/);
        if (numMatch) {
            let num = numMatch[1];
            if (num === '1620' || num === '162') num = '162';
            else if (num === '1216' || num === '121') num = '121';
            else if (num === '2025' || num === '202') num = '202';
            n = 'VTR' + num;
        }
    }

    if (n.startsWith('MYNX')) n = 'M' + n.substring(4);
    else if (n.startsWith('VMX')) n = 'M' + n.substring(3);
    else if (n.startsWith('VM') && !n.startsWith('VMX')) n = 'M' + n.substring(2);
    else if (n.startsWith('MP')) n = 'M' + n.substring(2);
    else if (n.startsWith('DNM')) n = 'D' + n.substring(3);
    else if (n.startsWith('DCM')) n = 'DC' + n.substring(3);
    else if (n.startsWith('DVF')) n = 'V' + n.substring(3);
    else if (n.startsWith('VCF')) n = 'V' + n.substring(3);
    else if (n.startsWith('VT') && !n.startsWith('VTR')) n = 'V' + n.substring(2);
    else if (n.startsWith('TT') && !n.startsWith('TTR')) n = 'T' + n.substring(2);
    
    if (n.startsWith('V') && !['VTR', 'VFC', 'VF'].some(p => n.startsWith(p))) {
        let digits = n.match(/\d{2,3}/);
        if (digits) n = 'V' + digits[0].substring(0, 2);
    }
    if (n.startsWith('TW')) {
        n = n.replace(/(\d+)(?:MZ|WB|W|B|Z|M)+\d*$/g, '$1');
        let base = n.match(/TW\d+/);
        if (base) n = base[0];
    }

    if (n.startsWith('GT2600')) {
        n = n.replace('XLMB', 'XLB').replace('XLMA', 'XLA').replace('XMB', 'XB').replace('XMA', 'XA');
    }

    let key = n.replace(/[^A-Z1-9]/g, '');
    key = key.replace(/\s+/g, '').replace(/[\u0000-\u001F\u007F-\u009F]/g, "");
    key = key.replace(/0/g, '');

    if (key.length >= 5) {
        key = key.replace(/[2-9]$/, '');
    }

    key = key.replace(/[A-Z]+$/, '');
    
    if (key.startsWith('DC')) {
        key = key.replace(/[A-Z]$/, '');
    }

    return key;
}

function extractMonth(s) {
    if (!s) return null;
    const str = s.toString().trim();
    const yearMatch = str.match(/26\.(\d+)/);
    if (yearMatch) {
        const n = parseInt(yearMatch[1]);
        if (n >= 1 && n <= 12) return n;
    }
    const monthWordMatch = str.match(/(?:^|\s)(\d+)\s*월/);
    if (monthWordMatch) {
        const n = parseInt(monthWordMatch[1]);
        if (n >= 1 && n <= 12) return n;
    }
    if (/^\d+$/.test(str)) {
        const n = parseInt(str);
        if (n >= 1 && n <= 12) return n;
    }
    return null;
}

function processMpsFile(input, rules = {}) {
    const siteMaster = rules.siteMaster || {};
    
    let wb;
    if (Buffer.isBuffer(input)) {
        wb = XLSX.read(input, { type: 'buffer' });
    } else {
        wb = XLSX.readFile(input);
    }

    const prodWsName = '생산배포용';
    const masterWsName = 'MPS';
    
    const prodWs = wb.Sheets[prodWsName] || wb.Sheets[wb.SheetNames[0]];
    const masterWs = wb.Sheets[masterWsName] || wb.Sheets[wb.SheetNames[1]] || wb.Sheets[wb.SheetNames[0]];
    
    const prodRaw = XLSX.utils.sheet_to_json(prodWs, { header: 1 });
    const masterRaw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });

    const finalResults = [];
    const unusedData = [];
    const masterPlanPool = [];
    const quotaPool = {};
    const masterPlan = {};
    const masterModelsByGroup = { "전체기종": new Set() };

    // [MASTER PLAN ANALYSIS]
    let masterHeaderIdx = -1;
    const masterMonthCols = [];

    // Find Month Row and Type Row
    let monthRowIdx = -1;
    let typeRowIdx = -1;

    for (let r = 0; r < Math.min(50, masterRaw.length); r++) {
        const rowStr = (masterRaw[r] || []).join('|');
        if (rowStr.includes('월')) monthRowIdx = r;
        if (rowStr.includes('생산') && rowStr.includes('판매') && r > monthRowIdx) {
            typeRowIdx = r;
            break;
        }
    }

    if (monthRowIdx !== -1 && typeRowIdx !== -1) {
        const monthRow = masterRaw[monthRowIdx];
        const typeRow = masterRaw[typeRowIdx];

        typeRow.forEach((cell, idx) => {
            const type = (cell || '').toString().trim();
            if (type === '생산') {
                for (let c = idx; c >= 0; c--) {
                    const mNum = extractMonth(monthRow[c]);
                    if (mNum !== null) {
                        masterMonthCols.push({ name: mNum + '월', col: idx });
                        console.log(`[Engine] Found Master Column: ${mNum}월 생산 at Col ${idx}`);
                        break;
                    }
                }
            }
        });
        masterHeaderIdx = typeRowIdx;
    }

    if (masterMonthCols.length === 0) {
        masterHeaderIdx = 2;
        masterMonthCols.push({ name: '3월', col: 8 }, { name: '4월', col: 14 }, { name: '5월', col: 20 }, { name: '6월', col: 26 }, { name: '7월', col: 32 }, { name: '8월', col: 38 });
    }

    masterRaw.forEach((row, idx) => {
        if (idx <= masterHeaderIdx) return;
        
        const mGroup = (row[2] || row[1] || '').toString().trim();
        const mCode = (row[3] || row[1] || '').toString().trim();
        const pName = (row[4] || row[2] || '').toString().trim();
        
        let finalSite = (row[6] || '').toString().trim();
        if (finalSite === '1842' || finalSite === 1842) finalSite = '성주';
        else if (finalSite === '1840' || finalSite === 1840) finalSite = '남산';
        
        if (pName || mCode) {
            const modelPart = (pName || mCode).split('-')[0].trim();
            const key = getMatchKey(modelPart);
            
            masterModelsByGroup["전체기종"].add(modelPart);
            if (mGroup) {
                if (!masterModelsByGroup[mGroup]) masterModelsByGroup[mGroup] = new Set();
                masterModelsByGroup[mGroup].add(modelPart);
            }

            masterMonthCols.forEach(mCol => {
                const q = parseInt(row[mCol.col]) || 0;
                if (q > 0) {
                    masterPlanPool.push({
                        Month: mCol.name, Group: mGroup, Model: modelPart,
                        Product: pName, Qty: q, Code: mCode
                    });

                    const mNum = parseInt(mCol.name);
                    if (!quotaPool[finalSite]) quotaPool[finalSite] = {};
                    if (!quotaPool[finalSite][key]) quotaPool[finalSite][key] = {};
                    if (!quotaPool[finalSite][key][mNum]) quotaPool[finalSite][key][mNum] = [];
                    quotaPool[finalSite][key][mNum].push({
                        Qty: q, Product: pName, Code: mCode
                    });
                }
            });
        }
    });

    // [PRODUCTION DATA ANALYSIS]
    let prodHeaderIdx = -1;
    const prodMonthCols = [];

    for (let r = 0; r < Math.min(50, prodRaw.length); r++) {
        const row = prodRaw[r] || [];
        const currentMonths = [];
        row.forEach((cell, idx) => {
            const s = (cell || '').toString().trim();
            const mNum = extractMonth(s);
            if (mNum !== null) currentMonths.push({ name: mNum + '월', col: idx });
        });

        if (currentMonths.length >= 2) {
            prodMonthCols.push(...currentMonths);
            prodHeaderIdx = r;
            break;
        }
    }

    let runningSite = '', runningGroup = '', runningModel = '';
    for (let r = prodHeaderIdx + 1; r < prodRaw.length; r++) {
        const row = prodRaw[r];
        if (!row || row.length < 3) continue;

        const site = (row[0] || '').toString().trim();
        const group = (row[1] || '').toString().trim();
        const model = (row[2] || '').toString().trim();
        const rpm = (row[3] || '').toString().trim();

        if (site) runningSite = site;
        if (group) runningGroup = group;
        if (model) runningModel = model;

        if (runningSite === '총합계') continue;

        const key = getMatchKey(runningModel);
        
        let finalSite = runningSite;
        if (finalSite === '1840' || finalSite === 1840) finalSite = '남산';
        if (finalSite === '1842' || finalSite === 1842) finalSite = '성주';
        if (siteMaster[runningSite]) finalSite = siteMaster[runningSite];
        finalSite = finalSite.replace(/^\d+\.\s*/, '').trim();

        prodMonthCols.forEach(mObj => {
            let remaining = parseInt(row[mObj.col]) || 0;
            if (remaining > 0) {
                const mNum = parseInt(mObj.name);
                const candidates = (quotaPool[key] && quotaPool[key][mNum]) || [];
                
                candidates.sort((a, b) => {
                    const aM = a.Model.toUpperCase();
                    const bM = b.Model.toUpperCase();
                    const rM = runningModel.toUpperCase();
                    const aScore = (aM === rM) ? 100 : (aM.includes(rM) || rM.includes(aM) ? 10 : 0);
                    const bScore = (bM === rM) ? 100 : (bM.includes(rM) || rM.includes(bM) ? 10 : 0);
                    return bScore - aScore;
                });

                for (const cand of candidates) {
                    if (remaining <= 0) break;
                    if (cand.Qty <= 0) continue;

                    const take = Math.min(remaining, cand.Qty);
                    finalResults.push({
                        Site: finalSite, Group: runningGroup, Model: runningModel,
                        RPM: rpm, Month: mObj.name, Code: cand.Code,
                        Product: cand.Product, Qty: take
                    });
                    remaining -= take;
                    cand.Qty -= take;

                    if (!masterPlan[runningGroup]) masterPlan[runningGroup] = {};
                    if (!masterPlan[runningGroup][runningModel]) masterPlan[runningGroup][runningModel] = {};
                    if (!masterPlan[runningGroup][runningModel][mObj.name]) masterPlan[runningGroup][runningModel][mObj.name] = [];
                    masterPlan[runningGroup][runningModel][mObj.name].push({ rpm: rpm, qty: take });
                }

                if (remaining > 0) {
                    unusedData.push({
                        Site: finalSite, Group: runningGroup, Model: runningModel,
                        RPM: rpm, Month: mObj.name, Qty: remaining, ModelCode: '', ProductName: '', Category: '미매칭'
                    });
                }
            }
        });
    }

    const finalMasterModels = {};
    for (const g in masterModelsByGroup) {
        finalMasterModels[g] = Array.from(masterModelsByGroup[g]).sort();
    }

    return { 
        finalResults, unusedData, masterPlanPool, masterPlan, masterModelsByGroup: finalMasterModels 
    };
}

module.exports = { processMpsFile, getMatchKey };
