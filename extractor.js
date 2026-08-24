const XLSX = require('xlsx');

function getMatchKey(s) {
    if (!s) return '';
    let n = s.toString().toUpperCase().trim();
    
    // [1. CLEANUP & BRAND PROTECTION]
    n = n.replace(/PUMA|LYNX/g, '').replace(/\s+/g, '').trim();
    // Only remove leading P/L if it's NOT LEO
    if (n.startsWith('P') || (n.startsWith('L') && !n.startsWith('LEO'))) {
        n = n.substring(1);
    }

    // [2. ROMAN TO NUMBER]
    n = n.replace(/VIII/g, '8').replace(/VII/g, '7').replace(/VI/g, '6').replace(/IV/g, '4').replace(/IX/g, '9');
    n = n.replace(/III/g, '3').replace(/II/g, '2');

    // [3. SMX SPECIAL]
    if (n.includes('SMX')) {
        n = n.replace(/SMX2(?![0-9])/g, 'SMX21');
        n = n.replace(/2100/g, '21').replace(/3100/g, '31').replace(/5100/g, '51');
        n = n.replace(/SYYB/g, 'SYY').replace(/STB/g, 'SB');
    }

    // [4. PINPOINT MAPPING - SHARED FOR BOTH SIDES]
    if (n.includes('DNM750L/50') || n === 'DNM755L') return 'D755L';
    if (n.includes('DNM750/50') || n === 'DNM7550' || n === 'DNM755') return 'D755';
    if (n.includes('ST38GS')) return 'ST38GS2'; 
    if (n.includes('ST10GS')) return 'ST1GS2';
    if (n.includes('DST20')) return 'DST20';
    if (n.includes('LEO16')) return 'LEO16';
    
    // VTR/VF Grouping
    if (n.startsWith('VTR')) {
        let m = n.match(/VTR(\d{2})/);
        if (m) return 'VTR' + m[1];
    }
    if (n.startsWith('DVF')) {
        let m = n.match(/DVF(\d)/);
        if (m) return 'DVF' + m[1];
    }
    if (n.startsWith('VCF') || n.startsWith('VF')) {
        let m = n.match(/(?:VCF|VF)(\d)/);
        if (m) return 'VF' + m[1];
    }

    // [5. LEGACY STABLE RULES]
    n = n.replace(/DNM(\d+)0\/(\d+)/, 'DNM$1$2');
    if (n.startsWith('MYNX')) n = 'M' + n.substring(4);
    else if (n.startsWith('VMX')) n = 'M' + n.substring(3);
    else if (n.startsWith('VM')) n = 'V' + n.substring(2);
    else if (n.startsWith('MP')) n = 'M' + n.substring(2);
    else if (n.startsWith('DNM')) n = 'D' + n.substring(3);
    else if (n.startsWith('DBC')) n = 'DB' + n.substring(3);
    else if (n.startsWith('DCM')) n = 'DC' + n.substring(3);
    else if (n.startsWith('DVF')) n = 'V' + n.substring(3);
    else if (n.startsWith('VT') && !n.startsWith('VTR')) n = 'V' + n.substring(2);
    else if (n.startsWith('TT') && !n.startsWith('TTR')) n = 'T' + n.substring(2);

    // [6. FINAL NORMALIZATION]
    let key = n.replace(/[^A-Z1-9]/g, '');
    
    if (key === 'DST2') key = 'DST20';

    key = key.replace(/0/g, '');
    
    // Protect L for series like D755L
    if (key === 'D755' && n.includes('L')) {
        return 'D755L';
    }

    key = key.replace(/[A-Z]+$/, '');
    if (key.startsWith('DC')) key = key.replace(/[A-Z]$/, '');
    
    return key;
}

function extractMonth(s) {
    if (!s) return null;
    if (s instanceof Date) return s.getMonth() + 1;
    
    const str = s.toString().trim();
    // 2026.05 or 26.05
    const dotMatch = str.match(/(?:20)?26\.(\d+)/);
    if (dotMatch) {
        const n = parseInt(dotMatch[1]);
        if (n >= 1 && n <= 12) return n;
    }
    // 5월
    const monthWordMatch = str.match(/(\d+)\s*월/);
    if (monthWordMatch) {
        const n = parseInt(monthWordMatch[1]);
        if (n >= 1 && n <= 12) return n;
    }
    // Just a number 1-12
    if (/^\d+$/.test(str)) {
        const n = parseInt(str);
        if (n >= 1 && n <= 12) return n;
    }
    return null;
}

function selectMeta(foundMetaList, finalSite, month, pName, mModel) {
    if (!foundMetaList || foundMetaList.length === 0) return null;
    if (foundMetaList.length === 1) return foundMetaList[0];
    
    // 1. Filter by site
    let mpsMainPlant = '남산';
    if (finalSite === '성주' || finalSite.includes('성주')) {
        mpsMainPlant = '성주';
    }
    
    const siteMatches = foundMetaList.filter(item => {
        const itemSiteClean = item.site.replace(/^\d+\.\s*/, '').trim();
        let itemMainPlant = '남산';
        if (itemSiteClean.includes('성주') || itemSiteClean.includes('성우')) {
            itemMainPlant = '성주';
        }
        return itemMainPlant === mpsMainPlant;
    });
    
    const candidates = siteMatches.length > 0 ? siteMatches : foundMetaList;
    if (candidates.length === 1) return candidates[0];
    
    // 2. Parse product suffix to find control type (H, S, F)
    const parts = pName.split('-');
    let targetKeyword = '';
    if (parts.length > 1) {
        const subCode = parts[1].toUpperCase();
        if (subCode.includes('H')) targetKeyword = 'H';
        else if (subCode.includes('S')) targetKeyword = 'S';
    }
    if (!targetKeyword) targetKeyword = 'F'; // Default to Fanuc
    
    // 3. Filter candidates by control type
    let typeFiltered = [];
    if (targetKeyword === 'H') {
        typeFiltered = candidates.filter(c => c.rpm.toUpperCase().includes('H/H'));
    } else if (targetKeyword === 'S') {
        typeFiltered = candidates.filter(c => {
            const up = c.rpm.toUpperCase();
            return up.includes('SONE') || up.includes('SIEMENS') || up.includes('지멘스');
        });
    } else if (targetKeyword === 'F') {
        typeFiltered = candidates.filter(c => {
            const up = c.rpm.toUpperCase();
            return !up.includes('H/H') && !up.includes('SONE') && !up.includes('SIEMENS') && !up.includes('지멘스');
        });
    }
    
    const finalCandidates = typeFiltered.length > 0 ? typeFiltered : candidates;
    if (finalCandidates.length === 1) return finalCandidates[0];
    
    // 4. Filter by monthly plan quantity in the production sheet
    const withPlanQty = finalCandidates.filter(c => c.monthlyPlan && c.monthlyPlan[month] > 0);
    if (withPlanQty.length > 0) {
        return withPlanQty.reduce((max, c) => (c.monthlyPlan[month] > max.monthlyPlan[month] ? c : max), withPlanQty[0]);
    }
    
    // 5. Fallback
    return finalCandidates[0];
}

function processMpsFile(input, rules = {}) {
    const siteMaster = rules.siteMaster || {};
    
    let wb;
    if (Buffer.isBuffer(input)) {
        wb = XLSX.read(input, { type: 'buffer' });
    } else {
        wb = XLSX.readFile(input);
    }

    const findSheet = (keywords) => {
        return wb.SheetNames.find(name => keywords.some(k => name.includes(k)));
    };

    const prodWsName = findSheet(['생산배포', '배포용', 'Production']) || wb.SheetNames[0];
    const masterWsName = findSheet(['MPS', 'Master', '마스터']) || wb.SheetNames[1] || wb.SheetNames[0];
    
    console.log(`[Engine] Sheets selected: Prod="${prodWsName}", Master="${masterWsName}"`);

    const prodWs = wb.Sheets[prodWsName];
    const masterWs = wb.Sheets[masterWsName];
    
    const prodRaw = XLSX.utils.sheet_to_json(prodWs, { header: 1 });
    const masterRaw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });

    const finalResults = [];
    const unusedData = [];
    const masterPlanPool = [];
    const quotaPool = {};
    const masterPlan = {};
    const masterModelsByGroup = { "전체기종": new Set() };

    // --- Dynamic Column Detection Helper ---
    const findCols = (raw, keywordsMap, maxRows = 20) => {
        for (let r = 0; r < Math.min(maxRows, raw.length); r++) {
            const row = raw[r] || [];
            const rowResult = {};

            // 1. Exact match pass for this row
            row.forEach((cell, idx) => {
                const val = String(cell || '').trim().toUpperCase();
                for (const [key, keywords] of Object.entries(keywordsMap)) {
                    if (rowResult[key] === undefined) {
                        const hasExact = keywords.some(k => val === k.toUpperCase());
                        if (hasExact) {
                            rowResult[key] = idx;
                        }
                    }
                }
            });

            // 2. Partial match pass for this row (only for keys not yet matched)
            row.forEach((cell, idx) => {
                const val = String(cell || '').trim().toUpperCase();
                for (const [key, keywords] of Object.entries(keywordsMap)) {
                    if (rowResult[key] === undefined) {
                        const hasPartial = keywords.some(k => val.includes(k.toUpperCase()));
                        if (hasPartial) {
                            rowResult[key] = idx;
                        }
                    }
                }
            });

            // Check if this row is the header row
            if (rowResult.Model !== undefined && (rowResult.Site !== undefined || rowResult.Group !== undefined)) {
                rowResult.headerRowIdx = r;
                return rowResult;
            }
        }
        return {};
    };

    // [MASTER PLAN ANALYSIS]
    const masterKeywords = {
        Model: ['기종', 'Model'],
        Group: ['그룹', 'Group', 'Series'],
        Site: ['사업장', '공장', 'Site'],
        PL: ['PL', '제품군'],
        Ver: ['Ver', '버전'],
        Pjt: ['PJT', '프로젝트', 'Product Name', 'Product']
    };
    const mCols = findCols(masterRaw, masterKeywords, 50);
    
    // [PHASE 1: Metadata Mapping from Production (배포용) Sheet]
    const prodKeywords = {
        Model: ['기종', 'Model'],
        Group: ['기종분류', '시리즈', 'Series', '그룹'],
        Site: ['생산처', '공장', '사업장', 'Site'],
        RPM: ['RPM', '주축', 'Spindle']
    };
    const pCols = findCols(prodRaw, prodKeywords, 50);
    const prodHeaderIdx = pCols.headerRowIdx !== undefined ? pCols.headerRowIdx : -1;

    // Legacy V3.2 logic: Use "배포용" as a lookup for Site, Group, and RPM
    const metaMap = {};
    let lastMetaSite = '', lastMetaGroup = '', lastMetaModel = '';
    
    // Find month columns in production sheet dynamically
    const prodMonthCols = {};
    if (prodHeaderIdx !== -1) {
        const headerRow = prodRaw[prodHeaderIdx];
        headerRow.forEach((cell, idx) => {
            if (!cell) return;
            const str = cell.toString().trim();
            const mMatch = str.match(/^(\d+)(?:월)?/);
            if (mMatch) {
                const mNum = parseInt(mMatch[1]);
                if (mNum >= 1 && mNum <= 12) {
                    prodMonthCols[mNum + '월'] = idx;
                }
            }
        });
    }
    
    prodRaw.forEach((row, idx) => {
        if (idx <= prodHeaderIdx) return;
        const s = (row[pCols.Site] || '').toString().trim();
        const g = (row[pCols.Group] || '').toString().trim();
        const m = (row[pCols.Model] || '').toString().trim();
        const rpm = (row[pCols.RPM] || '').toString().trim();
        
        if (s) lastMetaSite = s;
        if (g) lastMetaGroup = g;
        if (m) lastMetaModel = m;
        
        if (lastMetaModel) {
            const mKey = getMatchKey(lastMetaModel);
            if (!metaMap[mKey]) {
                metaMap[mKey] = [];
            }
            
            // Extract monthly planned quantities for this RPM row
            const monthlyPlan = {};
            for (const [monthName, colIdx] of Object.entries(prodMonthCols)) {
                monthlyPlan[monthName] = parseInt(row[colIdx]) || 0;
            }
            
            const exists = metaMap[mKey].some(item => item.site === lastMetaSite && item.model === lastMetaModel && item.rpm === rpm);
            if (!exists) {
                metaMap[mKey].push({ 
                    site: lastMetaSite, 
                    group: lastMetaGroup, 
                    model: lastMetaModel, 
                    rpm: rpm,
                    monthlyPlan: monthlyPlan
                });
            } else {
                const existingItem = metaMap[mKey].find(item => item.site === lastMetaSite && item.model === lastMetaModel && item.rpm === rpm);
                if (existingItem) {
                    for (const [monthName, colIdx] of Object.entries(prodMonthCols)) {
                        existingItem.monthlyPlan[monthName] = (existingItem.monthlyPlan[monthName] || 0) + (parseInt(row[colIdx]) || 0);
                    }
                }
            }
        }
    });
    console.log(`[Engine] Meta Map built: ${Object.keys(metaMap).length} models found in Production sheet`);

    let masterHeaderIdx = mCols.headerRowIdx !== undefined ? mCols.headerRowIdx : -1;
    const masterMonthCols = [];

    // Find Month Row and Type Row
    let monthRowIdx = -1;
    let typeRowIdx = -1;

    for (let r = 0; r < Math.min(50, masterRaw.length); r++) {
        const rowStr = (masterRaw[r] || []).join('|');
        if (rowStr.includes('월') && monthRowIdx === -1) monthRowIdx = r;
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
                        break;
                    }
                }
            }
        });

        // 6개월 계획 한계 설정 (감지된 달 중 상위 6개만 계획 수립에 반영하고 7번째 달은 제외)
        if (masterMonthCols.length > 6) {
            masterMonthCols.splice(6);
        }

        masterHeaderIdx = typeRowIdx;
    }

    if (masterMonthCols.length === 0) {
        console.warn('[Engine] No month columns found in Master, using defaults');
        masterHeaderIdx = 2;
        masterMonthCols.push({ name: '3월', col: 8 }, { name: '4월', col: 14 }, { name: '5월', col: 20 }, { name: '6월', col: 26 }, { name: '7월', col: 32 }, { name: '8월', col: 38 });
    }

    // [PHASE 2: Master Plan (MPS) Processing]
    masterRaw.forEach((row, idx) => {
        if (idx <= masterHeaderIdx) return;
        
        const mModelIdx = mCols.Model !== undefined ? mCols.Model : 1;
        const mGroupIdx = mCols.Group !== undefined ? mCols.Group : 2;
        const mPlIdx = mCols.PL !== undefined ? mCols.PL : 1;
        const mVerIdx = mCols.Ver !== undefined ? mCols.Ver : 7;
        const mSiteIdx = mCols.Site !== undefined ? mCols.Site : 6;
        const mPjtIdx = mCols.Pjt !== undefined ? mCols.Pjt : 4;

        const pName = (row[mPjtIdx] || row[mModelIdx] || '').toString().trim();
        const plCode = (row[mPlIdx] || '').toString().trim();
        const verCode = (row[mVerIdx] || '').toString().trim();
        let mModel = (row[mModelIdx] || '').toString().trim();
        
        // Match logic following V3.2: Use Product Name prefix or Model
        const pNamePrefix = pName.split('-')[0].trim();
        const mKey = getMatchKey(pNamePrefix || mModel);

        let finalSite = (row[mSiteIdx] || '').toString().trim();
        let mGroup = (row[mGroupIdx] || row[mPlIdx] || '').toString().trim();
        let mRPM = '';

        // Enrichment from Meta Map (Legacy V3.2 style)
        // Try exact match first, then partial match like V3.2
        let foundMetaList = metaMap[mKey];
        if (!foundMetaList || foundMetaList.length === 0) {
            const possibleKeys = Object.keys(metaMap);
            const bestMatch = possibleKeys.find(k => {
                if (k.length <= 1 || mKey.length <= 1) return false;
                return k.startsWith(mKey) || mKey.startsWith(k);
            });
            if (bestMatch) foundMetaList = metaMap[bestMatch];
        }

        let foundMeta = null;
        if (foundMetaList && foundMetaList.length > 0) {
            if (foundMetaList.length === 1) {
                foundMeta = foundMetaList[0];
            } else {
                // Determine main plant from originalSite (stored in finalSite initially)
                let mpsMainPlant = '남산';
                if (finalSite === '1842' || finalSite.includes('성주')) {
                    mpsMainPlant = '성주';
                }
                
                // Select matching main plant from list
                foundMeta = foundMetaList.find(item => {
                    const itemSiteClean = item.site.replace(/^\d+\.\s*/, '').trim();
                    let itemMainPlant = '남산';
                    if (itemSiteClean.includes('성주') || itemSiteClean.includes('성우')) {
                        itemMainPlant = '성주';
                    }
                    return itemMainPlant === mpsMainPlant;
                });
                
                if (!foundMeta) foundMeta = foundMetaList[0];
            }
        }

        if (foundMeta) {
            // Prioritize Production sheet (Meta Map) values
            finalSite = foundMeta.site;
            mGroup = foundMeta.group;
            mRPM = foundMeta.rpm;
            // Use canonical model name from Production sheet
            if (foundMeta.model) mModel = foundMeta.model;
        }

        // Standardize Site
        if (verCode && siteMaster[verCode]) {
            finalSite = siteMaster[verCode];
        } else if (foundMeta && foundMeta.site) {
            finalSite = foundMeta.site;
        } else if (siteMaster[finalSite]) {
            finalSite = siteMaster[finalSite];
        }
        
        if (verCode === '9ACE') finalSite = '지티테크';
        if (verCode === '9AYA' || pName.toUpperCase().startsWith('LEO') || mModel.toUpperCase().startsWith('LEO')) {
            finalSite = 'LEO';
        }
        
        if (finalSite === '1842' || finalSite === 1842 || finalSite.includes('성주')) finalSite = '성주';
        else if (finalSite === '1840' || finalSite === 1840 || finalSite.includes('남산')) finalSite = '남산';

        finalSite = finalSite.replace(/^\d+\.\s*/, '').trim();
        if (finalSite.includes('LEO') && !pName.toUpperCase().startsWith('LEO') && !mModel.toUpperCase().startsWith('LEO') && verCode !== '9AYA') {
            finalSite = (row[mSiteIdx] === 1842 || row[mSiteIdx] === '1842') ? '성주' : '남산';
        }
        if (finalSite.includes('지티')) finalSite = '지티테크';
        
        if (pName || mModel) {
            const modelPart = (pName || mModel).split('-')[0].trim();
            const mCode = verCode || plCode;
            
            masterModelsByGroup["전체기종"].add(modelPart);
            if (mGroup) {
                if (!masterModelsByGroup[mGroup]) masterModelsByGroup[mGroup] = new Set();
                masterModelsByGroup[mGroup].add(modelPart);
            }

            masterMonthCols.forEach(mCol => {

                const q = parseInt(row[mCol.col]) || 0;
                if (q > 0) {
                    // Resolve monthly RPM dynamically
                    let monthlyRPM = mRPM;
                    if (foundMetaList && foundMetaList.length > 1) {
                        const resolvedMeta = selectMeta(foundMetaList, finalSite, mCol.name, pName, mModel);
                        if (resolvedMeta) {
                            monthlyRPM = resolvedMeta.rpm;
                        }
                    }

                    // Create main results directly from MPS (This makes it look like V3.2)
                    finalResults.push({
                        Site: finalSite || '기타', 
                        Group: mGroup || '기타', 
                        Model: modelPart,
                        RPM: monthlyRPM, 
                        Month: mCol.name, 
                        Code: mCode,
                        Product: pName, 
                        Qty: q
                    });

                    masterPlanPool.push({
                        Month: mCol.name, Group: mGroup, Model: modelPart,
                        Product: pName, Qty: q, Code: mCode, Site: finalSite
                    });
                }
            });
        }
    });

    console.log(`[Engine] Extraction complete: ${finalResults.length} plan entries generated from MPS`);

    const finalMasterModels = {};
    for (const g in masterModelsByGroup) {
        finalMasterModels[g] = Array.from(masterModelsByGroup[g]).sort();
    }

    return { 
        finalResults, unusedData: [], masterPlanPool, masterPlan, masterModelsByGroup: finalMasterModels 
    };
}

module.exports = { processMpsFile, getMatchKey };
