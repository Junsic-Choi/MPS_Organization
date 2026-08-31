const express = require('express');
const path = require('path');
const cors = require('cors');
const fs = require('fs');

const app = express();
const PORT = 8895;

app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static(__dirname));

const DATA_FILE = path.join(__dirname, 'shopfloor_data.json');

// Default Bay Template (Organized by Real MES Building and Bay Codes)
function getDefaultBays() {
    const bays = [];
    
    // MC 1직 (C동) - D구역 1~14, E구역 1~10, F구역 1~6
    for (let i = 1; i <= 14; i++) bays.push({ id: `mc1-d${i}`, shift: 'MC1직', area: 'C동', bay: `D${i}`, assigned: false, model: '', serial: '', salesDoc: '', customer: '', worker: '', currentProcess: 'BASE', spec: '', issue: '', startDate: '', deliveryDate: '' });
    for (let i = 1; i <= 10; i++) bays.push({ id: `mc1-e${i}`, shift: 'MC1직', area: 'C동', bay: `E${i}`, assigned: false, model: '', serial: '', salesDoc: '', customer: '', worker: '', currentProcess: 'BASE', spec: '', issue: '', startDate: '', deliveryDate: '' });
    for (let i = 1; i <= 6; i++) bays.push({ id: `mc1-f${i}`, shift: 'MC1직', area: 'C동', bay: `F${i}`, assigned: false, model: '', serial: '', salesDoc: '', customer: '', worker: '', currentProcess: 'BASE', spec: '', issue: '', startDate: '', deliveryDate: '' });

    // MC 2직 (MC동) - C구역 1~19, D구역 1~19
    for (let i = 1; i <= 19; i++) bays.push({ id: `mc2-c${i}`, shift: 'MC2직', area: 'MC동', bay: `C${i}`, assigned: false, model: '', serial: '', salesDoc: '', customer: '', worker: '', currentProcess: 'BASE', spec: '', issue: '', startDate: '', deliveryDate: '' });
    for (let i = 1; i <= 19; i++) bays.push({ id: `mc2-d${i}`, shift: 'MC2직', area: 'MC동', bay: `D${i}`, assigned: false, model: '', serial: '', salesDoc: '', customer: '', worker: '', currentProcess: 'BASE', spec: '', issue: '', startDate: '', deliveryDate: '' });

    // MC 3직 (MC동) - B구역 1~19, A구역 1~18
    for (let i = 1; i <= 19; i++) bays.push({ id: `mc3-b${i}`, shift: 'MC3직', area: 'MC동', bay: `B${i}`, assigned: false, model: '', serial: '', salesDoc: '', customer: '', worker: '', currentProcess: 'BASE', spec: '', issue: '', startDate: '', deliveryDate: '' });
    for (let i = 1; i <= 18; i++) bays.push({ id: `mc3-a${i}`, shift: 'MC3직', area: 'MC동', bay: `A${i}`, assigned: false, model: '', serial: '', salesDoc: '', customer: '', worker: '', currentProcess: 'BASE', spec: '', issue: '', startDate: '', deliveryDate: '' });

    // MC 4직 (FA동) - A구역 1~14, B구역 1~12, C구역 5~13, D구역 1~6
    for (let i = 1; i <= 14; i++) bays.push({ id: `mc4-a${i}`, shift: 'MC4직', area: 'FA동', bay: `A${i}`, assigned: false, model: '', serial: '', salesDoc: '', customer: '', worker: '', currentProcess: 'BASE', spec: '', issue: '', startDate: '', deliveryDate: '' });
    for (let i = 1; i <= 12; i++) bays.push({ id: `mc4-b${i}`, shift: 'MC4직', area: 'FA동', bay: `B${i}`, assigned: false, model: '', serial: '', salesDoc: '', customer: '', worker: '', currentProcess: 'BASE', spec: '', issue: '', startDate: '', deliveryDate: '' });
    for (let i = 5; i <= 13; i++) bays.push({ id: `mc4-c${i}`, shift: 'MC4직', area: 'FA동', bay: `C${i}`, assigned: false, model: '', serial: '', salesDoc: '', customer: '', worker: '', currentProcess: 'BASE', spec: '', issue: '', startDate: '', deliveryDate: '' });
    for (let i = 1; i <= 6; i++) bays.push({ id: `mc4-d${i}`, shift: 'MC4직', area: 'FA동', bay: `D${i}`, assigned: false, model: '', serial: '', salesDoc: '', customer: '', worker: '', currentProcess: 'BASE', spec: '', issue: '', startDate: '', deliveryDate: '' });

    return bays;
}

// 1. Root / UI Routing
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'mc_shopfloor.html'));
});

app.get('/shopfloor', (req, res) => {
    res.sendFile(path.join(__dirname, 'mc_shopfloor.html'));
});

// 2. Load Shopfloor Bays API
app.get('/api/shopfloor', (req, res) => {
    try {
        if (fs.existsSync(DATA_FILE)) {
            const content = fs.readFileSync(DATA_FILE, 'utf8');
            const data = JSON.parse(content);
            return res.json({ success: true, bays: data.bays || [] });
        }
        // If file doesn't exist, create with clean empty templates
        const defaultBays = getDefaultBays();
        fs.writeFileSync(DATA_FILE, JSON.stringify({ bays: defaultBays, updatedAt: new Date().toISOString() }, null, 2), 'utf8');
        res.json({ success: true, bays: defaultBays });
    } catch (err) {
        console.error('[shopfloor] Load failed:', err.message);
        res.status(500).json({ success: false, error: err.message });
    }
});

// 3. Save Shopfloor Bays API
app.post('/api/shopfloor/save', (req, res) => {
    try {
        const { bays } = req.body;
        if (!Array.isArray(bays)) {
            return res.status(400).json({ success: false, error: 'bays array is required' });
        }
        fs.writeFileSync(DATA_FILE, JSON.stringify({ bays, updatedAt: new Date().toISOString() }, null, 2), 'utf8');
        console.log(`[shopfloor] Saved ${bays.length} bays successfully at ${new Date().toLocaleTimeString()}`);
        res.json({ success: true, count: bays.length });
    } catch (err) {
        console.error('[shopfloor] Save failed:', err.message);
        res.status(500).json({ success: false, error: err.message });
    }
});

// 3-1. Move / Swap Bay API
app.post('/api/shopfloor/move-bay', (req, res) => {
    try {
        const { sourceBayId, targetBayId } = req.body;
        if (!sourceBayId || !targetBayId) {
            return res.status(400).json({ success: false, error: 'sourceBayId and targetBayId required' });
        }

        let bays = [];
        if (fs.existsSync(DATA_FILE)) {
            const data = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
            bays = data.bays || [];
        } else {
            bays = getDefaultBays();
        }

        const srcIndex = bays.findIndex(b => b.id === sourceBayId);
        const tgtIndex = bays.findIndex(b => b.id === targetBayId);

        if (srcIndex === -1 || tgtIndex === -1) {
            return res.status(404).json({ success: false, error: 'Bay not found' });
        }

        const src = bays[srcIndex];
        const tgt = bays[tgtIndex];

        // Backup static properties
        const srcMeta = { id: src.id, shift: src.shift, bay: src.bay };
        const tgtMeta = { id: tgt.id, shift: tgt.shift, bay: tgt.bay };

        // Data to transfer
        const srcData = {
            assigned: src.assigned,
            model: src.model,
            serial: src.serial,
            salesDoc: src.salesDoc,
            customer: src.customer,
            worker: src.worker,
            currentProcess: src.currentProcess,
            spec: src.spec,
            issue: src.issue,
            startDate: src.startDate,
            deliveryDate: src.deliveryDate,
            source: src.source || 'MANUAL'
        };

        const tgtData = {
            assigned: tgt.assigned,
            model: tgt.model,
            serial: tgt.serial,
            salesDoc: tgt.salesDoc,
            customer: tgt.customer,
            worker: tgt.worker,
            currentProcess: tgt.currentProcess,
            spec: tgt.spec,
            issue: tgt.issue,
            startDate: tgt.startDate,
            deliveryDate: tgt.deliveryDate,
            source: tgt.source || 'MANUAL'
        };

        if (tgt.assigned) {
            // Swap if target is already assigned
            bays[srcIndex] = { ...srcMeta, ...tgtData };
            bays[tgtIndex] = { ...tgtMeta, ...srcData };
        } else {
            // Move to empty target bay and clear source bay
            bays[tgtIndex] = { ...tgtMeta, ...srcData };
            bays[srcIndex] = {
                ...srcMeta,
                assigned: false,
                model: '',
                serial: '',
                salesDoc: '',
                customer: '',
                worker: '',
                currentProcess: 'BASE',
                spec: '',
                issue: '',
                startDate: '',
                deliveryDate: '',
                source: 'MANUAL'
            };
        }

        fs.writeFileSync(DATA_FILE, JSON.stringify({ bays, updatedAt: new Date().toISOString() }, null, 2), 'utf8');
        res.json({ success: true, bays, swapped: tgtData.assigned });
    } catch (err) {
        console.error('[shopfloor] Move bay failed:', err.message);
        res.status(500).json({ success: false, error: err.message });
    }
});

// 3-2. MES Actual Performance Sync API
app.post('/api/mes-sync', async (req, res) => {
    try {
        let mesItems = [];
        const locBayMap = new Map();

        // 1. If client provided manual raw OUT_DATA or raw JSON directly
        if (Array.isArray(req.body.outData) && req.body.outData.length > 0) {
            mesItems = req.body.outData;
        } else if (req.body.rawJson) {
            let parsed = typeof req.body.rawJson === 'string' ? JSON.parse(req.body.rawJson) : req.body.rawJson;
            mesItems = Array.isArray(parsed) ? parsed : (parsed.OUT_DATA || (parsed.InDataList && parsed.InDataList.OUT_DATA) || []);
        } else {
            // 2. Fetch from remote MES API using DUAL Queries (Progress + Physical Bay Location)
            const authHeader = req.headers['authorization'] || req.body.token || '';
            const now = new Date();
            const yyyy = now.getFullYear();
            const mm = String(now.getMonth() + 1).padStart(2, '0');
            const nextMm = String((now.getMonth() + 2) > 12 ? 1 : (now.getMonth() + 2)).padStart(2, '0');
            const nextYyyy = (now.getMonth() + 2) > 12 ? (yyyy + 1) : yyyy;

            const fromYm = `${yyyy}${mm}`;
            const toYm = `${nextYyyy}${nextMm}`;

            // Query 1: Progress & Real-time Active Process (CUR_PROC_ID)
            const payloadProgress = {
                BizActId: "BR_DNS_MES_SEL_ProdProgressStatus_Assy",
                InDataList: {
                    IN_DATA: [
                        {
                            ENTERPRISE_ID: "1800",
                            PLANT_ID: "1840",
                            DEPT_ID: null,
                            PROD_YYYYMM_FROM: fromYm,
                            PROD_YYYYMM_TO: toYm,
                            ISFIRST: "FIRST",
                            ISINCLUDE_PLAN: "Y",
                            ISINCLUDE_START: "Y",
                            ISINCLUDE_COMPLETE_OPER: "N",
                            ISINCLUDE_COMPLETE: "Y",
                            INCLUDE_OTHERDEPT: "Y",
                            ISINCLUDE_PARALLEL: "Y",
                            LANG_ID: "ko-KR",
                            ORD_TYPE_CODE: "A"
                        }
                    ]
                },
                OutData: "OUT_DATA",
                port: "8082"
            };

            // Query 2: Workorder Process & Workers (CN0WBFAG - MC1/2조립)
            const payloadWorkorderG = {
                BizActId: "BR_DNS_MES_GET_WorkorderProcess",
                InDataList: {
                    IN_DATA: [
                        {
                            ENTERPRISE_ID: "1800",
                            PROD_ORD_ID: null,
                            PLANT_ID: "1840",
                            PROD_MDL_ID: null,
                            PROD_YYYYMM_FROM: fromYm,
                            PROD_YYYYMM_TO: toYm,
                            ORD_TYPE_CODE: "A",
                            MTRL_ID: null,
                            START_PLAN_DATE_FROM: null,
                            START_PLAN_DATE_TO: null,
                            DEPT_VENDOR_CHK: "Y",
                            DEPT_VENDOR_ID: "CN0WBFAG",
                            LANG_ID: "ko-KR",
                            EXCLD_PROC_END: "Y"
                        }
                    ]
                },
                OutData: "OUT_DATA,OUT_GPES",
                port: "8082"
            };

            // Query 3: Workorder Process & Workers (CN0WBFAH - MC3/4조립)
            const payloadWorkorderH = {
                BizActId: "BR_DNS_MES_GET_WorkorderProcess",
                InDataList: {
                    IN_DATA: [
                        {
                            ENTERPRISE_ID: "1800",
                            PROD_ORD_ID: null,
                            PLANT_ID: "1840",
                            PROD_MDL_ID: null,
                            PROD_YYYYMM_FROM: fromYm,
                            PROD_YYYYMM_TO: toYm,
                            ORD_TYPE_CODE: "A",
                            MTRL_ID: null,
                            START_PLAN_DATE_FROM: null,
                            START_PLAN_DATE_TO: null,
                            DEPT_VENDOR_CHK: "Y",
                            DEPT_VENDOR_ID: "CN0WBFAH",
                            LANG_ID: "ko-KR",
                            EXCLD_PROC_END: "Y"
                        }
                    ]
                },
                OutData: "OUT_DATA,OUT_GPES",
                port: "8082"
            };

            // Query 4: Physical Bay Locations (GI_LOC_ID) across all departments
            const payloadLocation = {
                BizActId: "BR_DNS_MES_GET_GIProdOrderOper",
                InDataList: {
                    IN_DATA: [
                        {
                            ENTERPRISE_ID: "1800",
                            PLANT_ID: "1840",
                            LANG_ID: "ko-KR",
                            PROD_ORD_ID: "",
                            PROD_MDL_ID: null,
                            PROD_YYYYMM_FROM: fromYm,
                            PROD_YYYYMM_TO: toYm,
                            ORD_TYPE_CODE: "A",
                            WC_ID: null,
                            MTRL_ID: "",
                            DEPT_VENDOR_CHK: null,
                            DEPT_VENDOR_ID: null,
                            EXCLD_GI_END: null
                        }
                    ]
                },
                OutData: "OUT_DATA,OUT_GPES",
                port: "8082"
            };

            const headers = { 'Content-Type': 'application/json' };
            if (authHeader) {
                headers['Authorization'] = authHeader.startsWith('Bearer ') ? authHeader : `Bearer ${authHeader}`;
            }

            const workerMap = new Map();

            try {
                const [resProgress, resWorkG, resWorkH, resLoc] = await Promise.all([
                    fetch('http://mes.dn-solutions.com:8081/api/json/query', { method: 'POST', headers, body: JSON.stringify(payloadProgress) }).catch(e => null),
                    fetch('http://mes.dn-solutions.com:8081/api/json/query', { method: 'POST', headers, body: JSON.stringify(payloadWorkorderG) }).catch(e => null),
                    fetch('http://mes.dn-solutions.com:8081/api/json/query', { method: 'POST', headers, body: JSON.stringify(payloadWorkorderH) }).catch(e => null),
                    fetch('http://mes.dn-solutions.com:8081/api/json/query', { method: 'POST', headers, body: JSON.stringify(payloadLocation) }).catch(e => null)
                ]);

                if (resProgress && resProgress.ok) {
                    const mesJson = await resProgress.json();
                    mesItems = mesJson.OUT_DATA || (mesJson.InDataList && mesJson.InDataList.OUT_DATA) || [];
                } else if (resProgress && resProgress.status === 401) {
                    return res.json({
                        success: false,
                        error: 'MES 인증 토큰(Bearer Token)이 만료되었거나 유효하지 않습니다. 최신 토큰을 설정해주세요.',
                        status: 401
                    });
                }

                // Process Workorder Process responses (Workers & Processes)
                const workorderRows = [];
                if (resWorkG && resWorkG.ok) {
                    try {
                        const gJson = await resWorkG.json();
                        const gRows = gJson.OUT_DATA || [];
                        const gGpes = gJson.OUT_GPES || [];
                        workorderRows.push(...gRows, ...gGpes);
                        console.log(`[MES] Fetched ${gRows.length} OUT_DATA + ${gGpes.length} OUT_GPES records from CN0WBFAG`);
                    } catch(e) {}
                }
                if (resWorkH && resWorkH.ok) {
                    try {
                        const hJson = await resWorkH.json();
                        const hRows = hJson.OUT_DATA || [];
                        const hGpes = hJson.OUT_GPES || [];
                        workorderRows.push(...hRows, ...hGpes);
                        console.log(`[MES] Fetched ${hRows.length} OUT_DATA + ${hGpes.length} OUT_GPES records from CN0WBFAH`);
                    } catch(e) {}
                }

                workorderRows.forEach(r => {
                    const ordKey = (r.PROD_ORD_ID || '').trim();
                    const serialKey = (r.PROD_MDL_CNT || '').trim();
                    
                    // Search all possible worker properties in MES WorkorderProcess row
                    let worker = '';
                    for (const [k, v] of Object.entries(r)) {
                        if (/WORKER|OPERATOR|USER_NAME|CHARGER|EMP_NAME|ACT_WRK/i.test(k) && typeof v === 'string' && v.trim() && !/ID|CODE/i.test(k)) {
                            worker = v.trim();
                            break;
                        }
                    }
                    if (!worker) {
                        worker = r.WORKER_NAME || r.WORKER || r.USER_NAME || r.OPERATOR_NAME || r.OPERATOR || r.CHARGER_NAME || r.CHARGER || r.EMP_NAME || r.ACT_WORKER_NAME || r.ACT_WORKER || '';
                    }

                    const loc = (r.GI_LOC_ID || r.LOC_ID || r.WORK_LOC_ID || '').trim().toUpperCase();

                    if (worker) {
                        if (ordKey) workerMap.set(ordKey, worker);
                        if (serialKey) workerMap.set(serialKey, worker);
                    }
                    if (loc) {
                        if (ordKey && !locBayMap.has(ordKey)) locBayMap.set(ordKey, { loc, area: r.GI_AREA_NAME || r.AREA_NAME, ver: r.PROD_VER_ID, raw: r });
                        if (serialKey && !locBayMap.has(serialKey)) locBayMap.set(serialKey, { loc, area: r.GI_AREA_NAME || r.AREA_NAME, ver: r.PROD_VER_ID, raw: r });
                    }
                });

                let locRows = [];
                if (resLoc && resLoc.ok) {
                    const locJson = await resLoc.json();
                    locRows = locJson.OUT_DATA || [];
                    console.log(`[MES] Fetched ${locRows.length} physical bay location records from Query 4`);

                    // Map bay locations by Order ID and Serial
                    locRows.forEach(r => {
                        const ordKey = (r.PROD_ORD_ID || '').trim();
                        const serialKey = (r.PROD_MDL_CNT || '').trim();
                        const loc = (r.GI_LOC_ID || '').trim().toUpperCase();

                        if (loc) {
                            if (ordKey && !locBayMap.has(ordKey)) locBayMap.set(ordKey, { loc, area: r.GI_AREA_NAME, ver: r.PROD_VER_ID, raw: r });
                            if (serialKey && !locBayMap.has(serialKey)) locBayMap.set(serialKey, { loc, area: r.GI_AREA_NAME, ver: r.PROD_VER_ID, raw: r });
                        }
                    });
                }

                // Debug save sample records
                try {
                    fs.writeFileSync('mes_debug_sample.json', JSON.stringify({
                        query1_sample: mesItems.slice(0, 2),
                        workorder_sample: workorderRows.slice(0, 3),
                        workorder_keys: workorderRows[0] ? Object.keys(workorderRows[0]) : [],
                        loc_sample: locRows.slice(0, 2)
                    }, null, 2), 'utf8');
                    console.log('[MES DEBUG] Saved mes_debug_sample.json with workorder rows');
                } catch(e) {}

            } catch (fetchErr) {
                console.warn('[MES] Remote connection warning:', fetchErr.message);
            }
        }

        if (mesItems.length === 0) {
            return res.json({ 
                success: false, 
                error: 'MES 데이터를 수신하지 못했습니다. (조회 조건 또는 로그인 토큰을 확인해주세요)',
                count: 0 
            });
        }

        // Load existing bays (and ensure full template coverage)
        let bays = [];
        if (fs.existsSync(DATA_FILE)) {
            try {
                const data = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
                bays = data.bays || [];
            } catch (e) {
                bays = [];
            }
        }
        const defaultBays = getDefaultBays();
        defaultBays.forEach(def => {
            if (!bays.some(b => b.id === def.id || (b.shift === def.shift && b.bay === def.bay))) {
                bays.push(def);
            }
        });

        // Strict MC Classifier
        function classifyMcShift(item) {
            const mdl = (item.PROD_MDL_NAME || item.MTRL_ID || '').toUpperCase();
            const wc = (item.WC_ID || '').toUpperCase();
            const ver = (item.PROD_VER_ID || '').toUpperCase();

            // 1. Strict Exclusion of TC/Lathe models
            if (mdl.includes('DNX') || mdl.includes('LYNX') || mdl.includes('PUMA') || mdl.includes('TW') || mdl.includes('TT') || mdl.includes('TL') || wc.includes('TC') || ver.includes('0AT')) {
                return null;
            }

            // 2. MC 4직: DHF 8000, NHP 8000, HM 1000, HM 1250, XC 4000
            if (mdl.includes('DHF') || mdl.includes('NHP 8') || mdl.includes('NHP8') || mdl.includes('HM 1') || mdl.includes('HM1') || mdl.includes('XC') || wc === 'A10MC40' || /0AM4|0AMD/.test(ver)) {
                return 'MC4직';
            }

            // 3. MC 3직: NHM 5000/6300/8000, NHP 5500/6300
            if (mdl.includes('NHM') || mdl.includes('NHP 55') || mdl.includes('NHP55') || mdl.includes('NHP 63') || mdl.includes('NHP63') || wc === 'A10MC30' || /0AM3|0AMC/.test(ver)) {
                return 'MC3직';
            }

            // 4. MC 2직: DVF 4000/5000/6500/8000 (5축기)
            if (mdl.includes('DVF') || wc === 'A10MC20' || ver.includes('0AM2')) {
                return 'MC2직';
            }

            // 5. MC 1직: NHP 4000/5000, NHC 4000/5000, HC 400/500
            if (mdl.includes('NHP 4') || mdl.includes('NHP4') || mdl.includes('NHP 50') || mdl.includes('NHP50') || mdl.includes('NHC') || mdl.includes('HC 4') || mdl.includes('HC 5') || wc === 'A10MC10' || /0AM1|0AMA|0AMB/.test(ver)) {
                return 'MC1직';
            }

            // Other MC lines
            if (wc === 'A10MCE0' || wc === 'Q10MC00' || wc === 'A10MC51') {
                if (mdl.includes('DVF')) return 'MC2직';
                if (mdl.includes('NHM')) return 'MC3직';
                if (mdl.includes('DHF') || mdl.includes('HM')) return 'MC4직';
                if (mdl.includes('NHP') || mdl.includes('NHC') || mdl.includes('HC')) return 'MC1직';
            }

            return null;
        }

        // Group unique machines from MES & Attach physical Bay Location (GI_LOC_ID)
        const machineMap = new Map();
        mesItems.forEach(i => {
            const shift = classifyMcShift(i);
            if (!shift) return; // Skip TC/non-MC

            const serial = i.PROD_MDL_ID ? `${i.PROD_MDL_ID}-${i.PROD_MDL_CNT}` : (i.PROD_MDL_CNT || '');
            const ordKey = (i.PROD_ORD_ID || '').trim();
            const serialKey = (i.PROD_MDL_CNT || '').trim();
            const key = ordKey || serial;

            const locInfo = (ordKey && locBayMap.get(ordKey)) || (serialKey && locBayMap.get(serialKey)) || {};
            const workerFound = (ordKey && workerMap.get(ordKey)) || (serialKey && workerMap.get(serialKey)) || i.WORKER_NAME || i.WORKER || i.USER_NAME || i.CHARGER || '';

            if (!machineMap.has(key) || (i.CUR_PROC_ID && !machineMap.get(key).CUR_PROC_ID)) {
                machineMap.set(key, {
                    ...i,
                    shift,
                    serial,
                    model: i.PROD_MDL_NAME || i.MTRL_ID,
                    salesDoc: i.PROD_ORD_ID,
                    loc: locInfo.loc || '',
                    area: locInfo.area || '',
                    worker: workerFound
                });
            }
        });

        const mesMachines = [...machineMap.values()];

        // Clean Bay code extractor (e.g. "D17", "D1(C동)" -> "D17", "D1")
        function extractBayCode(str) {
            if (!str) return '';
            const m = str.toString().toUpperCase().match(/([A-F]\d{1,2})/);
            return m ? m[1] : str.toString().toUpperCase().replace(/[^A-Z0-9]/g, '');
        }

        // Smart Physical Bay Allocation
        let autoAssignedCount = 0;
        let updatedCount = 0;
        let shippedCount = 0;

        bays.forEach(targetBay => {
            const targetBayCode = extractBayCode(targetBay.bay);
            const targetShift = targetBay.shift;
            const targetArea = targetBay.area;
            
            // Look for an MES machine assigned to this exact physical bay location and shift
            const locMatch = mesMachines.find(m => {
                if (!m.loc) return false;
                const mCode = extractBayCode(m.loc);
                if (mCode !== targetBayCode) return false;
                
                // Match shift
                if (m.shift && m.shift !== targetShift) return false;

                // Match area if both specified
                if (m.area && targetArea && m.area.includes(targetArea.substring(0, 1))) return true;
                return true;
            });

            if (locMatch) {
                targetBay.assigned = true;
                targetBay.model = locMatch.model;
                targetBay.serial = locMatch.serial;
                targetBay.salesDoc = locMatch.salesDoc || '';
                targetBay.customer = locMatch.customer || '';
                targetBay.currentProcess = locMatch.CUR_PROC_ID || locMatch.PROC_ID || 'BASE';
                targetBay.startDate = locMatch.START_PLAN_DATE || targetBay.startDate;
                targetBay.deliveryDate = locMatch.SHIP_TARGET_DATE || targetBay.deliveryDate;
                targetBay.spec = locMatch.MTRL_ID || targetBay.spec;
                
                // Extract worker name
                const ordKey = (locMatch.salesDoc || '').trim();
                const serialKey = (locMatch.PROD_MDL_CNT || locMatch.serial || '').trim();
                const workerFound = locMatch.worker || (ordKey && workerMap.get(ordKey)) || (serialKey && workerMap.get(serialKey)) || locMatch.WORKER_NAME || locMatch.WORKER || locMatch.USER_NAME || locMatch.CHARGER || '';
                if (workerFound) targetBay.worker = workerFound;

                targetBay.source = 'MES';
                targetBay.isShipped = false;
                if (locMatch.PROD_ORD_STATUS_NAME) {
                    targetBay.issue = `[상태: ${locMatch.PROD_ORD_STATUS_NAME}]`;
                }
                autoAssignedCount++;
                return;
            }

            // 2. If bay was already assigned manually, sync its in-progress stage
            if (targetBay.assigned) {
                const baySerial = (targetBay.serial || '').trim();
                const bayOrder = (targetBay.salesDoc || '').trim();

                const match = mesMachines.find(m => {
                    if (baySerial && (m.serial === baySerial || m.PROD_MDL_CNT === baySerial)) return true;
                    if (bayOrder && m.salesDoc === bayOrder) return true;
                    return false;
                });

                if (match) {
                    targetBay.currentProcess = match.CUR_PROC_ID || match.PROC_ID || targetBay.currentProcess || 'BASE';
                    targetBay.startDate = match.START_PLAN_DATE || targetBay.startDate;
                    targetBay.deliveryDate = match.SHIP_TARGET_DATE || targetBay.deliveryDate;
                    targetBay.spec = match.MTRL_ID || targetBay.spec;

                    const ordKey = (match.salesDoc || '').trim();
                    const serialKey = (match.PROD_MDL_CNT || match.serial || '').trim();
                    const workerFound = match.worker || (ordKey && workerMap.get(ordKey)) || (serialKey && workerMap.get(serialKey)) || match.WORKER_NAME || match.WORKER || match.USER_NAME || match.CHARGER || '';
                    if (workerFound) targetBay.worker = workerFound;

                    targetBay.source = 'MES';
                    targetBay.isShipped = false;
                    if (match.PROD_ORD_STATUS_NAME) {
                        targetBay.issue = `[상태: ${match.PROD_ORD_STATUS_NAME}]`;
                    }
                    updatedCount++;
                } else if (targetBay.source === 'MES') {
                    // Shipped
                    targetBay.isShipped = true;
                    targetBay.issue = '🏁 조립 및 출하 완료 (MES 종료)';
                    shippedCount++;
                }
            }
        });

        // Save updated bays
        fs.writeFileSync(DATA_FILE, JSON.stringify({ bays, updatedAt: new Date().toISOString() }, null, 2), 'utf8');
        console.log(`[MES] Synced: ${autoAssignedCount} bays auto-placed from GI_LOC_ID, ${updatedCount} bays stage-updated, ${shippedCount} shipped`);

        res.json({
            success: true,
            totalMesRecords: mesItems.length,
            validMcMachinesCount: mesMachines.length,
            autoAssignedBaysCount: autoAssignedCount,
            updatedBaysCount: updatedCount,
            shippedBaysCount: shippedCount,
            mesMachines,
            bays
        });
    } catch (err) {
        console.error('[shopfloor] MES sync failed:', err.message);
        res.status(500).json({ success: false, error: err.message });
    }
});

// 4. Fetch MPS Planned Machines from local SAP files
app.get('/api/mps-plan-machines', (req, res) => {
    try {
        const sapPath = path.join(__dirname, 'sap_1840.mhtml');
        if (!fs.existsSync(sapPath)) {
            return res.json({ success: true, machines: [] });
        }

        const mhtml = fs.readFileSync(sapPath, 'utf8');
        const decoded = mhtml
            .replace(/=\r?\n/g, '')
            .replace(/=([0-9A-F]{2})/gi, (_, hex) => String.fromCharCode(parseInt(hex, 16)));

        const trMatches = decoded.match(/<tr[\s\S]*?<\/tr>/gi) || [];
        const rows = trMatches.map(tr => {
            const tdMatches = tr.match(/<t[dh][\s\S]*?<\/t[dh]>/gi) || [];
            return tdMatches.map(td => td.replace(/<[^>]+>/g, '').trim());
        });

        if (rows.length < 2) {
            return res.json({ success: true, machines: [] });
        }

        let headerRow = rows[0];
        let headerCells = headerRow.map(c => c.toUpperCase());
        for (let i = 1; i < Math.min(10, rows.length); i++) {
            const cells = rows[i].map(c => c.toUpperCase());
            if (cells.some(c => c.includes('MATERIAL') || c.includes('PROD'))) {
                headerRow = rows[i];
                headerCells = cells;
                break;
            }
        }

        const findIdx = (keywords) => {
            let idx = headerCells.findIndex(cell => keywords.some(k => cell === k));
            if (idx !== -1) return idx;
            return headerCells.findIndex(cell => keywords.some(k => cell.includes(k)));
        };

        const map = {
            mon: findIdx(['PROD.MON', 'MONTH', '생산월']),
            model: findIdx(['MATERIAL DESCRIPTION', 'MODEL', '기종']),
            customer: findIdx(['CUSTOMER NAME', 'CUSTOMER', '고객']),
            serial: findIdx(['SERIAL NO', 'S/O SERIAL', '시리얼', 'SERIAL']),
            order: findIdx(['ORDER', '오더']),
            salesDoc: findIdx(['S/O ORDER', 'SALES DOC', 'SALESDOC', '판매문서']),
            startDate: findIdx(['START DATE', '시작일']),
            ver: findIdx(['VER.', 'VERSION', '버전']),
            deletedItem: findIdx(['DELETED ITEM', 'DELETED', '삭제'])
        };

        const machines = [];
        const startIndex = rows.indexOf(headerRow) + 1;

        for (let i = startIndex; i < rows.length; i++) {
            const cells = rows[i];
            if (cells.length < 5) continue;
            if (map.deletedItem !== -1 && (cells[map.deletedItem] || '').toUpperCase() === 'X') continue;

            const ver = (cells[map.ver] || '').toUpperCase();
            let shift = null;
            if (/0AM1|0AMA|0AMB|MC1|MC 1/.test(ver)) {
                shift = 'MC1직';
            } else if (/0AM2|MC2|MC 2/.test(ver)) {
                shift = 'MC2직';
            } else if (/0AM3|0AMC|MC3|MC 3/.test(ver)) {
                shift = 'MC3직';
            } else if (/0AM4|0AMD|MC4|MC 4/.test(ver)) {
                shift = 'MC4직';
            }

            // MC 1~4직이 아닌 타 라인(TC, VC 등) 기종은 완전 제외
            if (!shift) continue;

            const monVal = (cells[map.mon] || '').toString().trim();
            let mNum = null;
            const dateDigits = monVal.replace(/[^0-9]/g, '');
            if (dateDigits.length >= 6) mNum = parseInt(dateDigits.substring(4, 6));
            else if (dateDigits.length >= 1) mNum = parseInt(dateDigits);

            machines.push({
                month: mNum ? String(mNum).padStart(2, '0') + '월' : monVal,
                model: cells[map.model] || '',
                serial: cells[map.serial] || '',
                salesDoc: cells[map.salesDoc] || '',
                orderNum: cells[map.order] || '',
                customer: cells[map.customer] || '',
                startDate: cells[map.startDate] || '',
                shift
            });
        }

        res.json({ success: true, machines });
    } catch (err) {
        console.error('[shopfloor] MPS plan fetch failed:', err.message);
        res.status(500).json({ success: false, error: err.message });
    }
});

// Heartbeat & Auto-shutdown (10 min idle)
let lastHeartbeat = Date.now();
let hasReceivedHeartbeat = false;

app.post('/api/heartbeat', (req, res) => {
    lastHeartbeat = Date.now();
    hasReceivedHeartbeat = true;
    res.sendStatus(200);
});

app.post('/api/shutdown', (req, res) => {
    console.log('[shopfloor] Shutdown requested.');
    res.json({ success: true });
    setTimeout(() => process.exit(0), 1000);
});

const server = app.listen(PORT, '0.0.0.0', () => {
    console.log(`=======================================================`);
    console.log(`  남산 MC 조립 지번(Bay) 독립 관리 서버 LIVE on Port ${PORT}`);
    console.log(`  접속 URL: http://localhost:${PORT}`);
    console.log(`=======================================================`);
});

const HEARTBEAT_TIMEOUT = 600000; // 10 minutes
const GRACE_PERIOD = 180000; // 3 minutes grace period
const startupTime = Date.now();

setInterval(() => {
    const now = Date.now();
    if (hasReceivedHeartbeat) {
        if (now - lastHeartbeat > HEARTBEAT_TIMEOUT) {
            console.log('[AUTO-SHUTDOWN] No client connected for 10 minutes. Exiting...');
            process.exit(0);
        }
    } else {
        if (now - startupTime > GRACE_PERIOD) {
            console.log('[AUTO-SHUTDOWN] Grace period expired without client connection. Exiting...');
            process.exit(0);
        }
    }
}, 5000);
