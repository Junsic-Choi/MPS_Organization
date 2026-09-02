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
            const baysList = data.bays || [];
            baysList.forEach(b => {
                if (b.source === 'MES' && Array.isArray(b.routingSteps) && b.routingSteps.length > 0) {
                    const active = b.routingSteps.find(s => s.status === 'START' || s.status === 'RUN' || s.status === 'CONT');
                    if (active && active.code) {
                        b.currentProcess = active.code;
                    }
                }
            });
            return res.json({ success: true, bays: baysList });
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
        const workerMap = new Map();
        const orderRoutingMap = new Map();

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

            // Broad window: 2 months prior (for active machines from previous months like 543호기) to 2 months ahead
            const prevDate = new Date(now.getFullYear(), now.getMonth() - 2, 1);
            const fromYm = `${prevDate.getFullYear()}${String(prevDate.getMonth() + 1).padStart(2, '0')}`;
            const nextDate = new Date(now.getFullYear(), now.getMonth() + 2, 1);
            const toYm = `${nextDate.getFullYear()}${String(nextDate.getMonth() + 1).padStart(2, '0')}`;

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

            // Query 2: Daily Work Result (CN0WBFAG - MC1조립)
            const dd = String(now.getDate()).padStart(2, '0');
            const todayStart = `${yyyy}-${mm}-${dd}T00:00:00.000+09:00`;
            const todayEnd = `${yyyy}-${mm}-${dd}T23:59:59.000+09:00`;

            const createDailyPayload = (deptId) => ({
                BizActId: "BR_DNS_MES_SEL_DailyWorkResult",
                InDataList: {
                    IN_DATA: [
                        {
                            LANG_ID: "ko-KR",
                            ENTERPRISE_ID: "1800",
                            PLANT_ID: "1840",
                            DEPT_ID: deptId,
                            WORKER_ID: null,
                            WO_TYPE_CODE: null,
                            WC_ID: null,
                            SEARCH_DATE_FROM: todayStart,
                            SEARCH_DATE_TO: todayEnd,
                            INCLUDE_NORESULT: "N"
                        }
                    ]
                },
                OutData: "OUT_DATA",
                port: "8082"
            });

            const payloadDailyG = createDailyPayload("CN0WBFAG"); // MC 1직
            const payloadDailyH = createDailyPayload("CN0WBFAH"); // MC 2직
            const payloadDailyI = createDailyPayload("CN0WBFAI"); // MC 3직
            const payloadDailyJ = createDailyPayload("CN0WBFAJ"); // MC 4직

            const createWorkorderPayload = (deptId) => ({
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
                            DEPT_VENDOR_ID: deptId,
                            LANG_ID: "ko-KR",
                            EXCLD_PROC_END: "N"
                        }
                    ]
                },
                OutData: "OUT_DATA,OUT_GPES",
                port: "8082"
            });

            const payloadWorkorderG = createWorkorderPayload("CN0WBFAG"); // MC 1직
            const payloadWorkorderH = createWorkorderPayload("CN0WBFAH"); // MC 2직
            const payloadWorkorderI = createWorkorderPayload("CN0WBFAI"); // MC 3직
            const payloadWorkorderJ = createWorkorderPayload("CN0WBFAJ"); // MC 4직

            // Query 6: Physical Bay Locations (GI_LOC_ID) across all departments
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

            // Query 7: QMS Inspection Request Status Update Procedure Trigger
            const payloadQmsUpdate = {
                BizActId: "BR_QMS_UPD_PR_REQUEST_INSPECTION_STATUS",
                InDataList: {
                    IN_DATA: [
                        {
                            ENTERPRISE_ID: "1800",
                            PLANT_ID: "1840"
                        }
                    ]
                },
                OutData: "",
                port: "8082"
            };

            // Query 8: QMS Inspection Requests Table (GIN Inspection Bays - Highest Source of Truth)
            const payloadQmsInspTbl = {
                BizActId: "BR_QMS_GET_REQUEST_INSPECTION_TBL",
                InDataList: {
                    IN_DATA: [
                        {
                            LANG_ID: "ko-KR",
                            ENTERPRISE_ID: "1800",
                            PLANT_ID: "1840",
                            LOGIN_USER_ID: "i0215099",
                            DEPT_ID: "",
                            REQ_DEPT_ID: "",
                            PROD_MDL_ID: "",
                            PROD_MDL_CNT: null,
                            INSP_TYPE: "",
                            INSP_STATUS: "",
                            PROD_YYYYMM_FROM: fromYm,
                            PROD_YYYYMM_TO: toYm,
                            INSP_DATE_CHK: false,
                            INSP_DTTM_FROM: "",
                            INSP_DTTM_TO: "",
                            OINS_INSP_DATE_CHK: false,
                            OINS_INSP_DTTM_FROM: "",
                            OINS_INSP_DTTM_TO: "",
                            isPending: false,
                            WO_DATE_CHK: false,
                            WO_DATE: ""
                        }
                    ]
                },
                OutData: "OUT_DATA",
                port: "8082"
            };

            const headers = { 'Content-Type': 'application/json' };
            if (authHeader) {
                headers['Authorization'] = authHeader.startsWith('Bearer ') ? authHeader : `Bearer ${authHeader}`;
            }

            try {
                const [
                    resProgress,
                    resDailyG, resDailyH, resDailyI, resDailyJ,
                    resWorkG, resWorkH, resWorkI, resWorkJ,
                    resLoc,
                    resQmsUpdate, resQms
                ] = await Promise.all([
                    fetch('http://mes.dn-solutions.com:8081/api/json/query', { method: 'POST', headers, body: JSON.stringify(payloadProgress) }).catch(e => null),
                    fetch('http://mes.dn-solutions.com:8081/api/json/query', { method: 'POST', headers, body: JSON.stringify(payloadDailyG) }).catch(e => null),
                    fetch('http://mes.dn-solutions.com:8081/api/json/query', { method: 'POST', headers, body: JSON.stringify(payloadDailyH) }).catch(e => null),
                    fetch('http://mes.dn-solutions.com:8081/api/json/query', { method: 'POST', headers, body: JSON.stringify(payloadDailyI) }).catch(e => null),
                    fetch('http://mes.dn-solutions.com:8081/api/json/query', { method: 'POST', headers, body: JSON.stringify(payloadDailyJ) }).catch(e => null),
                    fetch('http://mes.dn-solutions.com:8081/api/json/query', { method: 'POST', headers, body: JSON.stringify(payloadWorkorderG) }).catch(e => null),
                    fetch('http://mes.dn-solutions.com:8081/api/json/query', { method: 'POST', headers, body: JSON.stringify(payloadWorkorderH) }).catch(e => null),
                    fetch('http://mes.dn-solutions.com:8081/api/json/query', { method: 'POST', headers, body: JSON.stringify(payloadWorkorderI) }).catch(e => null),
                    fetch('http://mes.dn-solutions.com:8081/api/json/query', { method: 'POST', headers, body: JSON.stringify(payloadWorkorderJ) }).catch(e => null),
                    fetch('http://mes.dn-solutions.com:8081/api/json/query', { method: 'POST', headers, body: JSON.stringify(payloadLocation) }).catch(e => null),
                    fetch('http://mes.dn-solutions.com:8081/api/json/query', { method: 'POST', headers, body: JSON.stringify(payloadQmsUpdate) }).catch(e => null),
                    fetch('http://mes.dn-solutions.com:8081/api/json/query', { method: 'POST', headers, body: JSON.stringify(payloadQmsInspTbl) }).catch(e => null)
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

                // Process QMS Inspection Table (GIN Inspection Location - Highest Priority Ground Truth)
                const qmsRows = [];
                if (resQms && resQms.ok) {
                    try {
                        const qmsJson = await resQms.json();
                        qmsRows.push(...(qmsJson.OUT_DATA || []));
                        console.log(`[QMS] Fetched ${qmsRows.length} Inspection Request records`);
                    } catch(e) {}
                }

                qmsRows.forEach(r => {
                    const ordKey = (r.PROD_ORD_ID || r.WO_ID || r.ORDER_ID || '').trim();
                    const serialKey = (r.PROD_MDL_CNT || r.SERIAL_NO || r.SERIAL || '').trim();
                    const fullSerial = (r.FULL_PROD_MDL_CNT || (r.PROD_MDL_ID ? `${r.PROD_MDL_ID}-${r.PROD_MDL_CNT}` : r.PROD_MDL_CNT) || '').trim();

                    let loc = (r.GI_LOC_ID || r.INSP_LOC_ID || r.REQ_LOC_ID || r.LOC_ID || r.BAY || '').trim().toUpperCase();
                    if (!loc) {
                        for (const [k, v] of Object.entries(r)) {
                            if (typeof v === 'string' && /([A-F]\d{1,2})/.test(v)) {
                                const m = v.toUpperCase().match(/([A-F]\d{1,2})/);
                                if (m) { loc = m[1]; break; }
                            }
                        }
                    }

                    const worker = (r.REQ_WORKER_NAME || r.INSP_CHARGER_NAME || r.CHARGER || r.USER_NAME || r.WORKER_NAME || '').trim();
                    if (worker) {
                        if (ordKey) workerMap.set(ordKey, worker);
                        if (serialKey) workerMap.set(serialKey, worker);
                        if (fullSerial) workerMap.set(fullSerial, worker);
                    }

                    if (loc) {
                        if (ordKey) locBayMap.set(ordKey, { loc, area: r.AREA_NAME || r.GI_AREA_NAME, isQms: true, isInsp: true, isWorkorder: true, worker, raw: r });
                        if (serialKey) locBayMap.set(serialKey, { loc, area: r.AREA_NAME || r.GI_AREA_NAME, isQms: true, isInsp: true, isWorkorder: true, worker, raw: r });
                        if (fullSerial) locBayMap.set(fullSerial, { loc, area: r.AREA_NAME || r.GI_AREA_NAME, isQms: true, isInsp: true, isWorkorder: true, worker, raw: r });
                    }
                });

                // Process Daily Work Results (Floor Activity Ground Truth across 4 Shifts)
                const dailyRows = [];
                const dailyResponses = [
                    { res: resDailyG, dept: 'CN0WBFAG' },
                    { res: resDailyH, dept: 'CN0WBFAH' },
                    { res: resDailyI, dept: 'CN0WBFAI' },
                    { res: resDailyJ, dept: 'CN0WBFAJ' }
                ];

                for (const d of dailyResponses) {
                    if (d.res && d.res.ok) {
                        try {
                            const dJson = await d.res.json();
                            const rows = dJson.OUT_DATA || [];
                            dailyRows.push(...rows);
                            console.log(`[MES] Fetched ${rows.length} DailyWorkResult records from ${d.dept}`);
                        } catch(e) {}
                    }
                }

                dailyRows.forEach(r => {
                    const ordKey = (r.PROD_ORD_ID || '').trim();
                    const serialKey = (r.PROD_MDL_CNT || '').trim();
                    const fullSerial = (r.FULL_PROD_MDL_CNT || (r.PROD_MDL_ID ? `${r.PROD_MDL_ID}-${r.PROD_MDL_CNT}` : r.PROD_MDL_CNT) || '').trim();
                    
                    let worker = '';
                    for (const [k, v] of Object.entries(r)) {
                        if (/WORKER|OPERATOR|USER_NAME|CHARGER|EMP_NAME|ACT_WRK/i.test(k) && typeof v === 'string' && v.trim() && !/ID|CODE/i.test(k)) {
                            worker = v.trim();
                            break;
                        }
                    }
                    if (!worker) worker = r.WORKER_NAME || r.WORKER || r.USER_NAME || r.OPERATOR_NAME || '';

                    // Check all possible location fields
                    let loc = (r.GI_LOC_ID || r.WORK_LOC_ID || r.LOC_ID || '').trim().toUpperCase();
                    if (!loc) {
                        for (const [k, v] of Object.entries(r)) {
                            if (typeof v === 'string' && /([A-F]\d{1,2})/.test(v)) {
                                const m = v.toUpperCase().match(/([A-F]\d{1,2})/);
                                if (m) { loc = m[1]; break; }
                            }
                        }
                    }

                    if (worker) {
                        if (ordKey && !workerMap.has(ordKey)) workerMap.set(ordKey, worker);
                        if (serialKey && !workerMap.has(serialKey)) workerMap.set(serialKey, worker);
                        if (fullSerial && !workerMap.has(fullSerial)) workerMap.set(fullSerial, worker);
                    }

                    if (loc) {
                        const existing = locBayMap.get(ordKey) || locBayMap.get(serialKey);
                        if (!existing || !existing.isQms) {
                            if (ordKey) locBayMap.set(ordKey, { loc, area: r.GI_AREA_NAME || r.AREA_NAME, ver: r.PROD_VER_ID, isDaily: true, worker, raw: r });
                            if (serialKey) locBayMap.set(serialKey, { loc, area: r.GI_AREA_NAME || r.AREA_NAME, ver: r.PROD_VER_ID, isDaily: true, worker, raw: r });
                        }
                    }
                });

                // Process Workorder Process responses (Workers & Processes across 4 Shifts)
                const workorderRows = [];
                const workResponses = [
                    { res: resWorkG, dept: 'CN0WBFAG' },
                    { res: resWorkH, dept: 'CN0WBFAH' },
                    { res: resWorkI, dept: 'CN0WBFAI' },
                    { res: resWorkJ, dept: 'CN0WBFAJ' }
                ];

                for (const w of workResponses) {
                    if (w.res && w.res.ok) {
                        try {
                            const wJson = await w.res.json();
                            const wRows = wJson.OUT_DATA || [];
                            const wGpes = wJson.OUT_GPES || [];
                            workorderRows.push(...wRows, ...wGpes);
                            console.log(`[MES] Fetched ${wRows.length + wGpes.length} WorkorderProcess records from ${w.dept}`);
                        } catch(e) {}
                    }
                }

                // Helper: Process Code to Friendly Descriptive Korean Name
                function getProcFriendlyName(procId) {
                    const p = (procId || '').toUpperCase();
                    if (p.includes('BDLEVEL') || p.includes('BASE')) return 'Base 레벨링 & 베드 안착 (GBDLEVEL)';
                    if (p === 'PGEB') return '테이블 안착 & 서보모터 (PGEB)';
                    if (p.startsWith('PCLM')) return '컬럼 조립 & 안착 (PCLM)';
                    if (p.startsWith('PACONF')) return '팔레트/부속 가공 & 서브 안착 (PACONF)';
                    if (p.includes('ATC') && (p.includes('30T') || p.includes('40T') || p.includes('TOOL') || p.includes('MAG'))) return `ATC 매거진 조립 (${p}) ⚡`;
                    if (p.includes('SCALE')) return `스케일 조립 & 에어 배관 (${p}) ⚡`;
                    if (p.startsWith('RPS')) return `Round Pallet System (${p}) 장착 & 시운전 ⚡`;
                    if (p.startsWith('APC') || p.startsWith('2PAL')) return `APC 팔레트 체인저 장착 & 시운전 ⚡`;
                    if (p.startsWith('MAT')) return `Matrix Magazine (${p}) 툴 매거진 장착 ⚡`;
                    if (p === 'EIF' || p === 'EIF1' || p === 'EIF2') return '전장 인터페이스 & 결선 (EIF)';
                    if (p === 'EAD') return '전장 시운전 & FSSB 파라미터 (EAD)';
                    if (p.includes('TRANS') || p.includes('XYHOME') || p.includes('XYZHOME')) return '원점 셋팅 & 팔레트 안착 (TRANS/XYHOME)';
                    if (p.startsWith('GAJ') || p.includes('LASER') || p.includes('ATC') || p.includes('ACC')) return `정밀도 측정 & 레이저 보정 (${p})`;
                    if (p.includes('SELFCUTT')) return 'ATC 센터링 & 셀프컷팅 시험 (SELFCUTT)';
                    if (p.startsWith('GIN')) return `기능 검사 & 쿨런트/정도 검사 (${p})`;
                    if (p.startsWith('GSG')) return 'Splash Guard 외주 조립 (GSG)';
                    if (p.startsWith('GSR') || p.startsWith('GSRAM')) return '연속 무부하 가동 시험 & 도어 조정 (GSR)';
                    if (p.startsWith('GRN') || p.startsWith('GRNE') || p.startsWith('GRNT')) return '절삭유 탱크 & 옵션 배관 작업 (GRN)';
                    if (p.startsWith('CSWRITE')) return 'CS 소프트웨어 라이팅 & 셋팅 (CSWRITE)';
                    if (p.startsWith('GOT')) return '출하 검사 & 최종 출하 준비 (GOT)';
                    if (p.startsWith('OINS') || p.startsWith('LK') || p.includes('출검')) return '최종 출하검사 & 축고정 포장 (OINS/LK)';
                    return `[${p}] 실시간 조립 공정`;
                }

                // Map of full actual routing steps per order / serial
                orderRoutingMap.clear();

                workorderRows.forEach(r => {
                    const ordKey = (r.PROD_ORD_ID || '').trim();
                    const serialKey = (r.PROD_MDL_CNT || '').trim();
                    const fullSerial = (r.FULL_PROD_MDL_CNT || '').trim();
                    const procId = (r.PROC_ID || '').trim().toUpperCase();
                    const procSeq = (r.PROC_SEQ || '').trim();
                    const stdTime = r.STD_TIME || 0;
                    const status = r.PROC_STATUS_CODE || r.LOT_STATUS_CODE || '';
                    
                    if (procId && (ordKey || serialKey || fullSerial)) {
                        const step = {
                            code: procId,
                            seq: procSeq,
                            stdTime: stdTime,
                            status: status
                        };

                        const keys = [ordKey, serialKey, fullSerial].filter(Boolean);
                        keys.forEach(k => {
                            if (!orderRoutingMap.has(k)) orderRoutingMap.set(k, []);
                            const list = orderRoutingMap.get(k);
                            if (!list.some(s => s.code === procId && s.seq === procSeq)) {
                                list.push(step);
                            }
                        });
                    }

                    let worker = '';
                    for (const [k, v] of Object.entries(r)) {
                        if (/WORKER|OPERATOR|USER_NAME|CHARGER|EMP_NAME|ACT_WRK/i.test(k) && typeof v === 'string' && v.trim() && !/ID|CODE/i.test(k)) {
                            worker = v.trim();
                            break;
                        }
                    }
                    if (!worker) worker = r.WORKER_NAME || r.WORKER || r.USER_NAME || r.OPERATOR_NAME || '';

                    const loc = (r.GI_LOC_ID || r.LOC_ID || r.WORK_LOC_ID || '').trim().toUpperCase();

                    if (worker) {
                        if (ordKey) workerMap.set(ordKey, worker);
                        if (serialKey) workerMap.set(serialKey, worker);
                        if (fullSerial) workerMap.set(fullSerial, worker);
                    }
                    if (loc) {
                        // MC 조립(CN0WBFAG, CN0WBFAH)에서 직접 지정된 최신 지번이므로 최우선 보존
                        if (ordKey && (!locBayMap.has(ordKey) || !locBayMap.get(ordKey).isDaily)) {
                            locBayMap.set(ordKey, { loc, area: r.GI_AREA_NAME || r.AREA_NAME, ver: r.PROD_VER_ID, isWorkorder: true, raw: r });
                        }
                        if (serialKey && (!locBayMap.has(serialKey) || !locBayMap.get(serialKey).isDaily)) {
                            locBayMap.set(serialKey, { loc, area: r.GI_AREA_NAME || r.AREA_NAME, ver: r.PROD_VER_ID, isWorkorder: true, raw: r });
                        }
                    }
                });

                // Format & sort each order's actual full routing
                for (const [k, list] of orderRoutingMap.entries()) {
                    list.sort((a, b) => (parseInt(a.seq) || 0) - (parseInt(b.seq) || 0));
                    list.forEach((s, idx) => {
                        s.dayIndex = idx + 1;
                        s.name = `${idx + 1}일차 / ${s.code} | ${getProcFriendlyName(s.code)}`;
                    });
                }

                let locRows = [];
                if (resLoc && resLoc.ok) {
                    const locJson = await resLoc.json();
                    locRows = locJson.OUT_DATA || [];

                    // Map bay locations by Order ID and Serial ONLY IF not already assigned by Workorder/Daily
                    locRows.forEach(r => {
                        const ordKey = (r.PROD_ORD_ID || '').trim();
                        const serialKey = (r.PROD_MDL_CNT || '').trim();
                        const loc = (r.GI_LOC_ID || '').trim().toUpperCase();

                        if (loc) {
                            if (ordKey && !locBayMap.has(ordKey)) {
                                locBayMap.set(ordKey, { loc, area: r.GI_AREA_NAME, ver: r.PROD_VER_ID, raw: r });
                            }
                            if (serialKey && !locBayMap.has(serialKey)) {
                                locBayMap.set(serialKey, { loc, area: r.GI_AREA_NAME, ver: r.PROD_VER_ID, raw: r });
                            }
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
            const dept = (item.DEPT_ID || item.DEPT_VENDOR_ID || '').toUpperCase();
            if (dept === 'CN0WBFAG') return 'MC1직';
            if (dept === 'CN0WBFAH') return 'MC2직';
            if (dept === 'CN0WBFAI') return 'MC3직';
            if (dept === 'CN0WBFAJ') return 'MC4직';

            const mdl = (item.PROD_MDL_NAME || item.MTRL_ID || '').toUpperCase();
            const wc = (item.WC_ID || '').toUpperCase();
            const ver = (item.PROD_VER_ID || '').toUpperCase();

            // 1. Strict Exclusion of TC/Lathe models
            if (mdl.includes('DNX') || mdl.includes('LYNX') || mdl.includes('PUMA') || mdl.includes('TW') || mdl.includes('TT') || mdl.includes('TL') || wc.includes('TC') || ver.includes('0AT')) {
                return null;
            }

            // 2. MC 4직: DHF 8000, NHP 8000, NHP 800, HM 1000, HM 1250, XC 4000
            if (mdl.includes('DHF') || mdl.includes('NHP 8') || mdl.includes('NHP8') || mdl.includes('NHP 800') || mdl.includes('NHP800') || mdl.includes('HM 1') || mdl.includes('HM1') || mdl.includes('XC') || wc === 'A10MC40' || /0AM4|0AMD/.test(ver)) {
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
                if (mdl.includes('DHF') || mdl.includes('HM') || mdl.includes('NHP 8') || mdl.includes('NHP8') || mdl.includes('NHP800') || mdl.includes('XC') || /0AM4|0AMD/.test(ver)) return 'MC4직';
                if (mdl.includes('NHP') || mdl.includes('NHC') || mdl.includes('HC')) return 'MC1직';
            }

            return null;
        }

        // Helper: Parse SAP 1840 production schedule to map customer and planMonth
        function loadSapMetaMap() {
            const customerMap = new Map();
            const monthMap = new Map();
            try {
                const sapPath = path.join(__dirname, 'sap_1840.mhtml');
                if (!fs.existsSync(sapPath)) return { customerMap, monthMap };

                const mhtml = fs.readFileSync(sapPath, 'utf8');
                const decoded = mhtml
                    .replace(/=\r?\n/g, '')
                    .replace(/=([0-9A-F]{2})/gi, (_, hex) => String.fromCharCode(parseInt(hex, 16)));

                const trMatches = decoded.match(/<tr[\s\S]*?<\/tr>/gi) || [];
                const rows = trMatches.map(tr => {
                    const tdMatches = tr.match(/<t[dh][\s\S]*?<\/t[dh]>/gi) || [];
                    return tdMatches.map(td => td.replace(/<[^>]+>/g, '').trim());
                });

                if (rows.length < 2) return { customerMap, monthMap };

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
                    customer: findIdx(['CUSTOMER NAME', 'CUSTOMER', '고객']),
                    serial: findIdx(['SERIAL NO', 'S/O SERIAL', '시리얼', 'SERIAL']),
                    order: findIdx(['ORDER', '오더']),
                    salesDoc: findIdx(['S/O ORDER', 'SALES DOC', 'SALESDOC', '판매문서'])
                };

                const startIndex = rows.indexOf(headerRow) + 1;
                for (let i = startIndex; i < rows.length; i++) {
                    const cells = rows[i];
                    if (cells.length < 5) continue;

                    const cust = map.customer !== -1 ? (cells[map.customer] || '').trim() : '';
                    const mon = map.mon !== -1 ? (cells[map.mon] || '').trim() : '';
                    const serial = map.serial !== -1 ? (cells[map.serial] || '').trim() : '';
                    const order = map.order !== -1 ? (cells[map.order] || '').trim() : '';
                    const sDoc = map.salesDoc !== -1 ? (cells[map.salesDoc] || '').trim() : '';

                    const keys = [serial, order, sDoc].filter(Boolean);
                    keys.forEach(k => {
                        if (cust && !customerMap.has(k)) customerMap.set(k, cust);
                        if (mon && !monthMap.has(k)) monthMap.set(k, mon);
                    });
                }
            } catch (e) {
                console.warn('[SAP Meta Map] Warning:', e.message);
            }
            return { customerMap, monthMap };
        }

        const sapMeta = loadSapMetaMap();

        // Group unique machines from MES & Attach physical Bay Location (GI_LOC_ID)
        const machineMap = new Map();
        mesItems.forEach(i => {
            const shift = classifyMcShift(i);
            if (!shift) return; // Skip TC/non-MC

            const serial = i.PROD_MDL_ID ? `${i.PROD_MDL_ID}-${i.PROD_MDL_CNT}` : (i.PROD_MDL_CNT || '');
            const ordKey = (i.PROD_ORD_ID || '').trim();
            const serialKey = (i.PROD_MDL_CNT || '').trim();
            const fullSerial = (i.FULL_PROD_MDL_CNT || serial || '').trim();
            const key = ordKey || serial;

            const locInfo = (ordKey && locBayMap.get(ordKey)) || (serialKey && locBayMap.get(serialKey)) || {};
            const workerFound = (ordKey && workerMap.get(ordKey)) || (serialKey && workerMap.get(serialKey)) || i.WORKER_NAME || i.WORKER || i.USER_NAME || i.CHARGER || '';
            const customerFound = i.CUST_NAME || i.CUSTOMER_NAME || i.CUSTOMER || i.BP_NAME || sapMeta.customerMap.get(ordKey) || sapMeta.customerMap.get(serialKey) || sapMeta.customerMap.get(fullSerial) || '';

            let rawMon = i.PROD_YYYYMM || sapMeta.monthMap.get(ordKey) || sapMeta.monthMap.get(serialKey) || sapMeta.monthMap.get(fullSerial) || '';
            let planMonth = '';
            if (rawMon) {
                const digits = rawMon.toString().replace(/\D/g, '');
                if (digits.length >= 6) {
                    planMonth = `${digits.slice(2, 4)}.${digits.slice(4, 6)}월분`;
                } else if (digits.length === 4) {
                    planMonth = `${digits.slice(0, 2)}.${digits.slice(2, 4)}월분`;
                } else {
                    planMonth = rawMon;
                }
            }

            const routingList = orderRoutingMap.get(ordKey) || orderRoutingMap.get(serialKey) || orderRoutingMap.get(fullSerial) || [];

            if (!machineMap.has(key) || (i.CUR_PROC_ID && !machineMap.get(key).CUR_PROC_ID)) {
                machineMap.set(key, {
                    ...i,
                    shift,
                    serial,
                    model: i.PROD_MDL_NAME || i.MTRL_ID,
                    salesDoc: i.PROD_ORD_ID,
                    customer: customerFound,
                    planMonth: planMonth,
                    loc: locInfo.loc || '',
                    area: locInfo.area || '',
                    worker: workerFound,
                    routingSteps: routingList
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

        const machine653 = mesMachines.find(m => (m.serial && m.serial.includes('653')) || (m.salesDoc && m.salesDoc.includes('4157634')));
        console.log('[DEBUG 653] in mesMachines:', machine653 ? { serial: machine653.serial, salesDoc: machine653.salesDoc, loc: machine653.loc, shift: machine653.shift, worker: machine653.worker, curProc: machine653.CUR_PROC_ID } : 'NOT FOUND IN mesMachines');

        // Smart Physical Bay Allocation
        let autoAssignedCount = 0;
        let updatedCount = 0;
        let shippedCount = 0;

        bays.forEach(targetBay => {
            const targetBayCode = extractBayCode(targetBay.bay);
            const targetShift = targetBay.shift;
            const targetArea = targetBay.area;
            
            // Look for an MES machine assigned to this exact physical bay location and shift (Daily Active machines first!)
            const candidates = mesMachines.filter(m => {
                if (!m.loc) return false;
                const mCode = extractBayCode(m.loc);
                if (mCode !== targetBayCode) return false;
                if (m.shift && m.shift !== targetShift) return false;
                return true;
            });

            if (targetBayCode === 'C17') {
                console.log('[DEBUG C17] candidates count:', candidates.length);
                candidates.forEach((c, idx) => {
                    console.log(`[DEBUG C17] cand #${idx}:`, {
                        serial: c.serial,
                        salesDoc: c.salesDoc,
                        model: c.model,
                        loc: c.loc,
                        proc: c.CUR_PROC_ID,
                        worker: c.worker
                    });
                });
            }

            // Check if machine is actually started / in active operation (Yellow/Purple status)
            function hasRealActiveProcess(m) {
                if (!m) return false;
                const aOrd = (m.salesDoc || '').trim();
                const aSerial = (m.PROD_MDL_CNT || m.serial || '').trim();

                // If from Daily work result or Workorder with started operation
                if (locBayMap.get(aOrd)?.isDaily || locBayMap.get(aSerial)?.isDaily) return true;

                // Check actual start date / time presence (Yellow / Purple)
                if (m.PROC_START_DTTM || m.WORK_START_DATE || m.ACT_START_DTTM) return true;
                if (m.LOT_STATUS_CODE === 'START' || m.LOT_STATUS_CODE === 'RUN' || m.LOT_STATUS_CODE === 'CONT') return true;
                if (m.PROC_STATUS_CODE === 'START' || m.PROC_STATUS_CODE === 'RUN' || m.PROC_STATUS_CODE === 'CONT') return true;

                const raw = locBayMap.get(aOrd)?.raw || locBayMap.get(aSerial)?.raw;
                if (raw && (raw.PROC_START_DTTM || raw.LOT_STATUS_CODE === 'START' || raw.WORK_START_DATE)) return true;

                return false;
            }

            function isMainProc(m) {
                if (!m) return false;
                const proc = (m.CUR_PROC_ID || m.PROC_ID || '').toUpperCase();
                // Sub processes starting with P (PGEB, PACONF, PCLM, PTBL, etc.)
                if (proc.startsWith('P') && (proc.includes('GEB') || proc.includes('CONF') || proc.includes('CLM') || proc.includes('TBL') || proc.includes('ACT') || proc.includes('10'))) {
                    return false;
                }
                return true; // GBDLEVEL, GCLM, GEB, GAJ, GIN, GSR, GOT, XYZHOME, ILUPIP, SCALE, etc.
            }

            // 4-Way Smart Priority Sort:
            // 0. Main Body Assembly (GBDLEVEL, GAJ, GIN 등) vs Sub/Pallet Assembly (PGEB, PACONF 등)
            // 1. QMS GIN Inspection Request Bay (Highest Ground Truth)
            // 2. Daily Live Work Result (CN0WBFAG/H)
            // 3. Real Active Started Status (Yellow/Purple)
            // 4. Newest Production Order (e.g. 700004161378 / 52호기 > 700004154500 / 1109호기)
            // 5. Newest Production Month (e.g. 202609 > 202608)
            candidates.sort((a, b) => {
                const aMain = isMainProc(a) ? 1 : 0;
                const bMain = isMainProc(b) ? 1 : 0;
                if (aMain !== bMain) return bMain - aMain;

                const aOrd = (a.salesDoc || a.PROD_ORD_ID || '').trim();
                const bOrd = (b.salesDoc || b.PROD_ORD_ID || '').trim();
                const aSerial = (a.PROD_MDL_CNT || a.serial || '').trim();
                const bSerial = (b.PROD_MDL_CNT || b.serial || '').trim();

                // 1. QMS Inspection Request Match (Highest Truth)
                const aQms = (locBayMap.get(aOrd)?.isQms || locBayMap.get(aSerial)?.isQms) ? 1 : 0;
                const bQms = (locBayMap.get(bOrd)?.isQms || locBayMap.get(bSerial)?.isQms) ? 1 : 0;
                if (aQms !== bQms) return bQms - aQms;

                // 2. Daily Active
                const aDaily = (locBayMap.get(aOrd)?.isDaily || locBayMap.get(aSerial)?.isDaily) ? 1 : 0;
                const bDaily = (locBayMap.get(bOrd)?.isDaily || locBayMap.get(bSerial)?.isDaily) ? 1 : 0;
                if (aDaily !== bDaily) return bDaily - aDaily;

                // 3. Real Active Started Status
                const aActive = hasRealActiveProcess(a) ? 1 : 0;
                const bActive = hasRealActiveProcess(b) ? 1 : 0;
                if (aActive !== bActive) return bActive - aActive;

                // 4. Newest Production Order ID
                const aOrdNum = parseInt(aOrd.replace(/\D/g, '')) || 0;
                const bOrdNum = parseInt(bOrd.replace(/\D/g, '')) || 0;
                if (aOrdNum !== bOrdNum) return bOrdNum - aOrdNum;

                // 5. Newest Production Month
                const aMon = parseInt((a.PROD_YYYYMM || '').replace(/\D/g, '')) || 0;
                const bMon = parseInt((b.PROD_YYYYMM || '').replace(/\D/g, '')) || 0;
                if (aMon !== bMon) return bMon - aMon;

                return 0;
            });

            const locMatch = candidates[0];

            if (locMatch) {
                targetBay.assigned = true;
                targetBay.model = locMatch.model;
                targetBay.serial = locMatch.serial;
                targetBay.salesDoc = locMatch.salesDoc || '';
                targetBay.customer = locMatch.customer || targetBay.customer || '';
                targetBay.planMonth = locMatch.planMonth || targetBay.planMonth || '';
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
                targetBay.routingSteps = locMatch.routingSteps || orderRoutingMap.get(ordKey) || orderRoutingMap.get(serialKey) || [];
                
                // Prioritize real-time active (START / RUN / CONT) process from actual routing
                const activeFromRouting = (targetBay.routingSteps || []).find(s => s.status === 'START' || s.status === 'RUN' || s.status === 'CONT');
                targetBay.currentProcess = (activeFromRouting && activeFromRouting.code) || locMatch.CUR_PROC_ID || locMatch.PROC_ID || 'BASE';

                if (locMatch.PROD_ORD_STATUS_NAME) {
                    targetBay.issue = `[상태: ${locMatch.PROD_ORD_STATUS_NAME}]`;
                }

                // Attach concurrent / sub machines at this bay
                targetBay.subMachines = candidates.slice(1).map(c => {
                    const cOrd = (c.salesDoc || '').trim();
                    const cSerial = (c.PROD_MDL_CNT || c.serial || '').trim();
                    const cWorker = c.worker || (cOrd && workerMap.get(cOrd)) || (cSerial && workerMap.get(cSerial)) || c.WORKER_NAME || '';
                    const cRouting = c.routingSteps || orderRoutingMap.get(cOrd) || orderRoutingMap.get(cSerial) || [];
                    const cActive = cRouting.find(s => s.status === 'START' || s.status === 'RUN' || s.status === 'CONT');
                    return {
                        model: c.model || '',
                        serial: c.serial || c.PROD_MDL_CNT || '',
                        salesDoc: c.salesDoc || '',
                        currentProcess: (cActive && cActive.code) || c.CUR_PROC_ID || c.PROC_ID || '',
                        worker: cWorker,
                        customer: c.customer || '',
                        planMonth: c.planMonth || '',
                        spec: c.MTRL_ID || '',
                        issue: c.PROD_ORD_STATUS_NAME ? `[상태: ${c.PROD_ORD_STATUS_NAME}]` : '',
                        routingSteps: cRouting
                    };
                });

                autoAssignedCount++;
                return;
            } else {
                targetBay.subMachines = [];
            }

            // 2. If bay was already assigned manually, sync its in-progress stage
            if (targetBay.assigned) {
                const baySerial = (targetBay.serial || '').trim();
                const bayOrder = (targetBay.salesDoc || '').trim();

                const match = mesMachines.find(m => {
                    const mSerial = (m.serial || '').trim();
                    const mCnt = (m.PROD_MDL_CNT || '').trim();
                    const mOrd = (m.salesDoc || m.PROD_ORD_ID || '').trim();

                    if (baySerial && (mSerial === baySerial || mCnt === baySerial || baySerial.includes(mCnt) || mSerial.includes(baySerial))) return true;
                    if (bayOrder && (mOrd === bayOrder || mOrd.includes(bayOrder) || bayOrder.includes(mOrd))) return true;
                    return false;
                });

                if (match) {
                    const ordKey = (match.salesDoc || '').trim();
                    const serialKey = (match.PROD_MDL_CNT || match.serial || '').trim();
                    targetBay.routingSteps = match.routingSteps || orderRoutingMap.get(ordKey) || orderRoutingMap.get(serialKey) || targetBay.routingSteps || [];
                    
                    const matchActive = (targetBay.routingSteps || []).find(s => s.status === 'START' || s.status === 'RUN' || s.status === 'CONT');
                    targetBay.currentProcess = (matchActive && matchActive.code) || match.CUR_PROC_ID || match.PROC_ID || targetBay.currentProcess || 'BASE';
                    
                    targetBay.startDate = match.START_PLAN_DATE || targetBay.startDate;
                    targetBay.deliveryDate = match.SHIP_TARGET_DATE || targetBay.deliveryDate;
                    targetBay.spec = match.MTRL_ID || targetBay.spec;
                    targetBay.customer = match.customer || targetBay.customer || '';
                    targetBay.planMonth = match.planMonth || targetBay.planMonth || '';
                    const workerFound = match.worker || (ordKey && workerMap.get(ordKey)) || (serialKey && workerMap.get(serialKey)) || match.WORKER_NAME || match.WORKER || match.USER_NAME || match.CHARGER || '';
                    if (workerFound) targetBay.worker = workerFound;

                    targetBay.source = 'MES';
                    targetBay.isShipped = false;
                    if (match.PROD_ORD_STATUS_NAME) {
                        targetBay.issue = `[상태: ${match.PROD_ORD_STATUS_NAME}]`;
                    } else if (targetBay.issue && targetBay.issue.includes('출하 완료')) {
                        targetBay.issue = '';
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

const HEARTBEAT_TIMEOUT = 1800000; // 30 minutes
const GRACE_PERIOD = 1800000; // 30 minutes grace period
const startupTime = Date.now();

setInterval(() => {
    const now = Date.now();
    if (hasReceivedHeartbeat) {
        if (now - lastHeartbeat > HEARTBEAT_TIMEOUT) {
            console.log('[AUTO-SHUTDOWN] No client connected for 30 minutes. Exiting...');
            process.exit(0);
        }
    } else {
        if (now - startupTime > GRACE_PERIOD) {
            console.log('[AUTO-SHUTDOWN] Grace period expired without client connection. Exiting...');
            process.exit(0);
        }
    }
}, 10000);
