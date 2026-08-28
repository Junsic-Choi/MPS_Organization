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

// Default Bay Template if no data exists
function getDefaultBays() {
    const bays = [];
    
    // MC 1직 (E구역 1~10, F구역 1~6)
    for (let i = 1; i <= 10; i++) bays.push({ id: `mc1-e${i}`, shift: 'MC1직', bay: `E${i}`, assigned: false, model: '', serial: '', salesDoc: '', customer: '', worker: '', currentProcess: 'BASE', spec: '', issue: '', startDate: '', deliveryDate: '' });
    for (let i = 1; i <= 6; i++) bays.push({ id: `mc1-f${i}`, shift: 'MC1직', bay: `F${i}`, assigned: false, model: '', serial: '', salesDoc: '', customer: '', worker: '', currentProcess: 'BASE', spec: '', issue: '', startDate: '', deliveryDate: '' });

    // MC 2직 (C구역 1~12, D구역 1~12)
    for (let i = 1; i <= 12; i++) bays.push({ id: `mc2-c${i}`, shift: 'MC2직', bay: `C${i}`, assigned: false, model: '', serial: '', salesDoc: '', customer: '', worker: '', currentProcess: 'BASE', spec: '', issue: '', startDate: '', deliveryDate: '' });
    for (let i = 1; i <= 12; i++) bays.push({ id: `mc2-d${i}`, shift: 'MC2직', bay: `D${i}`, assigned: false, model: '', serial: '', salesDoc: '', customer: '', worker: '', currentProcess: 'BASE', spec: '', issue: '', startDate: '', deliveryDate: '' });

    // MC 3직 (B구역 1~10)
    for (let i = 1; i <= 10; i++) bays.push({ id: `mc3-b${i}`, shift: 'MC3직', bay: `B${i}`, assigned: false, model: '', serial: '', salesDoc: '', customer: '', worker: '', currentProcess: 'BASE', spec: '', issue: '', startDate: '', deliveryDate: '' });

    // MC 4직 (A구역 1~10)
    for (let i = 1; i <= 10; i++) bays.push({ id: `mc4-a${i}`, shift: 'MC4직', bay: `A${i}`, assigned: false, model: '', serial: '', salesDoc: '', customer: '', worker: '', currentProcess: 'BASE', spec: '', issue: '', startDate: '', deliveryDate: '' });

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
        const authHeader = req.headers['authorization'] || req.body.token || '';
        const yyyymm = req.body.yyyymm || new Date().toISOString().slice(0, 7).replace('-', '');

        const payload = {
            BizActId: "BR_DNS_MES_SEL_ProdProgressStatus_Assy",
            InDataList: {
                IN_DATA: [
                    {
                        ENTERPRISE_ID: "1800",
                        PLANT_ID: "1840",
                        DEPT_ID: null,
                        PROD_YYYYMM_FROM: yyyymm
                    }
                ]
            },
            OutData: "OUT_DATA",
            port: "8082"
        };

        const headers = {
            'Content-Type': 'application/json'
        };
        if (authHeader) {
            headers['Authorization'] = authHeader.startsWith('Bearer ') ? authHeader : `Bearer ${authHeader}`;
        }

        let mesItems = [];
        try {
            const mesRes = await fetch('http://mes.dn-solutions.com:8081/api/json/query', {
                method: 'POST',
                headers,
                body: JSON.stringify(payload)
            });

            if (mesRes.ok) {
                const mesJson = await mesRes.json();
                mesItems = mesJson.OUT_DATA || (mesJson.InDataList && mesJson.InDataList.OUT_DATA) || [];
            } else {
                const errText = await mesRes.text().catch(() => '');
                console.warn(`[MES] Server response status: ${mesRes.status} - ${errText}`);
                if (mesRes.status === 401) {
                    return res.json({
                        success: false,
                        error: 'MES 인증 토큰(Bearer Token)이 만료되었거나 유효하지 않습니다. 최신 토큰을 설정해주세요.',
                        status: 401
                    });
                }
            }
        } catch (fetchErr) {
            console.warn('[MES] Remote connection warning:', fetchErr.message);
        }

        // If client provided manual raw OUT_DATA array directly
        if (Array.isArray(req.body.outData) && req.body.outData.length > 0) {
            mesItems = req.body.outData;
        }

        if (mesItems.length === 0) {
            return res.json({ 
                success: false, 
                error: 'MES 데이터를 수신하지 못했습니다. (사내망 연결 또는 로그인 토큰을 확인해주세요)',
                count: 0 
            });
        }

        // Load existing bays
        let bays = [];
        if (fs.existsSync(DATA_FILE)) {
            const data = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
            bays = data.bays || [];
        } else {
            bays = getDefaultBays();
        }

        let matchedCount = 0;

        // Parse Work Center (WC_ID) to bay matching
        const normalizeBayKey = (wc) => {
            if (!wc) return null;
            const clean = wc.toUpperCase().trim();
            
            // Match Patterns like E1~E10, F1~F6, C1~C12, D1~D12, B1~B10, A1~A10
            const m = clean.match(/([A-F])(0?[1-9]|1[0-2])/);
            if (m) {
                const zone = m[1];
                const num = parseInt(m[2], 10);
                return `${zone}${num}`;
            }
            return null;
        };

        mesItems.forEach(item => {
            const bayCode = normalizeBayKey(item.WC_ID);
            if (!bayCode) return;

            // Find matching bay in bays list
            const targetBay = bays.find(b => b.bay.toUpperCase() === bayCode);
            if (targetBay) {
                targetBay.assigned = true;
                targetBay.model = item.PROD_MDL_NAME || item.MTRL_ID || targetBay.model;
                
                const serialNo = (item.PROD_MDL_ID && item.PROD_MDL_CNT) 
                    ? `${item.PROD_MDL_ID}-${item.PROD_MDL_CNT}` 
                    : (item.PROD_MDL_CNT || targetBay.serial);
                targetBay.serial = serialNo || targetBay.serial;

                targetBay.currentProcess = item.CUR_PROC_ID || item.PROC_ID || targetBay.currentProcess || 'BASE';
                targetBay.salesDoc = item.PROD_ORD_ID || targetBay.salesDoc;
                targetBay.startDate = item.START_PLAN_DATE || targetBay.startDate;
                targetBay.deliveryDate = item.SHIP_TARGET_DATE || targetBay.deliveryDate;
                targetBay.spec = item.MTRL_ID || targetBay.spec;
                
                if (item.LOT_STATUS_CODE && item.LOT_STATUS_CODE !== 'NONE') {
                    targetBay.issue = `[${item.PROD_ORD_STATUS_NAME || item.LOT_STATUS_CODE}]`;
                }

                targetBay.source = 'MES';
                matchedCount++;
            }
        });

        // Save updated bays
        fs.writeFileSync(DATA_FILE, JSON.stringify({ bays, updatedAt: new Date().toISOString() }, null, 2), 'utf8');
        console.log(`[MES] Synced ${matchedCount} bays from ${mesItems.length} MES records`);

        res.json({
            success: true,
            totalMesRecords: mesItems.length,
            matchedBaysCount: matchedCount,
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
            if (ver.includes('0AM1') || ver.includes('0AMA') || ver.includes('MC1') || ver.includes('MC 1')) {
                shift = 'MC1직';
            } else if (ver.includes('0AM2') || ver.includes('MC2') || ver.includes('MC 2')) {
                shift = 'MC2직';
            } else if (ver.includes('0AM3') || ver.includes('MC3') || ver.includes('MC 3')) {
                shift = 'MC3직';
            } else if (ver.includes('0AM4') || ver.includes('MC4') || ver.includes('MC 4')) {
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
