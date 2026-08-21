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
