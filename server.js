const express = require('express');
const path = require('path');
const cors = require('cors');
const { exec, spawn } = require('child_process');
const fs = require('fs');
const multer = require('multer');
const XLSX = require('xlsx');
const { processMpsFile } = require('./extractor');


const upload = multer({ 
    storage: multer.memoryStorage(),
    limits: { fileSize: 50 * 1024 * 1024 } // 50MB 제한
});

const app = express();
console.log('--- Server Initializing ---');
const PORT = 8890;

app.use(cors());
app.use(express.json());
app.use(express.static(__dirname));

// 루트 경로(/) 접속 시 대시보드로 리다이렉트
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'dashboard.html'));
});

// 가용 파일 목록 조회 API
app.get('/api/list-files', (req, res) => {
    try {
        const uploadDir = process.pkg ? path.dirname(process.execPath) : __dirname;
        const files = fs.readdirSync(uploadDir)
            .filter(f => f.startsWith('MPS') && f.endsWith('.xlsx'))
            .sort().reverse(); // 최신순 (이름 기준)
        const filesWithDetails = files.map(filename => {
            const filePath = path.join(uploadDir, filename);
            const stats = fs.statSync(filePath);
            return {
                filename,
                size: stats.size,
                mtime: stats.mtime
            };
        });
        res.json({ success: true, files: filesWithDetails });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

// [Live] 서버 사이드 실시간 추출 API (브라우저 메모리 부족 해결용)
app.post('/api/extract-live', upload.single('file'), async (req, res) => {
    try {
        console.log(`[api] Received extract request: ${req.file ? req.file.originalname : 'No file'}`);
        
        if (!req.file) {
            return res.status(400).json({ success: false, error: '파일이 업로드되지 않았습니다. (Multipart field name: file)' });
        }

        let rules = {};
        if (req.body.rules) {
            try {
                rules = JSON.parse(req.body.rules);
            } catch (e) {
                console.error('[api] Failed to parse rules:', e.message);
            }
        }
        
        console.log(`[api] Live extract started: ${req.file.originalname} (${req.file.size} bytes)`);
        
        // 업로드된 파일을 서버에 저장 (exe 실행 시 exe 파일과 같은 폴더에 저장되도록 유도)
        const uploadDir = process.pkg ? path.dirname(process.execPath) : __dirname;
        const savePath = path.join(uploadDir, req.file.originalname);
        fs.writeFileSync(savePath, req.file.buffer);
        console.log(`[api] File saved to server: ${savePath}`);
        
        const result = await processMpsFile(req.file.buffer, rules);
        
        console.log(`[api] Live extract success: ${result.finalResults.length} rows`);
        res.json({ success: true, ...result });
    } catch (err) {
        console.error(`[api] Live extract failed:`, err);
        res.status(500).json({ 
            success: false, 
            error: err.message,
            stack: process.env.NODE_ENV === 'development' ? err.stack : undefined 
        });
    }
});

// [Saved] 서버에 저장된 파일 직접 추출 API
app.post('/api/extract-saved', async (req, res) => {
    try {
        const { filename, rules: rulesStr } = req.body;
        if (!filename) {
            return res.status(400).json({ success: false, error: '파일명이 제공되지 않았습니다.' });
        }
        
        const uploadDir = process.pkg ? path.dirname(process.execPath) : __dirname;
        const filePath = path.join(uploadDir, filename);
        if (!fs.existsSync(filePath)) {
            return res.status(404).json({ success: false, error: `파일을 찾을 수 없습니다: ${filename}` });
        }
        
        let rules = {};
        if (rulesStr) {
            try {
                rules = typeof rulesStr === 'object' ? rulesStr : JSON.parse(rulesStr);
            } catch (e) {
                console.error('[api] Failed to parse rules:', e.message);
            }
        }
        
        console.log(`[api] Saved file extract started: ${filename}`);
        const fileBuffer = fs.readFileSync(filePath);
        const result = await processMpsFile(fileBuffer, rules);
        
        console.log(`[api] Saved file extract success: ${result.finalResults.length} rows`);
        res.json({ success: true, ...result });
    } catch (err) {
        console.error(`[api] Saved file extract failed:`, err);
        res.status(500).json({ success: false, error: err.message });
    }
});

// 개인화 설정 로드 API
app.get('/api/preferences', (req, res) => {
    try {
        const uploadDir = process.pkg ? path.dirname(process.execPath) : __dirname;
        const prefPath = path.join(uploadDir, 'preferences.json');
        if (fs.existsSync(prefPath)) {
            const data = fs.readFileSync(prefPath, 'utf8');
            res.json({ success: true, preferences: JSON.parse(data) });
        } else {
            res.json({ success: true, preferences: null });
        }
    } catch (err) {
        console.error('[api] Load preferences failed:', err.message);
        res.status(500).json({ success: false, error: err.message });
    }
});

// 개인화 설정 저장 API
app.post('/api/preferences', (req, res) => {
    try {
        const uploadDir = process.pkg ? path.dirname(process.execPath) : __dirname;
        const prefPath = path.join(uploadDir, 'preferences.json');
        fs.writeFileSync(prefPath, JSON.stringify(req.body, null, 2), 'utf8');
        res.json({ success: true });
    } catch (err) {
        console.error('[api] Save preferences failed:', err.message);
        res.status(500).json({ success: false, error: err.message });
    }
});

// SAP 파일 업로드 API (덮어쓰기)
app.post('/api/upload-sap', upload.single('file'), (req, res) => {
    try {
        const { type } = req.body;
        if (!type || !['1842', '1840', 'component_1842', 'component_1840'].includes(type)) {
            return res.status(400).json({ success: false, error: '올바른 타입(1842, 1840, component_1842 또는 component_1840)을 지정해주세요.' });
        }
        if (!req.file) {
            return res.status(400).json({ success: false, error: '업로드된 파일이 없습니다.' });
        }
        const uploadDir = process.pkg ? path.dirname(process.execPath) : __dirname;
        const savePath = path.join(uploadDir, `sap_${type}.mhtml`);
        fs.writeFileSync(savePath, req.file.buffer);
        console.log(`[sap-upload] Saved sap_${type}.mhtml to server`);
        res.json({ success: true });
    } catch (err) {
        console.error('[sap-upload] Failed to save SAP file:', err);
        res.status(500).json({ success: false, error: err.message });
    }
});

// SAP 파일 삭제 API
app.post('/api/clear-sap', (req, res) => {
    try {
        const { type } = req.body;
        if (!type || !['1842', '1840', 'component_1842', 'component_1840'].includes(type)) {
            return res.status(400).json({ success: false, error: '올바른 타입(1842, 1840, component_1842 또는 component_1840)을 지정해주세요.' });
        }
        const uploadDir = process.pkg ? path.dirname(process.execPath) : __dirname;
        const savePath = path.join(uploadDir, `sap_${type}.mhtml`);
        if (fs.existsSync(savePath)) {
            fs.unlinkSync(savePath);
            console.log(`[sap-clear] Deleted sap_${type}.mhtml`);
        }
        res.json({ success: true });
    } catch (err) {
        console.error('[sap-clear] Failed to clear SAP file:', err);
        res.status(500).json({ success: false, error: err.message });
    }
});

// SAP 파일 로드 API
app.get('/api/load-sap/:type', (req, res) => {
    try {
        const { type } = req.params;
        if (!['1842', '1840', 'component_1842', 'component_1840'].includes(type)) {
            return res.status(400).json({ success: false, error: '올바른 타입(1842, 1840, component_1842 또는 component_1840)을 지정해주세요.' });
        }
        const uploadDir = process.pkg ? path.dirname(process.execPath) : __dirname;
        const filePath = path.join(uploadDir, `sap_${type}.mhtml`);
        if (fs.existsSync(filePath)) {
            const content = fs.readFileSync(filePath, 'utf8');
            const stats = fs.statSync(filePath);
            res.json({ success: true, exists: true, content, mtime: stats.mtime });
        } else {
            res.json({ success: true, exists: false });
        }
    } catch (err) {
        console.error('[sap-load] Failed to load SAP file:', err);
        res.status(500).json({ success: false, error: err.message });
    }
});

// Heartbeat state
let lastHeartbeat = Date.now();
let hasReceivedHeartbeat = false;

app.post('/api/heartbeat', (req, res) => {
    lastHeartbeat = Date.now();
    hasReceivedHeartbeat = true;
    res.sendStatus(200);
});

// SAP ERP 자동 동기화 API
const { runSapSync, getSyncStatus } = require('./sap_automation/sync_engine');

app.post('/api/sync-sap', async (req, res) => {
    try {
        const { startMonth, endMonth } = req.body || {};
        const status = getSyncStatus();
        if (status.running) {
            return res.json({ success: true, message: '이미 SAP 동기화가 진행 중입니다.', status });
        }
        // Run in background so request doesn't timeout
        runSapSync({ startMonth, endMonth }).catch(e => {
            console.error('[SAP Sync Error in background]:', e.message);
        });
        res.json({ success: true, message: 'SAP 동기화가 시작되었습니다.' });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

const sapFileCache = {};

function getSapFileStats(fn, fp) {
    if (!fs.existsSync(fp)) return null;
    const stat = fs.statSync(fp);
    const cached = sapFileCache[fn];
    if (cached && cached.mtimeMs === stat.mtimeMs) {
        return cached;
    }
    let rowCount = 0;
    try {
        const content = fs.readFileSync(fp, 'utf8');
        const trs = content.match(/<tr[^>]*>/gi);
        if (trs && trs.length > 1) rowCount = trs.length - 1;
    } catch (e) {}

    const result = {
        size: (stat.size / 1024 / 1024).toFixed(2) + ' MB',
        mtime: stat.mtime,
        mtimeMs: stat.mtimeMs,
        rows: rowCount
    };
    sapFileCache[fn] = result;
    return result;
}

app.get('/api/sync-sap/status', (req, res) => {
    try {
        const status = getSyncStatus();
        const files = {};
        ['sap_1840.mhtml', 'sap_1842.mhtml', 'sap_component_1840.mhtml', 'sap_component_1842.mhtml'].forEach(fn => {
            const fp = path.join(__dirname, fn);
            files[fn] = getSapFileStats(fn, fp);
        });
        res.json({ success: true, ...status, files });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

app.get('/api/download-sap/:filename', (req, res) => {
    const valid = ['sap_1840.mhtml', 'sap_1842.mhtml', 'sap_component_1840.mhtml', 'sap_component_1842.mhtml'];
    const fn = req.params.filename;
    if (!valid.includes(fn)) return res.status(400).send('Invalid file name');
    const fp = path.join(__dirname, fn);
    if (!fs.existsSync(fp)) return res.status(404).send('File not found');
    res.download(fp, fn);
});

app.get('/api/sap-preview', (req, res) => {
    try {
        const targetFiles = {
            '1840': { name: '1840 남산+ 생산계획', file: 'sap_1840.mhtml' },
            'comp_1840': { name: '1840 남산+ 부품소요량', file: 'sap_component_1840.mhtml' },
            '1842': { name: '1842 성주 생산계획', file: 'sap_1842.mhtml' },
            'comp_1842': { name: '1842 성주 부품소요량', file: 'sap_component_1842.mhtml' }
        };
        const result = {};
        for (const [key, info] of Object.entries(targetFiles)) {
            const fp = path.join(__dirname, info.file);
            if (!fs.existsSync(fp)) {
                result[key] = { exists: false, name: info.name, file: info.file };
                continue;
            }
            const stat = fs.statSync(fp);
            const content = fs.readFileSync(fp, 'utf8');
            const trs = content.match(/<tr[^>]*>[\s\S]*?<\/tr>/gi) || [];
            let headers = [];
            let rows = [];
            if (trs.length > 0) {
                headers = (trs[0].match(/<t[dh][^>]*>[\s\S]*?<\/t[dh]>/gi) || [])
                    .map(c => c.replace(/<[^>]+>/g, '').trim())
                    .filter(c => c.length > 0);
                for (let i = 1; i < Math.min(trs.length, 11); i++) {
                    const cells = (trs[i].match(/<td[^>]*>[\s\S]*?<\/td>/gi) || [])
                        .map(c => c.replace(/<[^>]+>/g, '').trim());
                    if (cells.length > 0) rows.push(cells);
                }
            }
            result[key] = {
                exists: true,
                name: info.name,
                file: info.file,
                size: (stat.size / 1024 / 1024).toFixed(2) + ' MB',
                mtime: stat.mtime,
                totalRows: Math.max(0, trs.length - 1),
                headers: headers.slice(0, 15),
                sampleRows: rows.map(r => r.slice(0, 15))
            };
        }
        res.json({ success: true, data: result });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});
// ==========================================
// SAP Diff & Engineering Change (설변) APIs
// ==========================================
const HISTORY_DIR = path.join(__dirname, 'sap_history');

function parseComponentMhtml(filePath) {
    if (!fs.existsSync(filePath)) return [];
    const content = fs.readFileSync(filePath, 'utf8');
    const trs = content.match(/<tr[^>]*>[\s\S]*?<\/tr>/gi) || [];
    if (trs.length < 2) return [];

    const headers = (trs[0].match(/<t[dh][^>]*>[\s\S]*?<\/t[dh]>/gi) || [])
        .map(c => c.replace(/<[^>]+>/g, '').trim().toUpperCase());

    const idxPlant = headers.findIndex(h => h.includes('PLANT') || h.includes('플랜트'));
    const idxSerial = headers.findIndex(h => h.includes('SERIAL') || h.includes('호기'));
    const idxSO = headers.findIndex(h => h.includes('S/O') || h.includes('SALES'));
    const idxOrder = headers.findIndex(h => h.includes('ORDER') && !h.includes('TYPE') && !h.includes('S/O'));
    const idxMonth = headers.findIndex(h => h.includes('PROD.MONTH') || h.includes('생산월'));
    const idxMat = headers.findIndex(h => h.includes('MATERIAL NUMBER') || h.includes('자재'));
    const idxMatDesc = headers.findIndex(h => h.includes('MATERIAL DESCRIPTION') || h.includes('자재내역'));
    const idxComp = headers.findIndex(h => h === 'COMPONENT' || h.includes('구성부품'));
    const idxCompDesc = headers.findIndex(h => h.includes('COMP.DESC') || h.includes('구성부품내역'));
    const idxQty = headers.findIndex(h => h.includes('REQUIREMENT QUANTITY') || h.includes('소요량'));
    const idxUnit = headers.findIndex(h => h.includes('BASE UNIT') || h.includes('단위'));

    const records = [];
    for (let i = 1; i < trs.length; i++) {
        const cells = (trs[i].match(/<td[^>]*>[\s\S]*?<\/td>/gi) || [])
            .map(c => c.replace(/<[^>]+>/g, '').trim());
        if (cells.length === 0) continue;

        const comp = cells[idxComp] || '';
        if (!comp) continue;

        records.push({
            plant: cells[idxPlant] || '',
            serial: cells[idxSerial] || '',
            soOrder: cells[idxSO] || '',
            order: cells[idxOrder] || '',
            prodMonth: cells[idxMonth] || '',
            material: cells[idxMat] || '',
            matDesc: cells[idxMatDesc] || '',
            component: comp,
            compDesc: cells[idxCompDesc] || '',
            qty: parseFloat((cells[idxQty] || '0').replace(/,/g, '')) || 0,
            unit: cells[idxUnit] || ''
        });
    }
    return records;
}

function getSnapshotDir(snapshotId) {
    if (!snapshotId || snapshotId === 'current') return __dirname;
    return path.join(HISTORY_DIR, snapshotId);
}

function loadComponentDataset(snapshotId) {
    const dir = getSnapshotDir(snapshotId);
    const f1840 = path.join(dir, 'sap_component_1840.mhtml');
    const f1842 = path.join(dir, 'sap_component_1842.mhtml');

    const records = [];
    if (fs.existsSync(f1840)) records.push(...parseComponentMhtml(f1840));
    if (fs.existsSync(f1842)) records.push(...parseComponentMhtml(f1842));
    return records;
}

function getSnapshotList() {
    const list = [
        {
            id: 'current',
            title: '현재 수집 데이터 (최신 Live)',
            timestamp: Date.now(),
            isCurrent: true
        }
    ];
    if (fs.existsSync(HISTORY_DIR)) {
        const dirs = fs.readdirSync(HISTORY_DIR).filter(d => {
            const full = path.join(HISTORY_DIR, d);
            return fs.statSync(full).isDirectory();
        });
        dirs.forEach(d => {
            const metaPath = path.join(HISTORY_DIR, d, 'metadata.json');
            if (fs.existsSync(metaPath)) {
                try {
                    const meta = JSON.parse(fs.readFileSync(metaPath, 'utf8'));
                    list.push({
                        id: meta.id || d,
                        title: meta.title || d,
                        timestamp: meta.timestamp || 0,
                        createdAt: meta.createdAt || ''
                    });
                } catch (e) {
                    list.push({ id: d, title: d, timestamp: 0 });
                }
            } else {
                list.push({ id: d, title: d, timestamp: 0 });
            }
        });
    }
    return list.sort((a, b) => {
        if (a.id === 'current') return -1;
        if (b.id === 'current') return 1;
        return (b.timestamp || 0) - (a.timestamp || 0);
    });
}

function calculateComponentDiff(recordsA, recordsB) {
    const aggregate = (records) => {
        const map = new Map();
        for (const r of records) {
            const key = [r.plant, r.soOrder || r.order || '', r.serial || '', r.component].join('|');
            if (!map.has(key)) {
                map.set(key, { ...r, qty: 0 });
            }
            map.get(key).qty += r.qty;
        }
        return map;
    };

    const mapA = aggregate(recordsA);
    const mapB = aggregate(recordsB);

    const changes = [];
    const ordersAffected = new Set();

    // Check items in B
    for (const [key, itemB] of mapB) {
        if (!mapA.has(key)) {
            changes.push({
                changeType: 'NEW',
                plant: itemB.plant,
                soOrder: itemB.soOrder,
                serial: itemB.serial,
                order: itemB.order,
                prodMonth: itemB.prodMonth,
                material: itemB.material,
                matDesc: itemB.matDesc,
                component: itemB.component,
                compDesc: itemB.compDesc,
                unit: itemB.unit,
                oldQty: 0,
                newQty: itemB.qty,
                diffQty: itemB.qty
            });
            ordersAffected.add(itemB.soOrder || itemB.order || itemB.serial);
        } else {
            const itemA = mapA.get(key);
            if (Math.abs(itemA.qty - itemB.qty) > 0.0001) {
                const diff = itemB.qty - itemA.qty;
                changes.push({
                    changeType: 'QTY_CHANGE',
                    plant: itemB.plant,
                    soOrder: itemB.soOrder,
                    serial: itemB.serial,
                    order: itemB.order,
                    prodMonth: itemB.prodMonth,
                    material: itemB.material,
                    matDesc: itemB.matDesc,
                    component: itemB.component,
                    compDesc: itemB.compDesc,
                    unit: itemB.unit,
                    oldQty: itemA.qty,
                    newQty: itemB.qty,
                    diffQty: diff
                });
                ordersAffected.add(itemB.soOrder || itemB.order || itemB.serial);
            }
        }
    }

    // Check items deleted from A
    for (const [key, itemA] of mapA) {
        if (!mapB.has(key)) {
            changes.push({
                changeType: 'DELETED',
                plant: itemA.plant,
                soOrder: itemA.soOrder,
                serial: itemA.serial,
                order: itemA.order,
                prodMonth: itemA.prodMonth,
                material: itemA.material,
                matDesc: itemA.matDesc,
                component: itemA.component,
                compDesc: itemA.compDesc,
                unit: itemA.unit,
                oldQty: itemA.qty,
                newQty: 0,
                diffQty: -itemA.qty
            });
            ordersAffected.add(itemA.soOrder || itemA.order || itemA.serial);
        }
    }

    const summary = {
        totalNew: changes.filter(c => c.changeType === 'NEW').length,
        totalDeleted: changes.filter(c => c.changeType === 'DELETED').length,
        totalQtyChanged: changes.filter(c => c.changeType === 'QTY_CHANGE').length,
        totalOrdersAffected: ordersAffected.size,
        totalChanges: changes.length
    };

    return { summary, changes };
}

app.get('/api/sap-diff/history', (req, res) => {
    try {
        const list = getSnapshotList();
        res.json({ success: true, snapshots: list });
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

app.post('/api/sap-diff/compare', (req, res) => {
    try {
        const snapshots = getSnapshotList();
        let targetId = req.body.targetId || 'current';
        let baseId = req.body.baseId;

        if (!baseId) {
            const prev = snapshots.find(s => s.id !== targetId);
            baseId = prev ? prev.id : targetId;
        }

        const baseTitle = (snapshots.find(s => s.id === baseId) || {}).title || baseId;
        const targetTitle = (snapshots.find(s => s.id === targetId) || {}).title || targetId;

        const baseRecords = loadComponentDataset(baseId);
        const targetRecords = loadComponentDataset(targetId);

        const diffResult = calculateComponentDiff(baseRecords, targetRecords);

        res.json({
            success: true,
            baseId,
            baseTitle,
            targetId,
            targetTitle,
            summary: diffResult.summary,
            changes: diffResult.changes
        });
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

app.post('/api/sap-diff/export-excel', (req, res) => {
    try {
        const { changes } = req.body;
        if (!Array.isArray(changes)) {
            return res.status(400).send('Invalid changes data');
        }

        const rows = changes.map(c => ({
            '변동유형': c.changeType === 'NEW' ? '신규 투입' : (c.changeType === 'DELETED' ? '삭제/제외' : '소요량 변동'),
            '플랜트': c.plant === '1840' ? '1840 (남산+)' : (c.plant === '1842' ? '1842 (성주)' : c.plant),
            'S/O번호': c.soOrder,
            '시리얼': c.serial,
            '계획오더': c.order,
            '생산월': c.prodMonth,
            '모품번(기종)': c.material,
            '기종내역': c.matDesc,
            '가공품번(Component)': c.component,
            '가공품명(Comp.Desc)': c.compDesc,
            '단위': c.unit,
            '이전 소요량': c.oldQty,
            '최신 소요량': c.newQty,
            '소요량 차이(+/-)': c.diffQty
        }));

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(rows);

        const colWidths = [
            { wch: 12 }, { wch: 14 }, { wch: 12 }, { wch: 15 },
            { wch: 14 }, { wch: 10 }, { wch: 22 }, { wch: 30 },
            { wch: 18 }, { wch: 28 }, { wch: 8 }, { wch: 12 },
            { wch: 12 }, { wch: 14 }
        ];
        ws['!cols'] = colWidths;

        XLSX.utils.book_append_sheet(wb, ws, 'ERP설변_가공품변동');
        const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

        const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        const filename = `ERP_Component_Changes_${dateStr}.xlsx`;

        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buf);
    } catch (e) {
        res.status(500).send(e.message);
    }
});

// 서버 종료 API
app.post('/api/shutdown', (req, res) => {
    console.log('[api] Shutdown requested. Exiting...');
    res.json({ success: true, message: 'Server is shutting down...' });
    setTimeout(() => {
        process.exit(0);
    }, 1000);
});

// Multer & General Error Handler
app.use((err, req, res, next) => {
    if (err instanceof multer.MulterError) {
        console.error('[Multer Error]', err);
        return res.status(400).json({ success: false, error: `파일 업로드 오류: ${err.message} (${err.code})` });
    }
    console.error('[Global Error]', err);
    res.status(500).json({ success: false, error: `서버 내부 오류: ${err.message}` });
});

let logClients = [];
app.get('/api/logs', (req, res) => {
    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    res.flushHeaders();
    const sendLog = (data) => res.write(`data: ${JSON.stringify(data)}\n\n`);
    logClients.push(sendLog);
    req.on('close', () => logClients = logClients.filter(c => c !== sendLog));
});

function broadcastLog(msg) {
    logClients.forEach(client => client({ msg, time: new Date().toLocaleTimeString() }));
}

const server = app.listen(PORT, '0.0.0.0', () => {
    console.log(`MPS Server LIVE on Port ${PORT}`);
});

server.on('error', (err) => {
    console.error(`[CRITICAL] Server failed to start: ${err.message}`);
    process.exit(1);
});

// 이벤트 루프 강제 유지용 (서버 상시 대기 모드)
setInterval(() => {
    if (!server.listening) {
        console.log('Server not listening, exiting...');
        process.exit(1);
    }
}, 60000);

// [상시 유지 모드] 컴퓨터를 오래 켜두거나 브라우저 탭이 절전 모드로 전환되어도
// 백그라운드 서버가 스스로 꺼지지 않고 상시 연결을 유지합니다.
// (서버 수동 종료는 대시보드의 '서버 종료' 버튼 또는 배치 파일 실행 시 자동 정리됩니다)

