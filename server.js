const express = require('express');
const path = require('path');
const cors = require('cors');
const { exec, spawn } = require('child_process');
const fs = require('fs');
const multer = require('multer');
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
        if (!type || !['1842', '1840'].includes(type)) {
            return res.status(400).json({ success: false, error: '올바른 타입(1842 또는 1840)을 지정해주세요.' });
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
        if (!type || !['1842', '1840'].includes(type)) {
            return res.status(400).json({ success: false, error: '올바른 타입(1842 또는 1840)을 지정해주세요.' });
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
        if (!['1842', '1840'].includes(type)) {
            return res.status(400).json({ success: false, error: '올바른 타입(1842 또는 1840)을 지정해주세요.' });
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

// 이벤트 루프 강제 유지용
setInterval(() => {
    if (!server.listening) {
        console.log('Server not listening, exiting...');
        process.exit(1);
    }
}, 60000);

// Auto-shutdown if no heartbeat is received
const HEARTBEAT_TIMEOUT = 600000; // 10 minutes (prevents shutdown from aggressive browser background throttling)
const GRACE_PERIOD = 120000; // 2 minutes grace period on startup
const startupTime = Date.now();
let lastCheckTime = Date.now();

setInterval(() => {
    const now = Date.now();
    
    // Sleep/wake detection: if the loop was suspended and more than 15 seconds passed (normally 5 seconds)
    if (now - lastCheckTime > 15000) {
        console.log('[SYSTEM] Sleep/wake detected. Resetting heartbeat timer to prevent premature shutdown.');
        lastHeartbeat = now;
    }
    lastCheckTime = now;
    
    if (hasReceivedHeartbeat) {
        if (now - lastHeartbeat > HEARTBEAT_TIMEOUT) {
            console.log('[AUTO-SHUTDOWN] No heartbeat received for 10 minutes. All dashboard pages closed. Shutting down...');
            process.exit(0);
        }
    } else {
        if (now - startupTime > GRACE_PERIOD) {
            console.log('[AUTO-SHUTDOWN] No client connected within grace period. Shutting down...');
            process.exit(0);
        }
    }
}, 5000);
