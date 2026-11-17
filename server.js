const express = require('express');
const path = require('path');
const cors = require('cors');
const { exec, execSync, spawn } = require('child_process');
const fs = require('fs');
const os = require('os');
const zlib = require('zlib');
const crypto = require('crypto');
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

// ==========================================
// [보안 및 사내망 공유 관리자 권한 인증 시스템]
// ==========================================
// 서버 시작 시 무작위 관리자 세션 토큰 생성
const ADMIN_SECRET = crypto.randomBytes(16).toString('hex');
const DEFAULT_ADMIN_PIN = 'dnpc1234';

function getPrimaryLanIp() {
    try {
        const ifaces = os.networkInterfaces();
        for (const name of Object.keys(ifaces)) {
            for (const iface of ifaces[name]) {
                if (iface.family === 'IPv4' && !iface.internal) {
                    return iface.address;
                }
            }
        }
    } catch (e) {}
    return '127.0.0.1';
}

function getLocalIpList() {
    const ips = new Set(['127.0.0.1', '::1', 'localhost', '::ffff:127.0.0.1']);
    try {
        const ifaces = os.networkInterfaces();
        for (const name of Object.keys(ifaces)) {
            for (const iface of ifaces[name]) {
                if (iface.address) {
                    const addr = iface.address.toLowerCase();
                    ips.add(addr);
                    if (addr.startsWith('::ffff:')) {
                        ips.add(addr.replace('::ffff:', ''));
                    }
                }
            }
        }
    } catch (e) {}
    return ips;
}

function cleanIp(ip) {
    if (!ip) return '';
    let str = String(ip).trim().toLowerCase();
    if (str.startsWith('::ffff:')) {
        str = str.replace('::ffff:', '');
    }
    return str;
}

function isLocalClient(req) {
    const rawIp = req.ip || req.connection?.remoteAddress || req.socket?.remoteAddress || '';
    const cleaned = cleanIp(rawIp);
    const localIps = getLocalIpList();
    return localIps.has(cleaned) || cleaned === '127.0.0.1' || cleaned === '::1';
}

function verifyAdmin(req) {
    // 1. 호스트 PC 로컬 접속(localhost, 127.0.0.1, 또는 호스트 본인 IP)은 자동으로 관리자 승인
    if (isLocalClient(req)) return true;
    
    // 2. 외부 사내망에서 관리자 인증을 거친 토큰 확인 (헤더 또는 쿼리)
    const token = req.headers['x-admin-token'] || req.query.admin_token;
    if (token && token === ADMIN_SECRET) return true;
    
    return false;
}

function requireAdmin(req, res, next) {
    if (verifyAdmin(req)) {
        return next();
    }
    return res.status(403).json({
        success: false,
        error: '관리자 권한이 필요합니다. (사내 조회 전용 모드에서는 데이터 동기화 및 서버 관리 기능이 제한됩니다.)'
    });
}

// 관리자 권한 상태 조회 API (클라이언트 UI 분기용)
app.get('/api/auth/status', (req, res) => {
    const isLocal = isLocalClient(req);
    const isAdmin = verifyAdmin(req);
    const clientIp = cleanIp(req.ip || req.connection?.remoteAddress || req.socket?.remoteAddress || '');
    const hostIp = getPrimaryLanIp();
    res.json({
        success: true,
        isAdmin,
        isLocal,
        clientIp,
        hostIp,
        port: PORT
    });
});

// 관리자 PIN 인증 API (원격 접속 시 PIN 번호로 관리자 권한 해제)
app.post('/api/auth/verify-pin', (req, res) => {
    const { pin } = req.body || {};
    const uploadDir = process.pkg ? path.dirname(process.execPath) : __dirname;
    const prefPath = path.join(uploadDir, 'preferences.json');
    let configuredPin = DEFAULT_ADMIN_PIN;
    try {
        if (fs.existsSync(prefPath)) {
            const prefData = JSON.parse(fs.readFileSync(prefPath, 'utf8'));
            if (prefData && prefData.adminPin) {
                configuredPin = String(prefData.adminPin);
            }
        }
    } catch (e) {}

    if (pin && String(pin) === configuredPin) {
        return res.json({ success: true, token: ADMIN_SECRET, message: '관리자 인증 성공' });
    }
    return res.status(401).json({ success: false, error: '관리자 PIN 번호가 일치하지 않습니다.' });
});

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

// 서버 저장 MPS 파일 삭제 API (관리자 전용)
app.post('/api/delete-file', requireAdmin, (req, res) => {
    try {
        const { filename } = req.body || {};
        if (!filename || typeof filename !== 'string') {
            return res.status(400).json({ success: false, error: '삭제할 파일명이 올바르지 않습니다.' });
        }

        // 보안 검증: 경로 조작(Path Traversal) 방지 및 MPS .xlsx 확장자 확인
        const cleanName = path.basename(filename);
        if (cleanName !== filename || !cleanName.startsWith('MPS') || !cleanName.endsWith('.xlsx')) {
            return res.status(400).json({ success: false, error: '유효한 MPS 엑셀(.xlsx) 파일만 삭제할 수 있습니다.' });
        }

        const uploadDir = process.pkg ? path.dirname(process.execPath) : __dirname;
        const targetPath = path.join(uploadDir, cleanName);

        if (!fs.existsSync(targetPath)) {
            return res.status(404).json({ success: false, error: '삭제 대상 파일이 서버에 존재하지 않습니다.' });
        }

        fs.unlinkSync(targetPath);
        console.log(`[api/delete-file] Deleted MPS file: ${cleanName}`);
        res.json({ success: true, message: `${cleanName} 파일이 삭제되었습니다.` });
    } catch (err) {
        console.error('[api/delete-file] Delete failed:', err);
        res.status(500).json({ success: false, error: err.message });
    }
});

// [Live] 서버 사이드 실시간 추출 API (브라우저 메모리 부족 해결용 - 관리자 전용)
app.post('/api/extract-live', requireAdmin, upload.single('file'), async (req, res) => {
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

// ==========================================
// [고속 인메모리 + 디스크 캐시 & Gzip 압축 전송 엔진]
// ==========================================
const SAP_CACHE_DIR = path.join(process.pkg ? path.dirname(process.execPath) : __dirname, '.sap_cache');
if (!fs.existsSync(SAP_CACHE_DIR)) {
    try { fs.mkdirSync(SAP_CACHE_DIR, { recursive: true }); } catch (e) {}
}

const mpsExtractMemoryCache = new Map();
const sapParsedMemoryCache = {};
const sapDiffMemoryCache = new Map();

function sendCompressedJson(req, res, jsonBuf, gzipBuf) {
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    const acceptEncoding = req.headers['accept-encoding'] || '';
    if (gzipBuf && acceptEncoding.includes('gzip')) {
        res.setHeader('Content-Encoding', 'gzip');
        return res.send(gzipBuf);
    }
    return res.send(jsonBuf);
}

function extractMhtmlRowCells(trHtml) {
    const cells = [];
    let p = 0;
    while (true) {
        let tdStart = trHtml.indexOf('<td', p);
        let thStart = trHtml.indexOf('<th', p);
        if (thStart !== -1 && (tdStart === -1 || thStart < tdStart)) tdStart = thStart;
        if (tdStart === -1) break;
        const gt = trHtml.indexOf('>', tdStart);
        if (gt === -1) break;
        let tdEnd = trHtml.indexOf('</td', gt);
        let thEnd = trHtml.indexOf('</th', gt);
        if (thEnd !== -1 && (tdEnd === -1 || thEnd < tdEnd)) tdEnd = thEnd;
        if (tdEnd === -1) break;
        let raw = trHtml.slice(gt + 1, tdEnd);
        if (raw.indexOf('<') !== -1) raw = raw.replace(/<[^>]+>/g, '');
        if (raw.indexOf('&') !== -1) {
            raw = raw
                .replace(/&amp;/g, '&')
                .replace(/&lt;/g, '<')
                .replace(/&gt;/g, '>')
                .replace(/&quot;/g, '"')
                .replace(/&#39;/g, "'")
                .replace(/&nbsp;/g, ' ');
        }
        cells.push(raw.trim());
        p = tdEnd + 5;
    }
    return cells;
}

function forEachMhtmlRow(rawContent, rowCallback) {
    if (!rawContent) return;
    // Check if header declares quoted-printable in first 2KB
    const headChunk = rawContent.slice(0, 2048);
    let content = rawContent;
    if (/Content-Transfer-Encoding:\s*quoted-printable/i.test(headChunk)) {
        content = rawContent
            .replace(/=\r?\n/g, '')
            .replace(/=([0-9A-F]{2})/gi, (_, hex) => String.fromCharCode(parseInt(hex, 16)));
    }
    let pos = 0;
    let rowIdx = 0;
    while (true) {
        const trStart = content.indexOf('<tr', pos);
        if (trStart === -1) break;
        const trOpenEnd = content.indexOf('>', trStart);
        if (trOpenEnd === -1) break;
        const trEnd = content.indexOf('</tr>', trOpenEnd);
        if (trEnd === -1) break;
        const trInner = content.slice(trOpenEnd + 1, trEnd);
        pos = trEnd + 5;
        const cells = extractMhtmlRowCells(trInner);
        if (cells.length > 0) {
            rowCallback(cells, rowIdx);
            rowIdx++;
        }
    }
}

function parseSapMhtmlServerSide(filePath, type) {
    const rawContent = fs.readFileSync(filePath, 'utf8');
    const isComponent = type.startsWith('component');
    const rows = [];
    const previewSampleRows = [];
    let previewHeaders = [];
    let headerCells = null;
    let mapping = null;
    let totalRawRows = 0;

    forEachMhtmlRow(rawContent, (cells, rowIdx) => {
        if (rowIdx === 0) {
            previewHeaders = cells.filter(c => c.length > 0).slice(0, 15);
        } else if (previewSampleRows.length < 10) {
            previewSampleRows.push(cells.slice(0, 15));
        }

        if (!mapping) {
            const upper = cells.map(c => c.toUpperCase());
            const isHeader = isComponent
                ? upper.some(c => c.includes('COMPONENT') || c.includes('COMP.DESC'))
                : upper.some(c => c.includes('MATERIAL') || c.includes('PROD'));

            if (isHeader && rowIdx < 10) {
                headerCells = upper;
                if (isComponent) {
                    const findIdx = (keywords) => {
                        let idx = headerCells.findIndex(cell => keywords.some(k => cell === k));
                        if (idx !== -1) return idx;
                        return headerCells.findIndex(cell => keywords.some(k => cell.includes(k)));
                    };
                    mapping = {
                        mon: findIdx(['PROD.MONTH', 'PROD.MON', '생산월', 'MONTH']),
                        salesDoc: findIdx(['S/O ORDER', 'SALES DOC', '판매문서', 'SALESDOC']),
                        serial: findIdx(['SERIAL NO', 'SERIALNO', '시리얼']),
                        order: findIdx(['ORDER', '오더']),
                        component: findIdx(['COMPONENT', '하위자재', '부품']),
                        compDesc: findIdx(['COMP.DESC.', 'COMP.DESC', '자재명', '부품명']),
                        qty: findIdx(['REQUIREMENT QUANTITY', 'REQUIREMENT QTY', '소요량', '소요']),
                        parentMat: findIdx(['MATERIAL NUMBER', 'MATERIAL', 'PARENT MAT.', '모자재']),
                        reqDate: findIdx(['REQUIREMENT DATE', '소요예정일']),
                        sortString: findIdx(['SORT STRING', '정렬 문자열', '정렬']),
                        deletedItem: findIdx(['DELETED ITEM', 'DELETED', 'DELETE ITEM', '삭제']),
                        wareCtrl: findIdx(['WAREHOUSE CONTROLLER', '창고관리자', '창고 관리자']),
                        plant: findIdx(['PLANT', '플랜트']),
                        procType: findIdx(['PROCUREMENT TYPE', 'PROCUREMENT', '조달유형', '조달']),
                        specProc: findIdx(['SPECIAL PROCUREMENT', '특별조달'])
                    };
                } else {
                    const findIdx = (keywords) => {
                        let idx = headerCells.findIndex(cell => keywords.some(k => cell === k));
                        if (idx !== -1) return idx;
                        return headerCells.findIndex(cell => keywords.some(k => {
                            if (k === 'MATERIAL' && cell.includes('DESCRIPTION')) return false;
                            return cell.includes(k);
                        }));
                    };
                    mapping = {
                        mon: findIdx(['PROD.MON', 'MONTH', '생산월']),
                        model: findIdx(['MATERIAL DESCRIPTION', 'MODEL', '기종']),
                        material: findIdx(['MATERIAL', '품명', '자재코드']),
                        customer: findIdx(['CUSTOMER NAME', 'CUSTOMER', '고객']),
                        serial: findIdx(['SERIAL NO', 'S/O SERIAL', '시리얼']),
                        order: findIdx(['ORDER', '오더', '주문']),
                        salesDoc: findIdx(['SALES DOC', 'SALESDOC', '판매문서']),
                        startDate: findIdx(['START DATE', '시작일', 'STARTDATE']),
                        ver: findIdx(['VER.', 'VERSION', '버전']),
                        deletedItem: findIdx(['DELETED ITEM', 'DELETED', 'DELETE ITEM', '삭제'])
                    };
                }
                return;
            }
        }

        if (!mapping) return;
        totalRawRows++;
        if (cells.length < 5) return;

        // 삭제 품목 (Deleted Item = X) 자동 제외
        if (mapping.deletedItem !== -1 && (cells[mapping.deletedItem] || '').trim().toUpperCase() === 'X') {
            return;
        }

        const monVal = cells[mapping.mon] || '';
        const dateDigits = monVal.replace(/[^0-9]/g, '');
        let mNum = null;
        if (dateDigits.length >= 6) mNum = parseInt(dateDigits.substring(4, 6), 10);
        else if (dateDigits.length >= 1) mNum = parseInt(dateDigits, 10);
        if (!mNum || mNum > 12) return;

        if (isComponent) {
            let wCtrl = '';
            let wDesc = '';
            if (mapping.wareCtrl !== -1) {
                wCtrl = (cells[mapping.wareCtrl] || '').trim();
                if (cells.length > mapping.wareCtrl + 1) {
                    wDesc = (cells[mapping.wareCtrl + 1] || '').trim();
                }
            }
            const sortStr = cells[mapping.sortString] || '';
            const isUnitA = (wCtrl === 'PS' || wCtrl === 'PZ' || wDesc.includes('A급') || sortStr.includes('A급'));
            const plant = mapping.plant !== -1 ? (cells[mapping.plant] || '').trim() : '';
            const procType = mapping.procType !== -1 ? (cells[mapping.procType] || '').trim().toUpperCase() : '';
            const specProc = mapping.specProc !== -1 ? (cells[mapping.specProc] || '').trim().toUpperCase() : '';
            const isInhouse = procType ? (procType === 'E' || (plant === '1842' && !isUnitA)) : (!isUnitA);

            rows.push([
                mNum + '월',
                cells[mapping.salesDoc] || '',
                cells[mapping.serial] || '',
                cells[mapping.order] || '',
                cells[mapping.component] || '',
                cells[mapping.compDesc] || '',
                parseFloat(cells[mapping.qty]) || 0,
                cells[mapping.parentMat] || '',
                cells[mapping.reqDate] || '',
                sortStr,
                wCtrl,
                wDesc,
                isUnitA ? 1 : 0,
                plant,
                procType,
                specProc,
                isInhouse ? 1 : 0
            ]);
        } else {
            rows.push([
                mNum + '월',
                cells[mapping.model] || '',
                cells[mapping.material] || '',
                cells[mapping.customer] || '',
                cells[mapping.serial] || '',
                cells[mapping.order] || '',
                cells[mapping.salesDoc] || '',
                cells[mapping.startDate] || '',
                cells[mapping.ver] || ''
            ]);
        }
    });

    return {
        format: isComponent ? 'compact_comp_v1' : 'compact_plan_v1',
        totalRows: totalRawRows,
        previewHeaders,
        previewSampleRows,
        rows
    };
}

function getOrBuildSapCache(type) {
    const uploadDir = process.pkg ? path.dirname(process.execPath) : __dirname;
    const filePath = path.join(uploadDir, `sap_${type}.mhtml`);
    if (!fs.existsSync(filePath)) {
        delete sapParsedMemoryCache[type];
        return null;
    }

    const stat = fs.statSync(filePath);
    const mem = sapParsedMemoryCache[type];
    if (mem && mem.mtimeMs === stat.mtimeMs && mem.size === stat.size) {
        return mem;
    }

    const diskCachePath = path.join(SAP_CACHE_DIR, `sap_${type}.cache.json`);
    if (fs.existsSync(diskCachePath)) {
        try {
            const diskData = JSON.parse(fs.readFileSync(diskCachePath, 'utf8'));
            if (diskData && diskData.mtimeMs === stat.mtimeMs && diskData.size === stat.size && Array.isArray(diskData.rows)) {
                const apiPayload = {
                    success: true,
                    exists: true,
                    mtime: stat.mtime,
                    mtimeMs: stat.mtimeMs,
                    format: diskData.format,
                    rows: diskData.rows
                };
                const jsonBuf = Buffer.from(JSON.stringify(apiPayload), 'utf8');
                const gzipBuf = zlib.gzipSync(jsonBuf);
                const entry = {
                    mtimeMs: stat.mtimeMs,
                    mtime: stat.mtime,
                    size: stat.size,
                    sizeStr: (stat.size / 1024 / 1024).toFixed(2) + ' MB',
                    totalRows: diskData.totalRows || diskData.rows.length,
                    previewHeaders: diskData.previewHeaders || [],
                    previewSampleRows: diskData.previewSampleRows || [],
                    format: diskData.format,
                    rows: diskData.rows,
                    jsonBuffer: jsonBuf,
                    gzipBuffer: gzipBuf
                };
                sapParsedMemoryCache[type] = entry;
                console.log(`[SAP Cache] Loaded pre-compiled cache from disk for sap_${type}.mhtml (${entry.rows.length} rows, gzip: ${(gzipBuf.length / 1024).toFixed(0)} KB)`);
                return entry;
            }
        } catch (e) {
            console.warn(`[SAP Cache] Disk cache read failed for ${type}, rebuilding:`, e.message);
        }
    }

    const t0 = Date.now();
    console.log(`[SAP Cache] Compiling sap_${type}.mhtml (${(stat.size / 1024 / 1024).toFixed(2)} MB)...`);
    const parsed = parseSapMhtmlServerSide(filePath, type);

    const diskPayload = {
        mtimeMs: stat.mtimeMs,
        size: stat.size,
        format: parsed.format,
        totalRows: parsed.totalRows,
        previewHeaders: parsed.previewHeaders,
        previewSampleRows: parsed.previewSampleRows,
        rows: parsed.rows
    };
    try {
        fs.writeFileSync(diskCachePath, JSON.stringify(diskPayload), 'utf8');
    } catch (e) {
        console.warn(`[SAP Cache] Failed to write disk cache for ${type}:`, e.message);
    }

    const apiPayload = {
        success: true,
        exists: true,
        mtime: stat.mtime,
        mtimeMs: stat.mtimeMs,
        format: parsed.format,
        rows: parsed.rows
    };
    const jsonBuf = Buffer.from(JSON.stringify(apiPayload), 'utf8');
    const gzipBuf = zlib.gzipSync(jsonBuf);

    const entry = {
        mtimeMs: stat.mtimeMs,
        mtime: stat.mtime,
        size: stat.size,
        sizeStr: (stat.size / 1024 / 1024).toFixed(2) + ' MB',
        totalRows: parsed.totalRows,
        previewHeaders: parsed.previewHeaders,
        previewSampleRows: parsed.previewSampleRows,
        format: parsed.format,
        rows: parsed.rows,
        jsonBuffer: jsonBuf,
        gzipBuffer: gzipBuf
    };
    sapParsedMemoryCache[type] = entry;
    console.log(`[SAP Cache] Compiled sap_${type}.mhtml in ${Date.now() - t0}ms (${parsed.rows.length} active rows, gzip: ${(gzipBuf.length / 1024).toFixed(0)} KB)`);
    return entry;
}

// [Saved] 서버에 저장된 파일 직접 추출 API (인메모리 + Gzip 캐시 적용)
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

        const stat = fs.statSync(filePath);
        const rulesKey = crypto.createHash('md5').update(JSON.stringify(rules || {})).digest('hex');
        const cacheKey = `${filename}:${stat.mtimeMs}:${stat.size}:${rulesKey}`;

        if (mpsExtractMemoryCache.has(cacheKey)) {
            const cached = mpsExtractMemoryCache.get(cacheKey);
            return sendCompressedJson(req, res, cached.jsonBuffer, cached.gzipBuffer);
        }
        
        console.log(`[api] Saved file extract started: ${filename}`);
        const fileBuffer = fs.readFileSync(filePath);
        const result = await processMpsFile(fileBuffer, rules);
        
        const responseObj = { success: true, ...result };
        const jsonBuffer = Buffer.from(JSON.stringify(responseObj), 'utf8');
        const gzipBuffer = zlib.gzipSync(jsonBuffer);

        if (mpsExtractMemoryCache.size > 15) {
            const firstKey = mpsExtractMemoryCache.keys().next().value;
            mpsExtractMemoryCache.delete(firstKey);
        }
        mpsExtractMemoryCache.set(cacheKey, { jsonBuffer, gzipBuffer });

        console.log(`[api] Saved file extract success: ${result.finalResults.length} rows (cached)`);
        return sendCompressedJson(req, res, jsonBuffer, gzipBuffer);
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

// 개인화 설정 저장 API (관리자 전용 - 기존 설정과 병합 저장)
app.post('/api/preferences', requireAdmin, (req, res) => {
    try {
        const uploadDir = process.pkg ? path.dirname(process.execPath) : __dirname;
        const prefPath = path.join(uploadDir, 'preferences.json');
        let existing = {};
        if (fs.existsSync(prefPath)) {
            try {
                existing = JSON.parse(fs.readFileSync(prefPath, 'utf8')) || {};
            } catch (e) {}
        }
        const updated = { ...existing, ...(req.body || {}) };
        fs.writeFileSync(prefPath, JSON.stringify(updated, null, 2), 'utf8');
        if (req.body && (req.body.siteMasterText !== undefined || req.body.ruleMasterText !== undefined)) {
            mpsExtractMemoryCache.clear();
        }
        res.json({ success: true, preferences: updated });
    } catch (err) {
        console.error('[api] Save preferences failed:', err.message);
        res.status(500).json({ success: false, error: err.message });
    }
});

// SAP 파일 업로드 API (덮어쓰기 - 관리자 전용)
app.post('/api/upload-sap', requireAdmin, upload.single('file'), (req, res) => {
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
        delete sapParsedMemoryCache[type];
        sapDiffMemoryCache.clear();
        setTimeout(() => { try { getOrBuildSapCache(type); } catch (e) {} }, 50);
        console.log(`[sap-upload] Saved sap_${type}.mhtml to server`);
        res.json({ success: true });
    } catch (err) {
        console.error('[sap-upload] Failed to save SAP file:', err);
        res.status(500).json({ success: false, error: err.message });
    }
});

// SAP 파일 삭제 API (관리자 전용)
app.post('/api/clear-sap', requireAdmin, (req, res) => {
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
        delete sapParsedMemoryCache[type];
        sapDiffMemoryCache.clear();
        const diskCachePath = path.join(SAP_CACHE_DIR, `sap_${type}.cache.json`);
        if (fs.existsSync(diskCachePath)) {
            try { fs.unlinkSync(diskCachePath); } catch (e) {}
        }
        res.json({ success: true });
    } catch (err) {
        console.error('[sap-clear] Failed to clear SAP file:', err);
        res.status(500).json({ success: false, error: err.message });
    }
});

// SAP 파일 로드 API (초고속 사전 파싱 + Gzip 압축 캐시 전송)
app.get('/api/load-sap/:type', (req, res) => {
    try {
        const { type } = req.params;
        if (!['1842', '1840', 'component_1842', 'component_1840'].includes(type)) {
            return res.status(400).json({ success: false, error: '올바른 타입(1842, 1840, component_1842 또는 component_1840)을 지정해주세요.' });
        }
        const cached = getOrBuildSapCache(type);
        if (!cached) {
            return res.json({ success: true, exists: false });
        }

        // 클라이언트가 이미 동일한 mtimeMs 데이터를 보유 중이면 전송 생략
        const clientMtime = Number(req.query.mtimeMs || 0);
        if (clientMtime && Math.abs(clientMtime - cached.mtimeMs) < 1) {
            return res.json({
                success: true,
                exists: true,
                unchanged: true,
                mtime: cached.mtime,
                mtimeMs: cached.mtimeMs
            });
        }

        return sendCompressedJson(req, res, cached.jsonBuffer, cached.gzipBuffer);
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

// SAP ERP 자동 동기화 API (관리자 전용)
const { runSapSync, getSyncStatus } = require('./sap_automation/sync_engine');

app.post('/api/sync-sap', requireAdmin, async (req, res) => {
    try {
        const { startMonth, endMonth } = req.body || {};
        const status = getSyncStatus();
        if (status.running) {
            return res.json({ success: true, message: '이미 SAP 동기화가 진행 중입니다.', status });
        }
        // Run in background so request doesn't timeout
        runSapSync({ startMonth, endMonth }).then(() => {
            sapDiffMemoryCache.clear();
            ['1842', '1840', 'component_1842', 'component_1840'].forEach(t => {
                try { getOrBuildSapCache(t); } catch (e) {}
            });
            // SAP 동기화 완료 직후 변동 내역 자동 메일 발송 (변동 0건일 때는 자동 제외)
            setTimeout(() => {
                sendSapDiffEmail({ isAutoAfterSync: true }).catch(err => {
                    console.error('[SAP Diff Auto Email] Error:', err.message);
                });
            }, 1000);
        }).catch(e => {
            console.error('[SAP Sync Error in background]:', e.message);
        });
        res.json({ success: true, message: 'SAP 동기화가 시작되었습니다.' });
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

const fnToTypeMap = {
    'sap_1840.mhtml': '1840',
    'sap_1842.mhtml': '1842',
    'sap_component_1840.mhtml': 'component_1840',
    'sap_component_1842.mhtml': 'component_1842'
};

function getSapFileStats(fn, fp) {
    if (!fs.existsSync(fp)) return null;
    const type = fnToTypeMap[fn];
    if (type) {
        const cached = getOrBuildSapCache(type);
        if (cached) {
            return {
                size: cached.sizeStr,
                mtime: cached.mtime,
                mtimeMs: cached.mtimeMs,
                rows: cached.totalRows
            };
        }
    }
    const stat = fs.statSync(fp);
    return {
        size: (stat.size / 1024 / 1024).toFixed(2) + ' MB',
        mtime: stat.mtime,
        mtimeMs: stat.mtimeMs,
        rows: 0
    };
}

app.get('/api/sync-sap/status', (req, res) => {
    try {
        const status = getSyncStatus();
        const files = {};
        let maxMtimeMs = 0;
        let minMtimeMs = Infinity;
        ['sap_1840.mhtml', 'sap_1842.mhtml', 'sap_component_1840.mhtml', 'sap_component_1842.mhtml'].forEach(fn => {
            const fp = path.join(__dirname, fn);
            const stat = getSapFileStats(fn, fp);
            files[fn] = stat;
            if (stat && stat.mtimeMs) {
                if (stat.mtimeMs > maxMtimeMs) maxMtimeMs = stat.mtimeMs;
                if (stat.mtimeMs < minMtimeMs) minMtimeMs = stat.mtimeMs;
            }
        });
        if (!status.running && maxMtimeMs > 0) {
            const finishedMs = status.finishedAt ? new Date(status.finishedAt).getTime() : 0;
            if (!finishedMs || maxMtimeMs > finishedMs + 5000) {
                status.finishedAt = new Date(maxMtimeMs).toISOString();
                if (!status.startedAt && minMtimeMs < maxMtimeMs && (maxMtimeMs - minMtimeMs) < 6 * 3600 * 1000) {
                    status.startedAt = new Date(minMtimeMs).toISOString();
                }
            }
        }
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
            '1840': { name: '1840 남산+ 생산계획', file: 'sap_1840.mhtml', type: '1840' },
            'comp_1840': { name: '1840 남산+ 부품소요량', file: 'sap_component_1840.mhtml', type: 'component_1840' },
            '1842': { name: '1842 성주 생산계획', file: 'sap_1842.mhtml', type: '1842' },
            'comp_1842': { name: '1842 성주 부품소요량', file: 'sap_component_1842.mhtml', type: 'component_1842' }
        };
        const result = {};
        for (const [key, info] of Object.entries(targetFiles)) {
            const cached = getOrBuildSapCache(info.type);
            if (!cached) {
                result[key] = { exists: false, name: info.name, file: info.file };
                continue;
            }
            result[key] = {
                exists: true,
                name: info.name,
                file: info.file,
                size: cached.sizeStr,
                mtime: cached.mtime,
                totalRows: cached.totalRows,
                headers: cached.previewHeaders || [],
                sampleRows: cached.previewSampleRows || []
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

function parsePlanMhtml(filePath) {
    const orderToSerial = new Map();
    const soToSerial = new Map();
    const orderToMonth = new Map();
    const orderToMatDesc = new Map();
    if (!fs.existsSync(filePath)) return { orderToSerial, soToSerial, orderToMonth, orderToMatDesc };

    try {
        const content = fs.readFileSync(filePath, 'utf8');
        let idxSerial = -1, idxSO = -1, idxOrder = -1, idxMonth = -1, idxMatDesc = -1;
        let hasHeaders = false;

        forEachMhtmlRow(content, (cells, rowIdx) => {
            if (!hasHeaders && rowIdx < 5) {
                const headers = cells.map(c => c.toUpperCase());
                if (headers.some(h => h.includes('MATERIAL') || h.includes('ORDER') || h.includes('SERIAL'))) {
                    idxSerial = headers.findIndex(h => h.includes('SERIAL') || h.includes('호기'));
                    idxSO = headers.findIndex(h => h.includes('SALES DOC') || h.includes('S/O') || h.includes('SALES'));
                    idxOrder = headers.findIndex(h => h === 'ORDER' || (h.includes('ORDER') && !h.includes('TYPE') && !h.includes('S/O')));
                    idxMonth = headers.findIndex(h => h.includes('PROD.MON') || h.includes('PROD.MONTH') || h.includes('생산월'));
                    idxMatDesc = headers.findIndex(h => h.includes('MATERIAL DESCRIPTION') || h.includes('자재내역'));
                    hasHeaders = true;
                    return;
                }
            }
            if (!hasHeaders) return;

            const order = idxOrder !== -1 ? (cells[idxOrder] || '') : '';
            const so = idxSO !== -1 ? (cells[idxSO] || '') : '';
            const serial = idxSerial !== -1 ? (cells[idxSerial] || '') : '';
            const month = idxMonth !== -1 ? (cells[idxMonth] || '') : '';
            const matDesc = idxMatDesc !== -1 ? (cells[idxMatDesc] || '') : '';

            if (order && serial) orderToSerial.set(order, serial);
            if (so && so !== '0' && so !== '-' && serial) soToSerial.set(so, serial);
            if (order && month) orderToMonth.set(order, month);
            if (order && matDesc) orderToMatDesc.set(order, matDesc);
        });
    } catch (e) {
        console.error('[SAP Diff] parsePlanMhtml error:', e.message);
    }
    return { orderToSerial, soToSerial, orderToMonth, orderToMatDesc };
}

function parseComponentMhtml(filePath, planData = null) {
    if (!fs.existsSync(filePath)) return [];
    const content = fs.readFileSync(filePath, 'utf8');

    let idxPlant = -1, idxSerial = -1, idxSO = -1, idxOrder = -1, idxMonth = -1;
    let idxMat = -1, idxMatDesc = -1, idxComp = -1, idxCompDesc = -1, idxQty = -1;
    let idxUnit = -1, idxDel = -1, idxWareCtrl = -1, idxSortStr = -1;
    let idxInvName = -1, idxInvCode = -1;
    let hasHeaders = false;
    const records = [];

    forEachMhtmlRow(content, (cells, rowIdx) => {
        if (!hasHeaders && rowIdx < 10) {
            const headers = cells.map(c => c.toUpperCase());
            if (headers.some(h => h.includes('COMPONENT') || h.includes('COMP.DESC'))) {
                idxPlant = headers.findIndex(h => h.includes('PLANT') || h.includes('플랜트'));
                idxSerial = headers.findIndex(h => h.includes('SERIAL') || h.includes('호기'));
                idxSO = headers.findIndex(h => h.includes('S/O') || h.includes('SALES'));
                idxOrder = headers.findIndex(h => h === 'ORDER' || (h.includes('ORDER') && !h.includes('TYPE') && !h.includes('S/O')));
                idxMonth = headers.findIndex(h => h.includes('PROD.MONTH') || h.includes('PROD.MON') || h.includes('생산월'));
                idxMat = headers.findIndex(h => h.includes('MATERIAL NUMBER') || (h.includes('MATERIAL') && !h.includes('DESCRIPTION')) || h.includes('자재'));
                idxMatDesc = headers.findIndex(h => h.includes('MATERIAL DESCRIPTION') || h.includes('자재내역'));
                idxComp = headers.findIndex(h => h === 'COMPONENT' || h.includes('구성부품'));
                idxCompDesc = headers.findIndex(h => h.includes('COMP.DESC') || h.includes('구성부품내역'));
                idxQty = headers.findIndex(h => h.includes('REQUIREMENT QUANTITY') || h.includes('REQUIREMENT QTY') || h.includes('소요량'));
                idxUnit = headers.findIndex(h => h.includes('BASE UNIT') || h.includes('단위'));
                idxDel = headers.findIndex(h => h.includes('DELETED ITEM') || h.includes('DELETED') || h.includes('삭제'));
                idxWareCtrl = headers.findIndex(h => h.includes('WAREHOUSE CONTROLLER') || h.includes('창고관리자') || h.includes('창고 관리자'));
                idxSortStr = headers.findIndex(h => h.includes('SORT STRING') || h.includes('정렬 문자열') || h.includes('정렬'));
                const invIndices = headers
                    .map((h, i) => (h.includes('INV.MANAGER') || h.includes('INV. MANAGER') || h.includes('INVENTORY MANAGER') || h.includes('재고관리자') || h.includes('담당자')) ? i : -1)
                    .filter(i => i !== -1);
                if (invIndices.length > 0) idxInvName = invIndices[0];
                if (invIndices.length > 1) idxInvCode = invIndices[1];
                hasHeaders = true;
                return;
            }
        }
        if (!hasHeaders) return;

        const comp = cells[idxComp] || '';
        if (!comp) return;

        // SAP 삭제 품목 (Deleted Item = 'X') 필터링
        if (idxDel !== -1) {
            const delVal = (cells[idxDel] || '').toUpperCase();
            if (delVal === 'X' || delVal === 'TRUE') return;
        }

        let wareCtrl = '';
        let wareCtrlDesc = '';
        if (idxWareCtrl !== -1) {
            wareCtrl = (cells[idxWareCtrl] || '').trim();
            if (cells.length > idxWareCtrl + 1) {
                wareCtrlDesc = (cells[idxWareCtrl + 1] || '').trim();
            }
        }
        const sortStr = idxSortStr !== -1 ? (cells[idxSortStr] || '').trim() : '';
        const isUnitA = (wareCtrl === 'PS' || wareCtrl === 'PZ' || wareCtrlDesc.includes('A급') || sortStr.includes('A급'));

        // A급 유니트(PS, PZ, A급)는 하위 소요량(가공품) 변동 분석 대상에서 조기 제외 (성능 극대화)
        if (isUnitA || wareCtrl === 'PS' || wareCtrl === 'PZ') return;

        const qty = parseFloat((cells[idxQty] || '0').replace(/,/g, '')) || 0;
        if (qty <= 0) return;

        const plant = cells[idxPlant] || '';
        const soOrder = cells[idxSO] || '';
        const order = cells[idxOrder] || '';
        let serial = cells[idxSerial] || '';
        let prodMonth = cells[idxMonth] || '';
        let matDesc = cells[idxMatDesc] || '';

        if (planData) {
            if (!serial || serial === '-') {
                const validSo = (soOrder && soOrder !== '0' && soOrder !== '-') ? soOrder : '';
                serial = planData.orderToSerial?.get(order) || (validSo ? planData.soToSerial?.get(validSo) : '') || '';
            }
            if (!prodMonth && planData.orderToMonth) {
                prodMonth = planData.orderToMonth.get(order) || '';
            }
            if (!matDesc && planData.orderToMatDesc) {
                matDesc = planData.orderToMatDesc.get(order) || '';
            }
        }

        const invManagerName = idxInvName !== -1 ? (cells[idxInvName] || '').trim() : '';
        const invManagerCode = idxInvCode !== -1 ? (cells[idxInvCode] || '').trim() : '';
        const invManager = invManagerName && invManagerCode
            ? `${invManagerName} (${invManagerCode})`
            : (invManagerName || invManagerCode || '');

        records.push({
            plant,
            serial,
            soOrder,
            order,
            prodMonth,
            material: cells[idxMat] || '',
            matDesc,
            component: comp,
            compDesc: cells[idxCompDesc] || '',
            qty,
            unit: cells[idxUnit] || '',
            wareCtrl,
            wareCtrlDesc,
            sortStr,
            isUnitA,
            invManagerName,
            invManagerCode,
            invManager
        });
    });
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
    const plan1840 = path.join(dir, 'sap_1840.mhtml');
    const plan1842 = path.join(dir, 'sap_1842.mhtml');

    const mtimeSig = [f1840, f1842, plan1840, plan1842].map(p => {
        try { return fs.existsSync(p) ? fs.statSync(p).mtimeMs : 0; } catch (e) { return 0; }
    }).join('_');

    const memKey = `v3:${snapshotId}:${mtimeSig}`;
    if (sapDiffMemoryCache.has(memKey)) {
        return sapDiffMemoryCache.get(memKey);
    }

    const safeSnapId = String(snapshotId || 'current').replace(/[^a-zA-Z0-9_-]/g, '_');
    const diskDiffCachePath = path.join(SAP_CACHE_DIR, `diff_v3_${safeSnapId}.cache.json`);
    if (fs.existsSync(diskDiffCachePath)) {
        try {
            const cachedDisk = JSON.parse(fs.readFileSync(diskDiffCachePath, 'utf8'));
            if (cachedDisk && cachedDisk.mtimeSig === mtimeSig && Array.isArray(cachedDisk.records)) {
                sapDiffMemoryCache.set(memKey, cachedDisk.records);
                return cachedDisk.records;
            }
        } catch (e) {}
    }

    const pData1840 = parsePlanMhtml(plan1840);
    const pData1842 = parsePlanMhtml(plan1842);

    const records = [];
    if (fs.existsSync(f1840)) records.push(...parseComponentMhtml(f1840, pData1840));
    if (fs.existsSync(f1842)) records.push(...parseComponentMhtml(f1842, pData1842));

    try {
        fs.writeFileSync(diskDiffCachePath, JSON.stringify({ mtimeSig, records }), 'utf8');
    } catch (e) {}
    sapDiffMemoryCache.set(memKey, records);
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

function scoreComponentSimilarity(delItem, newItem) {
    const compA = (delItem.component || '').trim().toUpperCase();
    const compB = (newItem.component || '').trim().toUpperCase();
    const prefixA = compA.split('-')[0];
    const prefixB = compB.split('-')[0];
    const samePrefix = prefixA && prefixA.length >= 4 && prefixA === prefixB;

    const descA = (delItem.compDesc || '').trim().toUpperCase();
    const descB = (newItem.compDesc || '').trim().toUpperCase();
    const mainA = descA.split(/[;,]/)[0].trim();
    const mainB = descB.split(/[;,]/)[0].trim();
    const subA = descA.split(/[;,]/).slice(1).join(' ').trim();
    const subB = descB.split(/[;,]/).slice(1).join(' ').trim();

    // Check directional/positional conflicts (e.g., LEFT vs RIGHT, UP vs BOTTOM, UPPER vs LOWER)
    const hasDirConflict = (a, b, w1, w2) => {
        const re1 = new RegExp(`\\b${w1}\\b`);
        const re2 = new RegExp(`\\b${w2}\\b`);
        return (re1.test(a) && re2.test(b)) || (re2.test(a) && re1.test(b));
    };
    if (
        hasDirConflict(descA, descB, 'LEFT', 'RIGHT') ||
        hasDirConflict(descA, descB, 'LH', 'RH') ||
        hasDirConflict(descA, descB, 'UP', 'BOTTOM') ||
        hasDirConflict(descA, descB, 'UPPER', 'LOWER') ||
        hasDirConflict(descA, descB, 'FRONT', 'REAR')
    ) {
        return -100;
    }

    let score = 0;
    if (descA && descA === descB) {
        score += 200;
    } else if (descA && descB && (descA.startsWith(descB) || descB.startsWith(descA))) {
        score += 150;
    }

    if (mainA && mainA === mainB) {
        score += 70;
    } else if (mainA && mainB && (mainA.includes(mainB) || mainB.includes(mainA))) {
        score += 40;
    }

    if (samePrefix) {
        score += 50;
    }

    const tokensA = subA.split(/[\s_(),/&;-]+/).filter(Boolean);
    const tokensB = subB.split(/[\s_(),/&;-]+/).filter(Boolean);
    const usedB = new Set();
    for (const tA of tokensA) {
        for (let j = 0; j < tokensB.length; j++) {
            if (usedB.has(j)) continue;
            const tB = tokensB[j];
            if (tA === tB) {
                score += 25;
                usedB.add(j);
                break;
            } else if (tA.length >= 3 && tB.length >= 3 && (tA.includes(tB) || tB.includes(tA))) {
                score += 20;
                usedB.add(j);
                break;
            }
        }
    }

    return score;
}

function pairComponentReplacements(changes) {
    // Group by machine (plant + serial, or plant + order if serial is absent)
    const groups = new Map();
    changes.forEach((c, idx) => {
        const validSo = (c.soOrder && c.soOrder !== '0' && c.soOrder !== '-') ? c.soOrder : '';
        const machineKey = `${c.plant || ''}|${c.serial || validSo || c.order || c.matDesc || idx}`;
        if (!groups.has(machineKey)) {
            groups.set(machineKey, []);
        }
        groups.get(machineKey).push(c);
    });

    const groupedChanges = [];

    for (const [, items] of groups) {
        const deletedList = items.filter(i => i.changeType === 'DELETED');
        const newList = items.filter(i => i.changeType === 'NEW');
        const qtyList = items.filter(i => i.changeType === 'QTY_CHANGE');

        const candidatePairs = [];
        deletedList.forEach((delItem, dIdx) => {
            newList.forEach((newItem, nIdx) => {
                const score = scoreComponentSimilarity(delItem, newItem);
                if (score >= 70) {
                    candidatePairs.push({ dIdx, nIdx, score });
                }
            });
        });

        // Sort candidate pairs by score descending for greedy best-first 1:1 matching
        candidatePairs.sort((a, b) => b.score - a.score);

        const matchedDel = new Set();
        const matchedNew = new Set();
        const machineResults = [];

        for (const pair of candidatePairs) {
            if (matchedDel.has(pair.dIdx) || matchedNew.has(pair.nIdx)) continue;
            matchedDel.add(pair.dIdx);
            matchedNew.add(pair.nIdx);

            const delItem = deletedList[pair.dIdx];
            const newItem = newList[pair.nIdx];
            machineResults.push({
                changeType: 'REPLACED',
                plant: newItem.plant || delItem.plant,
                soOrder: newItem.soOrder || delItem.soOrder,
                serial: newItem.serial || delItem.serial,
                order: newItem.order || delItem.order,
                prodMonth: newItem.prodMonth || delItem.prodMonth,
                material: newItem.material || delItem.material,
                matDesc: newItem.matDesc || delItem.matDesc,
                oldComponent: delItem.component,
                component: newItem.component,
                oldCompDesc: delItem.compDesc,
                compDesc: newItem.compDesc,
                unit: newItem.unit || delItem.unit,
                wareCtrl: newItem.wareCtrl || delItem.wareCtrl || '',
                wareCtrlDesc: newItem.wareCtrlDesc || delItem.wareCtrlDesc || '',
                isUnitA: !!(newItem.isUnitA || delItem.isUnitA),
                oldInvManager: delItem.invManager || '',
                invManager: newItem.invManager || '',
                invManagerName: newItem.invManagerName || '',
                invManagerCode: newItem.invManagerCode || '',
                oldQty: delItem.oldQty,
                newQty: newItem.newQty,
                diffQty: newItem.newQty - delItem.oldQty,
                matchScore: pair.score
            });
        }

        // Add remaining unmatched NEW, DELETED, and QTY_CHANGE items for this machine
        newList.forEach((newItem, nIdx) => {
            if (!matchedNew.has(nIdx)) machineResults.push(newItem);
        });
        deletedList.forEach((delItem, dIdx) => {
            if (!matchedDel.has(dIdx)) machineResults.push(delItem);
        });
        qtyList.forEach(qItem => {
            machineResults.push(qItem);
        });

        // Sort items within the same machine: REPLACED first, then by compDesc
        const typeOrder = { REPLACED: 1, NEW: 2, DELETED: 3, QTY_CHANGE: 4 };
        machineResults.sort((a, b) => {
            const tDiff = (typeOrder[a.changeType] || 9) - (typeOrder[b.changeType] || 9);
            if (tDiff !== 0) return tDiff;
            return (a.compDesc || '').localeCompare(b.compDesc || '', 'ko');
        });

        groupedChanges.push(...machineResults);
    }

    // Sort all grouped changes by plant -> serial/order -> type -> compDesc
    groupedChanges.sort((a, b) => {
        if (a.plant !== b.plant) return (a.plant || '').localeCompare(b.plant || '');
        const aHasSerial = a.serial ? 0 : 1;
        const bHasSerial = b.serial ? 0 : 1;
        if (aHasSerial !== bHasSerial) return aHasSerial - bHasSerial;
        const keyA = a.serial || a.order || a.soOrder || '';
        const keyB = b.serial || b.order || b.soOrder || '';
        if (keyA !== keyB) return keyA.localeCompare(keyB, 'ko');
        return (a.compDesc || '').localeCompare(b.compDesc || '', 'ko');
    });

    return groupedChanges;
}

function calculateComponentDiff(recordsA, recordsB) {
    const getAffectedOrderKey = (r) => {
        if (r.serial && r.serial !== '-') return `S:${r.serial}`;
        const validSo = (r.soOrder && r.soOrder !== '0' && r.soOrder !== '-') ? r.soOrder : '';
        if (validSo) return `SO:${validSo}`;
        return `ORD:${r.order || r.matDesc || ''}`;
    };

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
                wareCtrl: itemB.wareCtrl || '',
                wareCtrlDesc: itemB.wareCtrlDesc || '',
                isUnitA: !!itemB.isUnitA,
                invManager: itemB.invManager || '',
                invManagerName: itemB.invManagerName || '',
                invManagerCode: itemB.invManagerCode || '',
                oldQty: 0,
                newQty: itemB.qty,
                diffQty: itemB.qty
            });
            ordersAffected.add(getAffectedOrderKey(itemB));
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
                    wareCtrl: itemB.wareCtrl || itemA.wareCtrl || '',
                    wareCtrlDesc: itemB.wareCtrlDesc || itemA.wareCtrlDesc || '',
                    isUnitA: !!(itemB.isUnitA || itemA.isUnitA),
                    invManager: itemB.invManager || itemA.invManager || '',
                    invManagerName: itemB.invManagerName || itemA.invManagerName || '',
                    invManagerCode: itemB.invManagerCode || itemA.invManagerCode || '',
                    oldQty: itemA.qty,
                    newQty: itemB.qty,
                    diffQty: diff
                });
                ordersAffected.add(getAffectedOrderKey(itemB));
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
                wareCtrl: itemA.wareCtrl || '',
                wareCtrlDesc: itemA.wareCtrlDesc || '',
                isUnitA: !!itemA.isUnitA,
                invManager: itemA.invManager || '',
                invManagerName: itemA.invManagerName || '',
                invManagerCode: itemA.invManagerCode || '',
                oldQty: itemA.qty,
                newQty: 0,
                diffQty: -itemA.qty
            });
            ordersAffected.add(getAffectedOrderKey(itemA));
        }
    }

    // Also collect all active components in recordsB (current target dataset) that have NO Inv.Manager registered (초도품 / 담당자 미등록)
    const unregMap = new Map();
    for (const r of recordsB) {
        if (!r.invManager) {
            const uKey = [r.plant, r.soOrder || r.order || '', r.serial || '', r.component].join('|');
            if (!unregMap.has(uKey)) {
                unregMap.set(uKey, {
                    changeType: 'UNREGISTERED_ALL',
                    plant: r.plant,
                    soOrder: r.soOrder,
                    serial: r.serial,
                    order: r.order,
                    prodMonth: r.prodMonth,
                    material: r.material,
                    matDesc: r.matDesc,
                    component: r.component,
                    compDesc: r.compDesc,
                    unit: r.unit,
                    invManager: '',
                    oldQty: 0,
                    newQty: r.qty,
                    diffQty: r.qty
                });
            } else {
                unregMap.get(uKey).newQty += r.qty;
                unregMap.get(uKey).diffQty += r.qty;
            }
        }
    }
    const unregisteredAll = Array.from(unregMap.values()).sort((a, b) => {
        if (a.plant !== b.plant) return (a.plant || '').localeCompare(b.plant || '');
        const aHasSerial = a.serial ? 0 : 1;
        const bHasSerial = b.serial ? 0 : 1;
        if (aHasSerial !== bHasSerial) return aHasSerial - bHasSerial;
        const keyA = a.serial || a.order || a.soOrder || '';
        const keyB = b.serial || b.order || b.soOrder || '';
        if (keyA !== keyB) return keyA.localeCompare(keyB, 'ko');
        return (a.compDesc || '').localeCompare(b.compDesc || '', 'ko');
    });
    const uniqueUnregisteredParts = new Set(unregisteredAll.map(u => u.component)).size;

    // Sort raw changes by plant -> serial/order -> compDesc so even ungrouped view stays organized by machine
    changes.sort((a, b) => {
        if (a.plant !== b.plant) return (a.plant || '').localeCompare(b.plant || '');
        const aHasSerial = a.serial ? 0 : 1;
        const bHasSerial = b.serial ? 0 : 1;
        if (aHasSerial !== bHasSerial) return aHasSerial - bHasSerial;
        const keyA = a.serial || a.order || a.soOrder || '';
        const keyB = b.serial || b.order || b.soOrder || '';
        if (keyA !== keyB) return keyA.localeCompare(keyB, 'ko');
        return (a.compDesc || '').localeCompare(b.compDesc || '', 'ko');
    });

    const groupedChanges = pairComponentReplacements(changes);
    const totalReplaced = groupedChanges.filter(c => c.changeType === 'REPLACED').length;
    const totalPureNew = groupedChanges.filter(c => c.changeType === 'NEW').length;
    const totalPureDeleted = groupedChanges.filter(c => c.changeType === 'DELETED').length;
    const totalDiffUnregistered = groupedChanges.filter(c => c.changeType !== 'DELETED' && !c.invManager).length;

    const summary = {
        totalReplaced,
        totalPureNew,
        totalPureDeleted,
        totalNew: changes.filter(c => c.changeType === 'NEW').length,
        totalDeleted: changes.filter(c => c.changeType === 'DELETED').length,
        totalQtyChanged: changes.filter(c => c.changeType === 'QTY_CHANGE').length,
        totalOrdersAffected: ordersAffected.size,
        totalChanges: changes.length,
        totalGroupedChanges: groupedChanges.length,
        totalDiffUnregistered,
        totalAllUnregisteredRows: unregisteredAll.length,
        uniqueUnregisteredParts
    };

    return { summary, changes, groupedChanges, unregisteredAll };
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
            changes: diffResult.changes,
            groupedChanges: diffResult.groupedChanges,
            unregisteredAll: diffResult.unregisteredAll
        });
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

function buildDiffExcelBuffer(changes) {
    const rows = changes.map(c => ({
        '변동유형': c.changeType === 'REPLACED' ? '품번변경(설변교체)' : (c.changeType === 'UNREGISTERED_ALL' ? '초도/담당자미등록' : (c.changeType === 'NEW' ? '신규 투입' : (c.changeType === 'DELETED' ? '삭제/제외' : '소요량 변동'))),
        '플랜트': c.plant === '1840' ? '1840 (남산+)' : (c.plant === '1842' ? '1842 (성주)' : c.plant),
        'S/O번호': (c.soOrder && c.soOrder !== '0') ? c.soOrder : '-',
        '호기(시리얼)': c.serial || '-',
        '계획오더': c.order || '-',
        '생산월': c.prodMonth || '-',
        '모품번(기종)': c.material || '-',
        '기종내역': c.matDesc || '-',
        '이전 가공품번': c.changeType === 'REPLACED' ? (c.oldComponent || '') : (c.changeType === 'DELETED' ? c.component : '-'),
        '변경/최신 가공품번': c.changeType === 'DELETED' ? '-' : (c.component || ''),
        '가공품명(Comp.Desc)': c.changeType === 'REPLACED' && c.oldCompDesc && c.oldCompDesc !== c.compDesc
            ? `${c.compDesc} (이전: ${c.oldCompDesc})`
            : (c.compDesc || ''),
        '담당자(Inv.Manager)': c.invManager || '🚨 미등록',
        '이전품번 담당자': c.oldInvManager || '-',
        '단위': c.unit || '',
        '이전 소요량': c.oldQty,
        '최신 소요량': c.newQty,
        '소요량 차이(+/-)': c.diffQty
    }));

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = [
        { wch: 16 }, { wch: 14 }, { wch: 12 }, { wch: 16 },
        { wch: 14 }, { wch: 10 }, { wch: 14 }, { wch: 28 },
        { wch: 18 }, { wch: 18 }, { wch: 32 }, { wch: 16 },
        { wch: 16 }, { wch: 8 }, { wch: 12 }, { wch: 12 }, { wch: 14 }
    ];
    XLSX.utils.book_append_sheet(wb, ws, 'ERP설변_가공품변동');
    return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

app.post('/api/sap-diff/export-excel', (req, res) => {
    try {
        const { changes } = req.body;
        if (!Array.isArray(changes)) {
            return res.status(400).send('Invalid changes data');
        }
        const buf = buildDiffExcelBuffer(changes);
        const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        const filename = `ERP_Component_Changes_${dateStr}.xlsx`;

        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buf);
    } catch (e) {
        res.status(500).send(e.message);
    }
});

// ==========================================
// SAP 설변/가공품 변동 내역 Outlook 자동 메일 발송 시스템
// ==========================================
const DEFAULT_DIFF_EMAIL_CONFIG = {
    enabled: true,
    to: 'junsic.choi@dncompany.com',
    cc: '',
    lastSentAt: null,
    lastSentStatus: ''
};

function getDiffEmailConfig() {
    try {
        const uploadDir = process.pkg ? path.dirname(process.execPath) : __dirname;
        const prefPath = path.join(uploadDir, 'preferences.json');
        if (fs.existsSync(prefPath)) {
            const pref = JSON.parse(fs.readFileSync(prefPath, 'utf8')) || {};
            if (pref.diffEmailConfig) {
                return { ...DEFAULT_DIFF_EMAIL_CONFIG, ...pref.diffEmailConfig };
            }
        }
    } catch (e) {}
    return { ...DEFAULT_DIFF_EMAIL_CONFIG };
}

function saveDiffEmailConfig(patch) {
    const uploadDir = process.pkg ? path.dirname(process.execPath) : __dirname;
    const prefPath = path.join(uploadDir, 'preferences.json');
    let existing = {};
    if (fs.existsSync(prefPath)) {
        try { existing = JSON.parse(fs.readFileSync(prefPath, 'utf8')) || {}; } catch (e) {}
    }
    const merged = { ...getDiffEmailConfig(), ...(patch || {}) };
    existing.diffEmailConfig = merged;
    fs.writeFileSync(prefPath, JSON.stringify(existing, null, 2), 'utf8');
    return merged;
}

function escapeHtml(str) {
    return String(str ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function buildDiffEmailHtml(summary, groupedChanges, baseTitle, targetTitle) {
    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    const nowStr = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}`;
    const fontStack = `'Pretendard','Noto Sans KR','Noto Sans CJK KR','LG Smart UI','Segoe UI','Apple SD Gothic Neo','Malgun Gothic',sans-serif`;
    const numFontStack = `'Segoe UI','Pretendard','Noto Sans KR','Malgun Gothic',sans-serif`;

    const rowsHtml = groupedChanges.map((c, idx) => {
        let badgeBg = '#ede9fe', badgeColor = '#5b21b6', badgeText = '품번변경';
        if (c.changeType === 'NEW') {
            badgeBg = '#d1fae5'; badgeColor = '#065f46'; badgeText = '신규투입';
        } else if (c.changeType === 'DELETED') {
            badgeBg = '#fee2e2'; badgeColor = '#991b1b'; badgeText = '삭제/제외';
        } else if (c.changeType === 'QTY_CHANGE') {
            badgeBg = '#fef3c7'; badgeColor = '#92400e'; badgeText = '수량변동';
        }

        const plantBadge = c.plant === '1840'
            ? `<span style="font-family:${fontStack}; color:#1d4ed8; font-weight:700; font-size:10pt;">1840 남산+</span>`
            : (c.plant === '1842' ? `<span style="font-family:${fontStack}; color:#6d28d9; font-weight:700; font-size:10pt;">1842 성주</span>` : escapeHtml(c.plant || '-'));
        const monthSub = c.prodMonth ? `<div style="font-family:${numFontStack}; font-size:9pt; color:#64748b; margin-top:2px;">${escapeHtml(c.prodMonth)}</div>` : '';

        const serialMain = c.serial
            ? `<div style="font-family:${numFontStack}; font-weight:700; color:#0f172a; font-size:10pt; letter-spacing:0;">${escapeHtml(c.serial)}</div>`
            : `<div style="font-family:${fontStack}; color:#64748b; font-size:9.5pt;">호기 미지정</div>`;
        const orderOrSo = (c.soOrder && c.soOrder !== '0' && c.soOrder !== '-')
            ? `S/O ${escapeHtml(c.soOrder)}`
            : (c.order ? `오더 ${escapeHtml(c.order)}` : '');
        const orderSub = orderOrSo ? `<div style="font-family:${numFontStack}; font-size:9pt; color:#64748b; margin-top:2px;">${orderOrSo}</div>` : '';

        const matTitle = c.matDesc ? escapeHtml(c.matDesc) : escapeHtml(c.material || '-');
        const matSub = (c.material && c.material !== 'SEMI' && c.matDesc)
            ? `<div style="font-family:${numFontStack}; font-size:9pt; color:#64748b; margin-top:2px;">${escapeHtml(c.material)}</div>`
            : '';

        let partNumHtml = '';
        if (c.changeType === 'REPLACED') {
            partNumHtml = `
                <div style="font-family:${numFontStack}; font-size:9.5pt; color:#94a3b8; text-decoration:line-through;">${escapeHtml(c.oldComponent || '-')}</div>
                <div style="font-family:${numFontStack}; font-size:10.5pt; font-weight:700; color:#0f172a; margin-top:3px;">
                    <span style="color:#7c3aed; font-weight:700;">➔</span> <span style="background-color:#f5f3ff; color:#5b21b6; padding:2px 6px; font-weight:700;">${escapeHtml(c.component || '-')}</span>
                </div>`;
        } else if (c.changeType === 'DELETED') {
            partNumHtml = `<div style="font-family:${numFontStack}; font-size:10pt; font-weight:700; color:#dc2626; text-decoration:line-through;">${escapeHtml(c.component || '-')}</div>`;
        } else {
            partNumHtml = `<div style="font-family:${numFontStack}; font-size:10.5pt; font-weight:700; color:#0284c7;">${escapeHtml(c.component || '-')}</div>`;
        }

        let partDescHtml = `<div style="font-family:${fontStack}; font-weight:600; color:#0f172a; font-size:10pt;">${escapeHtml(c.compDesc || '-')}</div>`;
        if (c.changeType === 'REPLACED' && c.oldCompDesc && c.oldCompDesc !== c.compDesc) {
            partDescHtml += `<div style="font-family:${fontStack}; font-size:9pt; color:#64748b; margin-top:3px;">이전 품명: ${escapeHtml(c.oldCompDesc)}</div>`;
        }

        const isUnreg = c.changeType !== 'DELETED' && !c.invManager;
        let invHtml = '';
        if (c.invManager) {
            invHtml = `<span style="font-family:${fontStack}; background-color:#eff6ff; color:#1d4ed8; padding:3px 9px; font-weight:700; font-size:9.5pt;">${escapeHtml(c.invManager)}</span>`;
            if (c.changeType === 'REPLACED' && c.oldInvManager && c.oldInvManager !== c.invManager) {
                invHtml += `<div style="font-family:${fontStack}; font-size:8.5pt; color:#64748b; margin-top:3px;">이전: ${escapeHtml(c.oldInvManager)}</div>`;
            }
        } else if (c.changeType === 'DELETED') {
            invHtml = `<span style="font-family:${fontStack}; color:#94a3b8; font-size:9.5pt;">${escapeHtml(c.oldInvManager || '-')}</span>`;
        } else {
            invHtml = `<span style="font-family:${fontStack}; background-color:#ffe4e6; color:#be123c; padding:3px 9px; font-weight:700; font-size:9.5pt;">미등록</span>`;
            if (c.oldInvManager) {
                invHtml += `<div style="font-family:${fontStack}; font-size:8.5pt; color:#b45309; font-weight:600; margin-top:4px;">이전담당: ${escapeHtml(c.oldInvManager)}</div>`;
            }
        }

        const oldQ = Number(c.oldQty || 0).toLocaleString();
        const newQ = Number(c.newQty || 0).toLocaleString();
        const diffNum = Number(c.diffQty || 0);
        const unitStr = c.unit ? ` ${escapeHtml(c.unit)}` : '';
        let qtySubHtml = '';
        if (c.changeType === 'REPLACED' && diffNum === 0) {
            qtySubHtml = `<span style="font-family:${fontStack}; color:#6d28d9; font-weight:700; font-size:9pt;">1:1 교체</span>`;
        } else if (diffNum > 0) {
            qtySubHtml = `<span style="font-family:${numFontStack}; color:#059669; font-weight:700; font-size:9pt;">+${diffNum.toLocaleString()}${unitStr}</span>`;
        } else if (diffNum < 0) {
            qtySubHtml = `<span style="font-family:${numFontStack}; color:#dc2626; font-weight:700; font-size:9pt;">${diffNum.toLocaleString()}${unitStr}</span>`;
        }

        const rowBg = isUnreg ? '#fff5f7' : (idx % 2 === 0 ? '#ffffff' : '#f8fafc');

        return `<tr style="background-color:${rowBg};">
            <td style="font-family:${fontStack}; padding:12px 10px; border-bottom:1px solid #e2e8f0; text-align:center; vertical-align:middle;">
                <span style="font-family:${fontStack}; background-color:${badgeBg}; color:${badgeColor}; padding:4px 9px; font-weight:700; font-size:9pt; display:inline-block;">${badgeText}</span>
            </td>
            <td style="font-family:${fontStack}; padding:12px 10px; border-bottom:1px solid #e2e8f0; text-align:center; vertical-align:middle;">
                ${plantBadge}${monthSub}
            </td>
            <td style="font-family:${numFontStack}; padding:12px 10px; border-bottom:1px solid #e2e8f0; vertical-align:middle;">
                ${serialMain}${orderSub}
            </td>
            <td style="font-family:${numFontStack}; padding:12px 10px; border-bottom:1px solid #e2e8f0; vertical-align:middle; color:#1e293b; font-size:9.5pt;">
                <div style="font-weight:600;">${matTitle}</div>${matSub}
            </td>
            <td style="font-family:${numFontStack}; padding:12px 10px; border-bottom:1px solid #e2e8f0; vertical-align:middle;">
                ${partNumHtml}
            </td>
            <td style="font-family:${fontStack}; padding:12px 10px; border-bottom:1px solid #e2e8f0; vertical-align:middle;">
                ${partDescHtml}
            </td>
            <td style="font-family:${fontStack}; padding:12px 10px; border-bottom:1px solid #e2e8f0; text-align:center; vertical-align:middle;">
                ${invHtml}
            </td>
            <td style="font-family:${numFontStack}; padding:12px 12px; border-bottom:1px solid #e2e8f0; text-align:right; vertical-align:middle;">
                <div style="font-size:10pt; color:#334155;"><span style="color:#94a3b8;">${oldQ}</span> ➔ <b style="color:#0f172a;">${newQ}</b><span style="font-size:9pt; color:#64748b;">${unitStr}</span></div>
                <div style="margin-top:2px;">${qtySubHtml}</div>
            </td>
        </tr>`;
    }).join('\r\n');

    return `<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<style>
    body, table, th, td, div, p, span {
        font-family: 'Pretendard', 'Noto Sans KR', 'Noto Sans CJK KR', 'LG Smart UI', 'Segoe UI', 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif !important;
        mso-ascii-font-family: 'Segoe UI', 'Noto Sans KR', sans-serif;
        mso-fareast-font-family: 'Noto Sans KR', 'Noto Sans CJK KR', 'LG Smart UI', 'Malgun Gothic', sans-serif;
        mso-hansi-font-family: 'Segoe UI', 'Noto Sans KR', sans-serif;
        letter-spacing: -0.2px;
    }
</style>
</head>
<body style="margin:0; padding:0; background-color:#f1f5f9; font-family:${fontStack}; color:#0f172a;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#f1f5f9; padding:24px 16px; font-family:${fontStack};">
        <tr>
            <td align="center" style="font-family:${fontStack};">
                <table width="100%" cellpadding="0" cellspacing="0" border="0" style="max-width:1180px; background-color:#ffffff; border:1px solid #e2e8f0; border-top:4px solid #2563eb; font-family:${fontStack};">
                    <!-- Header Section -->
                    <tr>
                        <td style="padding:26px 30px 20px 30px; border-bottom:1px solid #f1f5f9; font-family:${fontStack};">
                            <div style="font-family:${numFontStack}; font-size:9pt; font-weight:700; color:#2563eb; letter-spacing:0.8px; margin-bottom:5px;">SAP ERP ENGINEERING CHANGE &amp; BOM MONITOR</div>
                            <div style="font-family:${fontStack}; font-size:16pt; font-weight:700; color:#0f172a; margin-bottom:8px; letter-spacing:-0.4px;">설계변경(설변) 및 하위 가공품 소요량 변동 리포트</div>
                            <div style="font-family:${fontStack}; font-size:10pt; color:#64748b;">
                                비교 구간: <b style="color:#1e293b;">${escapeHtml(baseTitle)}</b> &nbsp;➔&nbsp; <b style="color:#2563eb;">${escapeHtml(targetTitle)}</b>
                                &nbsp;&nbsp;|&nbsp;&nbsp; 발송 일시: <b style="color:#1e293b;">${nowStr}</b>
                            </div>
                        </td>
                    </tr>

                    <!-- KPI Cards Section -->
                    <tr>
                        <td style="padding:20px 30px 12px 30px; font-family:${fontStack};">
                            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="font-family:${fontStack};">
                                <tr>
                                    <td width="19%" style="font-family:${fontStack}; background-color:#f8fafc; border:1px solid #e2e8f0; border-top:3px solid #7c3aed; padding:14px 16px;">
                                        <div style="font-size:9.5pt; color:#475569; font-weight:700;">품번 변경 (설변 A➔B)</div>
                                        <div style="font-family:${numFontStack}; font-size:17pt; font-weight:700; color:#6d28d9; margin-top:5px;">${summary.totalReplaced || 0}<span style="font-family:${fontStack}; font-size:10pt; font-weight:600; color:#64748b;"> 쌍</span></div>
                                    </td>
                                    <td width="1%">&nbsp;</td>
                                    <td width="19%" style="font-family:${fontStack}; background-color:#f8fafc; border:1px solid #e2e8f0; border-top:3px solid #10b981; padding:14px 16px;">
                                        <div style="font-size:9.5pt; color:#475569; font-weight:700;">순수 신규 투입</div>
                                        <div style="font-family:${numFontStack}; font-size:17pt; font-weight:700; color:#059669; margin-top:5px;">${summary.totalPureNew || 0}<span style="font-family:${fontStack}; font-size:9.5pt; font-weight:600; color:#64748b;"> 건 (전체 ${summary.totalNew || 0})</span></div>
                                    </td>
                                    <td width="1%">&nbsp;</td>
                                    <td width="19%" style="font-family:${fontStack}; background-color:#f8fafc; border:1px solid #e2e8f0; border-top:3px solid #ef4444; padding:14px 16px;">
                                        <div style="font-size:9.5pt; color:#475569; font-weight:700;">순수 삭제/제외</div>
                                        <div style="font-family:${numFontStack}; font-size:17pt; font-weight:700; color:#dc2626; margin-top:5px;">${summary.totalPureDeleted || 0}<span style="font-family:${fontStack}; font-size:9.5pt; font-weight:600; color:#64748b;"> 건 (전체 ${summary.totalDeleted || 0})</span></div>
                                    </td>
                                    <td width="1%">&nbsp;</td>
                                    <td width="19%" style="font-family:${fontStack}; background-color:#f8fafc; border:1px solid #e2e8f0; border-top:3px solid #f59e0b; padding:14px 16px;">
                                        <div style="font-size:9.5pt; color:#475569; font-weight:700;">소요 수량 변동</div>
                                        <div style="font-family:${numFontStack}; font-size:17pt; font-weight:700; color:#d97706; margin-top:5px;">${summary.totalQtyChanged || 0}<span style="font-family:${fontStack}; font-size:10pt; font-weight:600; color:#64748b;"> 건</span></div>
                                    </td>
                                    <td width="1%">&nbsp;</td>
                                    <td width="20%" style="font-family:${fontStack}; background-color:#fff1f2; border:1px solid #fecdd3; border-top:3px solid #e11d48; padding:14px 16px;">
                                        <div style="font-size:9.5pt; color:#be123c; font-weight:700;">변동중 담당자 미등록</div>
                                        <div style="font-family:${numFontStack}; font-size:17pt; font-weight:700; color:#e11d48; margin-top:5px;">${summary.totalDiffUnregistered || 0}<span style="font-family:${fontStack}; font-size:9.5pt; font-weight:600; color:#9f1239;"> 건 (전체 ${summary.uniqueUnregisteredParts || 0}품번)</span></div>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>

                    <!-- Table Section -->
                    <tr>
                        <td style="padding:12px 30px 30px 30px; font-family:${fontStack};">
                            <div style="font-family:${fontStack}; font-size:10pt; color:#64748b; margin-bottom:12px;">
                                총 <b style="color:#0f172a;">${groupedChanges.length}건</b>의 변동 내역이 감지되었습니다. 상세 원본 데이터는 첨부된 Excel 파일을 참고해 주세요.
                            </div>
                            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse; border-top:2px solid #0f172a; font-family:${fontStack};">
                                <thead>
                                    <tr style="background-color:#f8fafc; color:#334155; font-size:9.5pt; text-align:left;">
                                        <th style="font-family:${fontStack}; padding:11px 10px; border-bottom:1px solid #cbd5e1; text-align:center; width:78px;">구분</th>
                                        <th style="font-family:${fontStack}; padding:11px 10px; border-bottom:1px solid #cbd5e1; text-align:center; width:88px;">플랜트/월</th>
                                        <th style="font-family:${fontStack}; padding:11px 10px; border-bottom:1px solid #cbd5e1; width:130px;">호기(Serial)/오더</th>
                                        <th style="font-family:${fontStack}; padding:11px 10px; border-bottom:1px solid #cbd5e1; width:175px;">완제품 기종내역</th>
                                        <th style="font-family:${fontStack}; padding:11px 10px; border-bottom:1px solid #cbd5e1; width:170px;">가공품 품번 (이전 ➔ 변경)</th>
                                        <th style="font-family:${fontStack}; padding:11px 10px; border-bottom:1px solid #cbd5e1;">가공품 품명 (Comp.Desc)</th>
                                        <th style="font-family:${fontStack}; padding:11px 10px; border-bottom:1px solid #cbd5e1; text-align:center; width:115px;">담당자 (Inv.Mgr)</th>
                                        <th style="font-family:${fontStack}; padding:11px 12px; border-bottom:1px solid #cbd5e1; text-align:right; width:110px;">소요량 (이전➔최신)</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    ${rowsHtml}
                                </tbody>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>
</html>`;
}

async function sendSapDiffEmail(options = {}) {
    const cfg = getDiffEmailConfig();
    const isAuto = !!options.isAutoAfterSync;
    if (isAuto && !cfg.enabled) {
        console.log('[SAP Diff Email] Auto-send is disabled in config. Skipping.');
        return { skipped: true, reason: '자동 메일 발송 설정이 꺼져 있습니다.' };
    }

    const toAddr = (options.customTo !== undefined ? options.customTo : cfg.to || '').trim();
    const ccAddr = (options.customCc !== undefined ? options.customCc : cfg.cc || '').trim();
    if (!toAddr) {
        return { skipped: true, reason: '수신자(To) 이메일 주소가 설정되지 않았습니다.' };
    }

    const snapshots = getSnapshotList();
    const targetId = options.targetId || 'current';
    let baseId = options.baseId;
    if (!baseId) {
        const prev = snapshots.find(s => s.id !== targetId);
        baseId = prev ? prev.id : targetId;
    }

    if (baseId === targetId) {
        return { skipped: true, reason: '비교할 이전 스냅샷 데이터가 없습니다.' };
    }

    const baseTitle = (snapshots.find(s => s.id === baseId) || {}).title || baseId;
    const targetTitle = (snapshots.find(s => s.id === targetId) || {}).title || targetId;

    const baseRecords = loadComponentDataset(baseId);
    const targetRecords = loadComponentDataset(targetId);
    const diffResult = calculateComponentDiff(baseRecords, targetRecords);
    const { summary, groupedChanges } = diffResult;

    // 변동 0건일 때는 자동 발송 제외
    if ((summary.totalChanges || 0) === 0 || groupedChanges.length === 0) {
        const msg = `변동 내역 0건으로 발송 생략 (${baseTitle} -> ${targetTitle})`;
        console.log(`[SAP Diff Email] ${msg}`);
        saveDiffEmailConfig({ lastSentStatus: msg });
        return { skipped: true, reason: msg, summary };
    }

    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    const dateStr = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
    const fileDateStr = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}`;

    const subject = `[ERP 설변/가공품 변동 알림] ${dateStr} (품번변경 ${summary.totalReplaced || 0}쌍 · 신규 ${summary.totalPureNew || 0}건 · 담당자미등록 ${summary.totalDiffUnregistered || 0}건)`;
    const htmlBody = buildDiffEmailHtml(summary, groupedChanges, baseTitle, targetTitle);
    const excelBuf = buildDiffExcelBuffer(groupedChanges);

    const tempDir = SAP_CACHE_DIR;
    const excelPath = path.join(tempDir, `ERP_설변_가공품변동_${fileDateStr}.xlsx`);
    const htmlPath = path.join(tempDir, `diff_mail_body_${fileDateStr}.html`);
    const payloadPath = path.join(tempDir, `diff_mail_payload_${fileDateStr}.json`);

    fs.writeFileSync(excelPath, excelBuf);
    fs.writeFileSync(htmlPath, htmlBody, 'utf8');
    fs.writeFileSync(payloadPath, JSON.stringify({
        subject,
        to: toAddr,
        cc: ccAddr,
        htmlBodyPath: htmlPath,
        attachmentPath: excelPath
    }, null, 2), 'utf8');

    const psScript = path.join(__dirname, 'sap_automation', 'send_outlook_mail.ps1');
    try {
        execSync(`powershell -NoProfile -ExecutionPolicy Bypass -File "${psScript}" -PayloadPath "${payloadPath}"`, {
            stdio: 'pipe',
            timeout: 30000
        });
        const sentAt = new Date().toISOString();
        const statusMsg = `발송 성공 (${toAddr}, 총 ${groupedChanges.length}건)`;
        saveDiffEmailConfig({
            to: toAddr,
            cc: ccAddr,
            lastSentAt: sentAt,
            lastSentStatus: statusMsg
        });
        console.log(`[SAP Diff Email] Successfully sent via Outlook to ${toAddr}`);
        return { sent: true, sentAt, to: toAddr, cc: ccAddr, summary, totalRows: groupedChanges.length };
    } finally {
        try { fs.unlinkSync(htmlPath); } catch (e) {}
        try { fs.unlinkSync(payloadPath); } catch (e) {}
        try { fs.unlinkSync(excelPath); } catch (e) {}
    }
}

app.get('/api/sap-diff/email-config', (req, res) => {
    try {
        res.json({ success: true, config: getDiffEmailConfig() });
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

app.post('/api/sap-diff/email-config', requireAdmin, (req, res) => {
    try {
        const { enabled, to, cc } = req.body || {};
        const updated = saveDiffEmailConfig({
            enabled: typeof enabled === 'boolean' ? enabled : true,
            to: typeof to === 'string' ? to.trim() : '',
            cc: typeof cc === 'string' ? cc.trim() : ''
        });
        res.json({ success: true, config: updated });
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

app.post('/api/sap-diff/send-email', requireAdmin, async (req, res) => {
    try {
        const { baseId, targetId, to, cc } = req.body || {};
        const result = await sendSapDiffEmail({
            baseId,
            targetId,
            customTo: to,
            customCc: cc,
            isAutoAfterSync: false
        });
        res.json({ success: true, ...result, config: getDiffEmailConfig() });
    } catch (e) {
        console.error('[SAP Diff Email] Send error:', e.message);
        res.status(500).json({ success: false, error: e.message });
    }
});

// 서버 종료 API (관리자 전용)
app.post('/api/shutdown', requireAdmin, (req, res) => {
    console.log('[api] Shutdown requested by admin. Exiting...');
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
    // Pre-warm SAP cache asynchronously right after server starts
    setTimeout(() => {
        try {
            ['1842', '1840', 'component_1842', 'component_1840'].forEach(t => {
                try { getOrBuildSapCache(t); } catch (e) {}
            });
        } catch (e) {}
    }, 100);
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

