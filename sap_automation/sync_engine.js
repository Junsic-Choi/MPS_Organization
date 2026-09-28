const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");

const SAP_GUI_DIR = "C:\\Users\\i0215099\\Documents\\SAP\\SAP GUI";
const WORKSPACE_DIR = path.resolve(__dirname, "..");
const USER_PROFILE = process.env.USERPROFILE || "C:\\Users\\i0215099";

// SAP XXL export can save to various default directories based on user history
const WATCH_DIRS = [
    SAP_GUI_DIR,
    path.join(USER_PROFILE, "Desktop", "원본데이터"),
    path.join(USER_PROFILE, "Desktop"),
    path.join(USER_PROFILE, "Downloads"),
    process.env.TEMP || path.join(USER_PROFILE, "AppData", "Local", "Temp")
];

let currentSyncState = {
    running: false,
    currentStep: 0,
    totalSteps: 4,
    statusText: "",
    error: null,
    startedAt: null,
    finishedAt: null,
    results: {}
};

function getSyncStatus() {
    return { ...currentSyncState };
}

function cleanSapExportDir() {
    WATCH_DIRS.forEach(dir => {
        if (!fs.existsSync(dir)) return;
        try {
            const files = fs.readdirSync(dir);
            files.forEach(f => {
                if (/^export.*\.mhtml$/i.test(f)) {
                    try {
                        fs.unlinkSync(path.join(dir, f));
                    } catch (e) {}
                }
            });
        } catch (e) {}
    });
}

function closeSapExcel(specificPath = null) {
    // 1. First try to cleanly close only SAP export workbooks without killing user's own work
    try {
        const script = path.join(__dirname, "close_sap_excel.vbs");
        if (fs.existsSync(script)) {
            execSync(`cscript //Nologo "${script}"`, { stdio: "ignore" });
        }
    } catch (e) {}

    // 2. Release locks on candidate export files
    const targets = specificPath ? [specificPath] : WATCH_DIRS.map(d => path.join(d, "export.MHTML"));
    for (const p of targets) {
        if (fs.existsSync(p)) {
            try {
                fs.unlinkSync(p);
            } catch (e) {
                try {
                    execSync("taskkill /f /im excel.exe", { stdio: "ignore" });
                } catch (kErr) {}
                try { fs.unlinkSync(p); } catch (uErr) {}
            }
        }
    }
}

function extractSalesDocsFromMhtml(filePath) {
    if (!fs.existsSync(filePath)) return [];
    try {
        const content = fs.readFileSync(filePath, "utf8");
        const tableMatch = content.match(/<table[^>]*>([\s\S]*?)<\/table>/i);
        if (!tableMatch) return [];
        const rows = tableMatch[1].match(/<tr[^>]*>([\s\S]*?)<\/tr>/gi) || [];
        if (rows.length < 2) return [];

        let salesDocIdx = -1;
        for (let i = 0; i < Math.min(10, rows.length); i++) {
            const cells = (rows[i].match(/<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi) || [])
                .map(c => c.replace(/<[^>]+>/g, "").trim().toUpperCase());
            const idx = cells.findIndex(c => c.includes("SALES DOC") || c.includes("SALESDOC") || c.includes("판매문서") || c.includes("S/O ORDER"));
            if (idx !== -1) {
                salesDocIdx = idx;
                break;
            }
        }
        if (salesDocIdx === -1) {
            salesDocIdx = 11;
        }

        const docs = new Set();
        for (let i = 1; i < rows.length; i++) {
            const cells = (rows[i].match(/<td[^>]*>([\s\S]*?)<\/td>/gi) || [])
                .map(c => c.replace(/<[^>]+>/g, "").trim());
            const doc = cells[salesDocIdx];
            if (doc && /^\d{5,12}$/.test(doc)) {
                docs.add(doc);
            }
        }
        return Array.from(docs);
    } catch (e) {
        console.error("[Engine] extractSalesDocs error:", e.message);
        return [];
    }
}

async function waitForNewExportFile(sinceTimestamp, maxWaitSec = 180) {
    const startTime = Date.now();
    let lastPath = null;
    let lastSize = -1;
    let stableCount = 0;

    while ((Date.now() - startTime) < (maxWaitSec * 1000)) {
        await new Promise(r => setTimeout(r, 1000));

        let candidateFiles = [];
        for (const dir of WATCH_DIRS) {
            if (!fs.existsSync(dir)) continue;
            try {
                const files = fs.readdirSync(dir)
                    .filter(f => /\.mhtml$/i.test(f))
                    .map(f => {
                        const full = path.join(dir, f);
                        const stat = fs.statSync(full);
                        return { full, size: stat.size, mtime: stat.mtimeMs };
                    })
                    .filter(f => f.mtime >= sinceTimestamp - 3000 && f.size > 1000);
                candidateFiles.push(...files);
            } catch (e) {}
        }

        if (candidateFiles.length > 0) {
            candidateFiles.sort((a, b) => b.mtime - a.mtime);
            const newest = candidateFiles[0];
            if (newest.full === lastPath && newest.size === lastSize) {
                stableCount++;
                if (stableCount >= 2) {
                    console.log(`[Engine] Detected export file: ${newest.full} (${(newest.size / 1024 / 1024).toFixed(2)} MB)`);
                    return newest.full;
                }
            } else {
                lastPath = newest.full;
                lastSize = newest.size;
                stableCount = 0;
            }
        }
    }
    return null;
}

function archiveCurrentSapSnapshot() {
    try {
        const files = ['sap_1840.mhtml', 'sap_1842.mhtml', 'sap_component_1840.mhtml', 'sap_component_1842.mhtml'];
        const existing = files.filter(f => fs.existsSync(path.join(WORKSPACE_DIR, f)));
        if (existing.length === 0) return null;

        const historyDir = path.join(WORKSPACE_DIR, 'sap_history');
        if (!fs.existsSync(historyDir)) fs.mkdirSync(historyDir, { recursive: true });

        const now = new Date();
        const pad = n => String(n).padStart(2, '0');
        const tsId = `${now.getFullYear()}${pad(now.getMonth()+1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
        const snapshotDir = path.join(historyDir, tsId);
        fs.mkdirSync(snapshotDir, { recursive: true });

        const stats = {};
        existing.forEach(f => {
            const src = path.join(WORKSPACE_DIR, f);
            const dst = path.join(snapshotDir, f);
            fs.copyFileSync(src, dst);
            const st = fs.statSync(src);
            stats[f] = { size: st.size, mtime: st.mtime };
        });

        const meta = {
            id: tsId,
            title: `${now.getFullYear()}-${pad(now.getMonth()+1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}`,
            timestamp: now.getTime(),
            createdAt: now.toISOString(),
            files: existing,
            fileStats: stats
        };
        fs.writeFileSync(path.join(snapshotDir, 'metadata.json'), JSON.stringify(meta, null, 2), 'utf8');
        console.log(`[Snapshot] Archived current SAP files to sap_history/${tsId}`);
        return tsId;
    } catch (e) {
        console.error('[Snapshot] Archive error:', e.message);
        return null;
    }
}

function mergeComponentMhtml(baseFilePath, appendFilePath) {
    if (!fs.existsSync(appendFilePath)) return;
    if (!fs.existsSync(baseFilePath)) {
        fs.copyFileSync(appendFilePath, baseFilePath);
        return;
    }
    try {
        const baseContent = fs.readFileSync(baseFilePath, 'utf8');
        const appendContent = fs.readFileSync(appendFilePath, 'utf8');

        const baseTrs = baseContent.match(/<tr[^>]*>[\s\S]*?<\/tr>/gi) || [];
        const appendTrs = appendContent.match(/<tr[^>]*>[\s\S]*?<\/tr>/gi) || [];

        if (appendTrs.length <= 1) return;

        // Build existing row keys from base to prevent duplicate counting
        const existingKeys = new Set();
        for (let i = 1; i < baseTrs.length; i++) {
            const cells = (baseTrs[i].match(/<t[dh][^>]*>[\s\S]*?<\/td>/gi) || [])
                .map(c => c.replace(/<[^>]+>/g, '').trim());
            const key = [cells[1] || '', cells[2] || '', cells[3] || '', cells[4] || '', cells[7] || '', cells[12] || ''].join('|');
            if (key) existingKeys.add(key);
        }

        const rowsToAdd = [];
        for (let i = 1; i < appendTrs.length; i++) {
            const cells = (appendTrs[i].match(/<t[dh][^>]*>[\s\S]*?<\/td>/gi) || [])
                .map(c => c.replace(/<[^>]+>/g, '').trim());
            const key = [cells[1] || '', cells[2] || '', cells[3] || '', cells[4] || '', cells[7] || '', cells[12] || ''].join('|');
            if (!existingKeys.has(key)) {
                existingKeys.add(key);
                rowsToAdd.push(appendTrs[i]);
            }
        }

        if (rowsToAdd.length === 0) {
            console.log(`[Sync] Items are already present in base file (0 new unique rows).`);
            return 0;
        }

        const tableEndIdx = baseContent.lastIndexOf('</table>');
        if (tableEndIdx !== -1) {
            const merged = baseContent.substring(0, tableEndIdx) +
                '\r\n' + rowsToAdd.join('\r\n') + '\r\n' +
                baseContent.substring(tableEndIdx);
            fs.writeFileSync(baseFilePath, merged, 'utf8');
            console.log(`[Sync] Successfully merged ${rowsToAdd.length} unique component rows into ${path.basename(baseFilePath)}`);
            return rowsToAdd.length;
        }
        return 0;
    } catch (err) {
        console.error(`[Sync] mergeComponentMhtml error:`, err.message);
        return 0;
    }
}

function copyExportToWorkspace(sourcePath, targetFilename) {
    const dest = path.join(WORKSPACE_DIR, targetFilename);
    let copied = false;
    for (let retry = 0; retry < 5; retry++) {
        try {
            fs.copyFileSync(sourcePath, dest);
            copied = true;
            break;
        } catch (e) {
            closeSapExcel(sourcePath);
            const waitEnd = Date.now() + 500;
            while (Date.now() < waitEnd) {}
        }
    }
    if (!copied) {
        fs.copyFileSync(sourcePath, dest);
    }

    const sizeMb = (fs.statSync(dest).size / 1024 / 1024).toFixed(2);
    console.log(`[Sync] Saved ${targetFilename} (${sizeMb} MB) from ${sourcePath}`);
    
    // Close Excel view and remove exported file so the next step has a clean path
    closeSapExcel(sourcePath);
    try { fs.unlinkSync(sourcePath); } catch (e) {}
    return dest;
}

function setClipboardText(textList) {
    if (!textList || textList.length === 0) return;
    const data = textList.join("\r\n");
    execSync("clip", { input: data });
}

function getDefaultMonths() {
    const now = new Date();
    const y = now.getFullYear();
    const m = now.getMonth();
    const dStart = new Date(y, m - 1, 1);
    const dEnd = new Date(y, m + 4, 1);
    
    const fmt = d => d.getFullYear() + "." + String(d.getMonth() + 1).padStart(2, "0");
    return {
        start: fmt(dStart),
        end: fmt(dEnd)
    };
}

function runVbs(vbsFile, args = []) {
    const vbsPath = path.join(__dirname, vbsFile);
    const quotedArgs = args.map(a => `"${a}"`).join(" ");
    const cmd = `cscript //Nologo "${vbsPath}" ${quotedArgs}`;
    try {
        const out = execSync(cmd, { encoding: "utf8" });
        return out;
    } catch (err) {
        const out = ((err.stdout || "") + " " + (err.stderr || "")).trim();
        if (out.includes("ERROR_NO_SAPGUI")) {
            throw new Error("SAP Logon이 실행되어 있지 않습니다. SAP Logon을 먼저 실행하고 로그인해 주세요.");
        }
        if (out.includes("ERROR_NO_SCRIPTING")) {
            throw new Error("SAP GUI 스크립팅이 비활성화되어 있습니다. SAP 옵션에서 스크립팅 설정을 확인해 주세요.");
        }
        if (out.includes("ERROR_NO_CONNECTION")) {
            throw new Error("로그인된 SAP 연결이 없습니다. SAP 시스템에 로그인해 주세요.");
        }
        if (out.includes("ERROR_NO_SESSION")) {
            throw new Error("활성화된 SAP 세션 창(화면)이 없습니다.");
        }
        if (out.includes("ERROR_ZERO_ROWS")) {
            throw new Error(`SAP ERP 조회 결과가 0건입니다. (${out.trim()}) SAP의 조회 조건(기간/버전/플랜트)을 확인해 주세요.`);
        }
        throw new Error(out || err.message);
    }
}

async function runSapSync(options = {}) {
    if (currentSyncState.running) {
        throw new Error("이미 SAP 동기화 작업이 진행 중입니다.");
    }

    const defaults = getDefaultMonths();
    const startMonth = options.startMonth || defaults.start;
    const endMonth = options.endMonth || defaults.end;
    const onProgress = options.onProgress || (() => {});

    currentSyncState = {
        running: true,
        currentStep: 0,
        totalSteps: 4,
        statusText: "동기화 시작...",
        error: null,
        startedAt: new Date().toISOString(),
        finishedAt: null,
        results: { startMonth, endMonth }
    };

    try {
        console.log("\n=======================================================");
        console.log(`  [SAP One-Stop Sync] Starting Sync for ${startMonth} ~ ${endMonth}`);
        console.log("=======================================================\n");

        // 0. Auto Archive previous sync files to sap_history
        archiveCurrentSapSnapshot();

        // Initial cleanup: Close prior SAP export windows and return SAP to home
        closeSapExcel();
        cleanSapExportDir();
        try { runVbs("return_home.vbs"); } catch (e) {}

        // --- STEP 1: ZPPM6680 for 1840 (남산+) ---
        currentSyncState.currentStep = 1;
        currentSyncState.statusText = `[1/4] 남산+(1840) 생산계획 데이터 수집 중 (${startMonth} ~ ${endMonth}, ZPPM6680)...`;
        onProgress(currentSyncState);
        console.log(currentSyncState.statusText);

        let stepStart = Date.now();
        runVbs("zppm6680.vbs", ["1840", startMonth, endMonth]);
        
        const file1 = await waitForNewExportFile(stepStart, 180);
        if (!file1) throw new Error("남산+(1840) 생산계획 MHTML 파일 생성 대기시간 초과 (180초)");
        copyExportToWorkspace(file1, "sap_1840.mhtml");
        
        const docs1840 = extractSalesDocsFromMhtml(path.join(WORKSPACE_DIR, "sap_1840.mhtml"));
        console.log(`[Sync] Extracted ${docs1840.length} Sales Docs for 남산+(1840)`);
        currentSyncState.results.docs1840Count = docs1840.length;

        // Reset SAP to home before step 2
        try { runVbs("return_home.vbs"); } catch (e) {}

        // --- STEP 2: ZPPR6470 for 1840 (남산+) ---
        currentSyncState.currentStep = 2;
        currentSyncState.statusText = `[2/4] 남산+(1840) 사내 가공품(e) 소요량 수집 중 (${docs1840.length}개 오더, ZPPR6470)...`;
        onProgress(currentSyncState);
        console.log(currentSyncState.statusText);

        if (docs1840.length > 0) {
            // 2-1: 사내 가공품 (조달구분 = e)
            setClipboardText(docs1840);
            stepStart = Date.now();
            runVbs("zppr6470.vbs", ["1840", "e", "", "18", ""]);
            
            const file2 = await waitForNewExportFile(stepStart, 240);
            if (!file2) throw new Error("남산+(1840) 가공품 소요량 MHTML 파일 생성 대기시간 초과 (240초)");
            copyExportToWorkspace(file2, "sap_component_1840.mhtml");

            // Reset SAP to home before step 2-2
            try { runVbs("return_home.vbs"); } catch (e) {}

            // 2-2: Warehouse Controller PS, PZ (A급 유니트) 품목 분할 수집 (조달구분='', wareCtrl='PS,PZ')
            // 전체 5,000+개 오더를 한 번에 조회 시 SAP 메모리 부족(TSV_TNEW_PAGE_ALLOC_FAILED)이 발생하므로 1,000개 단위 분할 조회
            const BATCH_SIZE = 1000;
            const batches = [];
            for (let i = 0; i < docs1840.length; i += BATCH_SIZE) {
                batches.push(docs1840.slice(i, i + BATCH_SIZE));
            }
            console.log(`[Sync] Splitting ${docs1840.length} orders into ${batches.length} batches (size: ${BATCH_SIZE}) for A-class unit query...`);

            let totalUnitAMerged = 0;
            const targetCompFile = path.join(WORKSPACE_DIR, "sap_component_1840.mhtml");

            for (let bIdx = 0; bIdx < batches.length; bIdx++) {
                const batch = batches[bIdx];
                currentSyncState.statusText = `[2/4] 남산+(1840) A급 유니트(PS, PZ) 분할 수집 중 [${bIdx + 1}/${batches.length}] (${batch.length}개 오더)...`;
                onProgress(currentSyncState);
                console.log(currentSyncState.statusText);

                try { runVbs("return_home.vbs"); } catch (e) {}

                setClipboardText(batch);
                stepStart = Date.now();
                try {
                    runVbs("zppr6470.vbs", ["1840", "", "", "18", "PS,PZ"]);
                    const fileUnitA = await waitForNewExportFile(stepStart, 180);
                    if (fileUnitA) {
                        const tempFile = path.join(WORKSPACE_DIR, `sap_component_1840_batch_${bIdx}.mhtml`);
                        fs.copyFileSync(fileUnitA, tempFile);
                        closeSapExcel();
                        try { fs.unlinkSync(fileUnitA); } catch (e) {}

                        const mergedCount = mergeComponentMhtml(targetCompFile, tempFile);
                        try { fs.unlinkSync(tempFile); } catch (e) {}
                        totalUnitAMerged += (mergedCount || 0);
                        console.log(`[Sync] Batch ${bIdx + 1}/${batches.length}: merged ${mergedCount || 0} unique unit rows.`);
                    }
                } catch (bErr) {
                    console.warn(`[Sync] Batch ${bIdx + 1}/${batches.length} skipped or no data:`, bErr.message);
                }
            }
            currentSyncState.results.pypzMergedRows = totalUnitAMerged;
            console.log(`[Sync] Total PS/PZ (A급 유니트) items successfully merged: ${totalUnitAMerged} rows`);
        } else {
            console.warn("[Sync] No Sales Docs found for 1840, skipping ZPPR6470");
        }

        // Reset SAP to home before step 3
        try { runVbs("return_home.vbs"); } catch (e) {}

        // --- STEP 3: ZPPM6680 for 1842 (성주) ---
        currentSyncState.currentStep = 3;
        currentSyncState.statusText = `[3/4] 성주(1842) 생산계획 데이터 수집 중 (${startMonth} ~ ${endMonth}, ZPPM6680)...`;
        onProgress(currentSyncState);
        console.log(currentSyncState.statusText);

        stepStart = Date.now();
        runVbs("zppm6680.vbs", ["1842", startMonth, endMonth]);
        
        const file3 = await waitForNewExportFile(stepStart, 180);
        if (!file3) throw new Error("성주(1842) 생산계획 MHTML 파일 생성 대기시간 초과 (180초)");
        copyExportToWorkspace(file3, "sap_1842.mhtml");

        const docs1842 = extractSalesDocsFromMhtml(path.join(WORKSPACE_DIR, "sap_1842.mhtml"));
        console.log(`[Sync] Extracted ${docs1842.length} Sales Docs for 성주(1842)`);
        currentSyncState.results.docs1842Count = docs1842.length;

        // Reset SAP to home before step 4
        try { runVbs("return_home.vbs"); } catch (e) {}

        // --- STEP 4: ZPPR6470 for 1842 (성주) ---
        currentSyncState.currentStep = 4;
        currentSyncState.statusText = `[4/4] 성주(1842) 가공품 소요량 데이터 수집 중 (${docs1842.length}개 오더, ZPPR6470)...`;
        onProgress(currentSyncState);
        console.log(currentSyncState.statusText);

        if (docs1842.length > 0) {
            setClipboardText(docs1842);
            stepStart = Date.now();
            runVbs("zppr6470.vbs", ["1842", "F", "44", "18"]);
            
            const file4 = await waitForNewExportFile(stepStart, 240);
            if (!file4) throw new Error("성주(1842) 가공품 소요량 MHTML 파일 생성 대기시간 초과 (240초)");
            copyExportToWorkspace(file4, "sap_component_1842.mhtml");
        } else {
            console.warn("[Sync] No Sales Docs found for 1842, skipping ZPPR6470");
        }

        // Return SAP to home screen
        try { runVbs("return_home.vbs"); } catch (e) {}

        currentSyncState.statusText = "✅ SAP 최신 데이터 동기화 완료! (남산+/성주 4개 파일 갱신됨)";
        currentSyncState.finishedAt = new Date().toISOString();
        onProgress(currentSyncState);
        console.log("\n=======================================================");
        console.log("  [SAP One-Stop Sync] ALL 4 STEPS COMPLETED SUCCESSFULLY!");
        console.log("=======================================================\n");

    } catch (err) {
        console.error("[Sync Error]", err.message);
        currentSyncState.error = err.message;
        currentSyncState.statusText = "❌ 동기화 실패: " + err.message;
        onProgress(currentSyncState);
        throw err;
    } finally {
        currentSyncState.running = false;
    }

    return currentSyncState;
}

if (require.main === module) {
    const args = process.argv.slice(2);
    let startArg = null, endArg = null;
    for (let i = 0; i < args.length; i++) {
        if (args[i] === "--start" && args[i + 1]) startArg = args[i + 1];
        if (args[i] === "--end" && args[i + 1]) endArg = args[i + 1];
    }
    runSapSync({ startMonth: startArg, endMonth: endArg })
        .then(() => process.exit(0))
        .catch(err => {
            console.error("Fatal Sync Failure:", err.message);
            process.exit(1);
        });
}

module.exports = { runSapSync, getSyncStatus };