const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");

const SAP_EXPORT_FILE = "C:\\Users\\i0215099\\Documents\\SAP\\SAP GUI\\export.MHTML";
const WORKSPACE_DIR = path.resolve(__dirname, "..");

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

function extractSalesDocsFromMhtml(filePath) {
    if (!fs.existsSync(filePath)) return [];
    try {
        const content = fs.readFileSync(filePath, "utf8");
        const tableMatch = content.match(/<table[^>]*>([\s\S]*?)<\/table>/i);
        if (!tableMatch) return [];
        const rows = tableMatch[1].match(/<tr[^>]*>([\s\S]*?)<\/tr>/gi) || [];
        
        let salesDocIdx = -1;
        for (let i = 0; i < Math.min(10, rows.length); i++) {
            const cells = (rows[i].match(/<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi) || []).map(c => c.replace(/<[^>]+>/g, "").trim().toUpperCase());
            const idx = cells.findIndex(c => c.includes("SALES DOC") || c.includes("SALESDOC") || c.includes("판매문서") || c.includes("S/O ORDER"));
            if (idx !== -1) {
                salesDocIdx = idx;
                break;
            }
        }
        if (salesDocIdx === -1) return [];
        
        const docs = new Set();
        for (let i = 1; i < rows.length; i++) {
            const cells = (rows[i].match(/<td[^>]*>([\s\S]*?)<\/td>/gi) || []).map(c => c.replace(/<[^>]+>/g, "").trim());
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

async function waitForExportFile(maxWaitSec = 90) {
    const startTime = Date.now();
    let lastSize = -1;
    let stableCount = 0;

    while ((Date.now() - startTime) < (maxWaitSec * 1000)) {
        await new Promise(r => setTimeout(r, 1000));
        if (fs.existsSync(SAP_EXPORT_FILE)) {
            try {
                const stat = fs.statSync(SAP_EXPORT_FILE);
                if (stat.size > 1000 && stat.size === lastSize) {
                    stableCount++;
                    if (stableCount >= 2) {
                        return true;
                    }
                } else {
                    lastSize = stat.size;
                    stableCount = 0;
                }
            } catch (e) {}
        }
    }
    return false;
}

function clearSapExportFile() {
    if (fs.existsSync(SAP_EXPORT_FILE)) {
        try { fs.unlinkSync(SAP_EXPORT_FILE); } catch (e) {}
    }
}

function copyExportToWorkspace(targetFilename) {
    const dest = path.join(WORKSPACE_DIR, targetFilename);
    fs.copyFileSync(SAP_EXPORT_FILE, dest);
    const sizeMb = (fs.statSync(dest).size / 1024 / 1024).toFixed(2);
    console.log("[Sync] Saved " + targetFilename + " (" + sizeMb + " MB)");
    return dest;
}

function setClipboardText(textList) {
    const tmpFile = path.join(__dirname, "temp_clipboard.txt");
    fs.writeFileSync(tmpFile, textList.join("\r\n"), "utf8");
    const safePath = tmpFile.replace(/\\/g, "/");
    const psCmd = 'powershell -NoProfile -Command "Set-Clipboard -Value (Get-Content -Raw -Encoding UTF8 \'' + safePath + '\')"';
    execSync(psCmd, { stdio: "ignore" });
}

function getDefaultMonths() {
    const now = new Date();
    const y1 = now.getFullYear();
    const m1 = now.getMonth();
    const dStart = new Date(y1, m1 - 1, 1);
    const dEnd = new Date(y1, m1 + 4, 1);
    
    const fmt = d => d.getFullYear() + "." + String(d.getMonth() + 1).padStart(2, "0");
    return {
        start: fmt(dStart),
        end: fmt(dEnd)
    };
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
        console.log("  [SAP One-Stop Sync] Starting Sync for " + startMonth + " ~ " + endMonth);
        console.log("=======================================================\n");

        // --- STEP 1: ZPPM6680 for 1840 ---
        currentSyncState.currentStep = 1;
        currentSyncState.statusText = "[1/4] 남산(1840) 생산계획 데이터 수집 중 (ZPPM6680)...";
        onProgress(currentSyncState);
        console.log(currentSyncState.statusText);

        clearSapExportFile();
        const vbs1 = path.join(__dirname, "zppm6680.vbs");
        execSync('cscript //Nologo "' + vbs1 + '" 1840 ' + startMonth + ' ' + endMonth);
        
        const ok1 = await waitForExportFile(90);
        if (!ok1) throw new Error("1840 생산계획 MHTML 파일 생성 대기 시간 초과");
        copyExportToWorkspace("sap_1840.mhtml");
        
        const docs1840 = extractSalesDocsFromMhtml(path.join(WORKSPACE_DIR, "sap_1840.mhtml"));
        console.log("[Sync] Extracted " + docs1840.length + " Sales Docs for 1840");
        currentSyncState.results.docs1840Count = docs1840.length;

        // --- STEP 2: ZPPR6470 for 1840 ---
        currentSyncState.currentStep = 2;
        currentSyncState.statusText = "[2/4] 남산(1840) 가공품 소요량 데이터 수집 중 (ZPPR6470)...";
        onProgress(currentSyncState);
        console.log(currentSyncState.statusText);

        if (docs1840.length > 0) {
            setClipboardText(docs1840);
            clearSapExportFile();
            const vbs2 = path.join(__dirname, "zppr6470.vbs");
            execSync('cscript //Nologo "' + vbs2 + '" 1840 e "" 18');
            
            const ok2 = await waitForExportFile(90);
            if (!ok2) throw new Error("1840 가공품 소요량 MHTML 파일 생성 대기 시간 초과");
            copyExportToWorkspace("sap_component_1840.mhtml");
        } else {
            console.warn("[Sync] No Sales Docs found for 1840, skipping ZPPR6470");
        }

        // --- STEP 3: ZPPM6680 for 1842 ---
        currentSyncState.currentStep = 3;
        currentSyncState.statusText = "[3/4] 성주(1842) 생산계획 데이터 수집 중 (ZPPM6680)...";
        onProgress(currentSyncState);
        console.log(currentSyncState.statusText);

        clearSapExportFile();
        execSync('cscript //Nologo "' + vbs1 + '" 1842 ' + startMonth + ' ' + endMonth);
        
        const ok3 = await waitForExportFile(90);
        if (!ok3) throw new Error("1842 생산계획 MHTML 파일 생성 대기 시간 초과");
        copyExportToWorkspace("sap_1842.mhtml");

        const docs1842 = extractSalesDocsFromMhtml(path.join(WORKSPACE_DIR, "sap_1842.mhtml"));
        console.log("[Sync] Extracted " + docs1842.length + " Sales Docs for 1842");
        currentSyncState.results.docs1842Count = docs1842.length;

        // --- STEP 4: ZPPR6470 for 1842 ---
        currentSyncState.currentStep = 4;
        currentSyncState.statusText = "[4/4] 성주(1842) 가공품 소요량 데이터 수집 중 (ZPPR6470)...";
        onProgress(currentSyncState);
        console.log(currentSyncState.statusText);

        if (docs1842.length > 0) {
            setClipboardText(docs1842);
            clearSapExportFile();
            const vbs4 = path.join(__dirname, "zppr6470.vbs");
            execSync('cscript //Nologo "' + vbs4 + '" "" F 44 18');
            
            const ok4 = await waitForExportFile(90);
            if (!ok4) throw new Error("1842 가공품 소요량 MHTML 파일 생성 대기 시간 초과");
            copyExportToWorkspace("sap_component_1842.mhtml");
        } else {
            console.warn("[Sync] No Sales Docs found for 1842, skipping ZPPR6470");
        }

        // Return SAP to home
        try {
            execSync('cscript //Nologo "' + path.join(__dirname, "return_home.vbs") + '"');
        } catch (e) {}

        currentSyncState.statusText = "✅ SAP 최신 데이터 동기화 완료! (4개 파일 갱신됨)";
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
        const tmpFile = path.join(__dirname, "temp_clipboard.txt");
        if (fs.existsSync(tmpFile)) {
            try { fs.unlinkSync(tmpFile); } catch (e) {}
        }
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