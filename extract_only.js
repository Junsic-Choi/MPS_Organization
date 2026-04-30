const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

if (global.gc) { global.gc(); }

const TARGET_FILENAME = process.argv[2] || 'MPS2603-1.xlsx';
const FILE_PATH = path.join(__dirname, TARGET_FILENAME);

// [FIX] 410 -> 4100 등 숫자 규격 불일치 해결을 위한 정규화
function getMatchKey(s) {
    if (!s) return "";
    let n = s.toString().toUpperCase().trim().split('-')[0].split(' ')[0];
    n = n.replace(/PUMA|LYNX/g, '');
    n = n.replace(/^P|^L/, '');
    // 0 제거를 통해 410과 4100을 동일하게 만듬 (주의: 위험할 수 있으나 이 데이터셋에선 유효)
    return n.replace(/0/g, '').replace(/[^A-Z0-9]/g, '');
}

// [FIX] 보조 사이트 매핑 (MPS 시트 코드용)
const mpsSiteMap = { 
    "1840":"01. 남산", "1841":"01. 남산", "1":"01. 남산", "10":"01. 남산(보조)",
    "1848":"02. 성주", "2":"02. 성주", 
    "1842":"03. 창원", "3":"03. 창원",
    "5":"05. 세양", "6":"06. 원진"
};
const mpsGroupMap = { "10": "PUMA Series", "20": "LYNX Series", "30": "Horizontal Series", "50": "V-Series/DVF" };

try {
    console.log(`[STEP] 엑셀 파일 로드 중: ${TARGET_FILENAME}`);
    const wb = XLSX.readFile(FILE_PATH, { type: 'file', cellStyles: false, cellNF: false }); 
    const sheetNames = wb.SheetNames;
    
    // 1. 마스터 정보 구축 (Sheet 0) - 메모리 절약을 위해 먼저 처리 후 삭제
    console.log(`[STEP] 시트 0(배포용) 분석 및 마스터 매핑 테이블 생성 중...`);
    const 배포ws = wb.Sheets[sheetNames[0]]; 
    let 배포Raw = XLSX.utils.sheet_to_json(배포ws, { header: 1 });
    delete wb.Sheets[sheetNames[0]]; // 사용 완료 시트 즉시 제거
    if (global.gc) { global.gc(); }

    const masterLookup = {};
    let lastSite = "", lastGroup = "";
    
    배포Raw.forEach((row, idx) => {
        if (idx < 6) return;
        let s = (row[0] || '').toString().trim();
        let g = (row[1] || '').toString().trim();
        let m = (row[2] || '').toString().trim();
        let r = (row[3] || '').toString().trim();
        
        if (s) lastSite = s;
        if (g) lastGroup = g;
        
        if (m) {
            const key = getMatchKey(m);
            if (!masterLookup[key] || (r && r !== '0')) {
                masterLookup[key] = { site: lastSite, group: lastGroup, model: m, rpm: r };
            }
        }
    });
    console.log(`[STEP] 마스터 데이터 ${Object.keys(masterLookup).length}건 매핑 완료.`);
    배포Raw = null; 
    if (global.gc) { global.gc(); }

    // 2. MPS 시트 처리
    console.log(`[STEP] MPS 시트 데이터 로드 및 공급 POOL 구성 중...`);
    let mpsWs = wb.Sheets['MPS'] || wb.Sheets['mps'] || wb.Sheets['Sheet2'];
    
    if (!mpsWs) {
        console.error(`[ERROR] 'MPS' 시트를 찾을 수 없습니다. 가용 시트: ${sheetNames.join(', ')}`);
        process.exit(1);
    }

    let mpsRaw = XLSX.utils.sheet_to_json(mpsWs, { header: 1 });
    if (!mpsRaw || mpsRaw.length === 0) {
        console.error(`[ERROR] MPS 시트에 데이터가 없습니다.`);
        process.exit(1);
    }
    delete wb.Sheets[mpsWs];
    if (global.gc) { global.gc(); }

    const monthNames = ["2월", "3월", "4월", "5월", "6월", "7월"];
    const mpsMonthIdxs = [8, 12, 17, 22, 28, 34]; 
    
    const mpsPool = {};
    monthNames.forEach(m => { mpsPool[m] = {}; });

    // 2. 공급 POOL (Sheet 1)
    for (let r = 5; r < mpsRaw.length; r++) {
        const row = mpsRaw[r] || [];
        const code = (row[3] || '').toString().trim();
        const prod = (row[4] || '').toString().trim();
        if (!code || !prod || code.includes('계') || prod.includes('합계')) continue;
        
        const key = getMatchKey(prod);
        
        mpsMonthIdxs.forEach((colIdx, i) => {
            const m = monthNames[i];
            const q = parseInt(row[colIdx]) || 0;
            if (q > 0) {
                if (!mpsPool[m][key]) mpsPool[m][key] = [];
                for (let k = 0; k < q; k++) {
                    // 객체 대신 콤팩트한 문자열로 저장하여 메모리 절약
                    mpsPool[m][key].push(`${prod}|${code}`);
                }
            }
        });
    }
    // mpsRaw는 아래 DEMAND 매핑 루프에서도 재사용되므로 null 하지 않음
    // (GC는 DEMAND 루프 이후에 수행)

    // 3. 요구 DEMAND 및 하이브리드 미러링
    console.log(`[STEP] 시트 간 하이브리드 결합 및 월별 유닛 매칭 시작...`);
    const finalResults = [];

    monthNames.forEach((month, mIdx) => {
        const colIdx = mpsMonthIdxs[mIdx];
        let runningSite = "", runningGroup = "";

        for (let r = 5; r < mpsRaw.length; r++) {
            const row = mpsRaw[r] || [];
            
            // MPS 시트 원본 정보 실시간 추적
            if (row[6]) runningSite = mpsSiteMap[row[6].toString().trim()] || row[6].toString(); 
            if (row[2]) runningGroup = mpsGroupMap[row[2].toString().trim()] || row[2].toString();

            const mCode = (row[3] || '').toString().trim();
            const prodNameRaw = (row[4] || '').toString().trim();
            if (!mCode || !prodNameRaw || mCode.includes('계') || mCode.includes('합계')) continue;

            const q = parseInt(row[colIdx]) || 0;
            if (q > 0) {
                const key = getMatchKey(prodNameRaw);
                // Priority 1: Sheet 0 (Master), Priority 2: MPS Sheet (Running)
                const info = masterLookup[key];
                
                const finalSite = info ? info.site : runningSite;
                const finalGroup = info ? info.group : runningGroup;
                const finalModel = info ? info.model : prodNameRaw.split('-')[0];
                const finalRPM = info ? info.rpm : "";

                for (let k = 0; k < q; k++) {
                    finalResults.push({ 
                        site: finalSite, group: finalGroup, model: finalModel, 
                        rpm: finalRPM, month, code: mCode, productRaw: prodNameRaw, key: key, match: null 
                    });
                }
            }
        }

        // 월별 유닛 매칭
        const currentMonthData = finalResults.filter(f => f.month === month);
        currentMonthData.forEach(need => {
            const pool = mpsPool[month][need.key];
            if (pool && pool.length > 0) {
                const pooledItem = pool.shift(); 
                const parts = pooledItem.split('|');
                need.match = { product: parts[0], code: parts[1] };
            }
        });
    });

    mpsRaw = null; // 여기서 해제 (DEMAND 루프 완료 후)
    if (global.gc) { global.gc(); }

    console.log(`[STEP] 결과 데이터 CSV 변환 및 저장 중...`);
    const outputRows = [['Site', 'Group', 'Model', 'RPM', 'Month', 'Code', 'Product']];
    finalResults.forEach(r => {
        outputRows.push([r.site, r.group, r.model, r.rpm, r.month, r.code, r.match ? r.match.product : "UNMAPPED"]);
    });
    
    fs.writeFileSync('_MPS_Final_Data_v3.csv', "\ufeff" + outputRows.map(r => r.map(v => `"${(v || '').toString().replace(/"/g, '""')}"`).join(',')).join('\n'), 'utf8');
    
    console.log(`[INFO] Final Done (Rev 10). Count=${finalResults.length}`);
    fs.writeFileSync('server_startup.log', `OK:${finalResults.length}`);
    process.exit(0);

} catch (err) {
    const errMsg = `[FATAL ERROR]: ${err.message}\nStack: ${err.stack}`;
    console.error(errMsg);
    fs.writeFileSync('extract_error.log', errMsg, 'utf8');
}
