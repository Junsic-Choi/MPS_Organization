const XLSX = require('xlsx');

function getMatchKey(s) {
    if (!s) return '';
    // [STABLE ENGINE]: 접두사 및 0 제거, 특수문자 제외하여 결합력 극대화
    let n = s.toString().toUpperCase().trim().split('-')[0];
    n = n.replace(/PUMA|LYNX/g, '').trim();
    n = n.replace(/^P|^L/, '');
    // 로마자 -> 숫자 변환 (II -> 2)
    n = n.replace(/VIII/g, '8').replace(/VII/g, '7').replace(/VI/g, '6').replace(/IV/g, '4').replace(/IX/g, '9');
    n = n.replace(/III/g, '3').replace(/II/g, '2');
    return n.replace(/[^A-Z1-9]/g, '');
}

/**
 * MPS2603-1 (4,650줄), MPS2604-1 (5,106줄) 등 파일에 따라
 * 실제 데이터 수량을 100% 정확하게 자동으로 읽어내던 '진짜' 안정적인 엔진
 */
function processMpsFile(input, rules = {}) {
    const siteMaster = rules.siteMaster || {};
    const realSiteLookup = rules.realSiteLookup || {};

    let wb;
    if (Buffer.isBuffer(input)) {
        wb = XLSX.read(input, { type: 'buffer' });
    } else {
        wb = XLSX.readFile(input);
    }
    const sheetNames = wb.SheetNames;
    
    // 1. [MPS] 탭 분석 (기준 정보 구축)
    const masterWs = wb.Sheets[sheetNames.find(n => n.toUpperCase() === 'MPS') || 'MPS'];
    const codeMap = {};
    if (masterWs) {
        const masterRaw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });
        masterRaw.forEach((row, idx) => {
            if (idx < 5) return;
            const mCode = (row[3] || '').toString().trim(); // D열
            const pName = (row[4] || '').toString().trim(); // E열
            if (pName) {
                const modelPart = pName.split('-')[0].trim();
                const key = getMatchKey(modelPart);
                if (!codeMap[key]) {
                    codeMap[key] = { code: mCode, product: pName };
                }
            }
        });
    }

    // 2. [생산배포용] 탭 분석
    const mpsWsName = sheetNames.find(n => n.includes('배포')) || '생산배포용';
    const mpsWs = wb.Sheets[mpsWsName];
    const mpsRaw = XLSX.utils.sheet_to_json(mpsWs, { header: 1 });

    // [중요]: 월 헤더 자동 탐지
    const monthInfo = [];
    let headerRowIdx = -1;
    for(let r=0; r<Math.min(15, mpsRaw.length); r++) {
        const row = mpsRaw[r] || [];
        const monthsInRow = row.filter(cell => /^\d+월$/.test((cell || '').toString().trim()));
        if (monthsInRow.length >= 3) {
            headerRowIdx = r;
            row.forEach((cell, idx) => {
                const cellStr = (cell || '').toString().trim();
                if (/^\d+월$/.test(cellStr)) {
                    monthInfo.push({ name: cellStr, col: idx });
                }
            });
            break;
        }
    }

    if (monthInfo.length === 0) {
        [4, 7, 8, 9, 10, 12].forEach((col, i) => {
            monthInfo.push({ name: (i+2) + '월', col: col });
        });
        headerRowIdx = 5;
    }

    const finalResults = [];
    const unusedData = [];
    let runningSite = '', runningGroup = '', runningModel = '';

    for (let r = headerRowIdx + 1; r < mpsRaw.length; r++) {
        const row = mpsRaw[r] || [];
        const site = (row[0] || '').toString().trim();
        const group = (row[1] || '').toString().trim();
        const model = (row[2] || '').toString().trim();
        const rpm = (row[3] || '').toString().trim();

        if (site) runningSite = site;
        if (group) runningGroup = group;
        if (model) runningModel = model;

        if (site === '총합계' || (site === '' && runningSite === '총합계')) continue;
        
        // 모델명이 아예 없는 경우 (공백 등) - 전수검사 대상
        const isModelMissing = !runningModel || runningModel === 'Model';

        const key = getMatchKey(runningModel);
        const mapped = codeMap[key];

        // 사이트 매핑 적용
        let finalSite = runningSite;
        if (siteMaster[runningSite]) {
            finalSite = siteMaster[runningSite];
        } else if (runningSite.match(/^\d+$/) || runningSite.startsWith('184')) {
            // 원본 코드로 남아있는 경우 사이트 마스터에서 찾아봄
            const found = Object.entries(siteMaster).find(([k, v]) => k.includes(runningSite) || runningSite.includes(k));
            if (found) finalSite = found[1];
        }
        finalSite = finalSite.replace(/^\d+\.\s*/, '').trim();

        monthInfo.forEach(mObj => {
            const qty = parseInt(row[mObj.col]) || 0;
            if (qty > 0) {
                for (let i = 0; i < Math.min(qty, 1000); i++) {
                    // 매칭 실패 (코드 없음) 또는 모델명 없음
                    if (!mapped || !mapped.code || isModelMissing) {
                        unusedData.push({
                            Site: finalSite,
                            Group: runningGroup,
                            ModelCode: mapped ? mapped.code : '',
                            ProductName: runningModel || '(기종명 없음)',
                            Month: mObj.name,
                            Category: isModelMissing ? '기종명누락' : '코드미매칭'
                        });
                    } else {
                        finalResults.push({
                            Site: finalSite,
                            Group: runningGroup,
                            Model: runningModel,
                            RPM: rpm,
                            Month: mObj.name,
                            Code: mapped.code,
                            Product: mapped.product
                        });
                    }
                }
            }
        });
    }

    return { finalResults, unusedData, masterPlan: {}, masterModelsByGroup: {} };
}


module.exports = { processMpsFile };
