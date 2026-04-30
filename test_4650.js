const XLSX = require('xlsx');

function getMatchKey(prod) {
    if (!prod) return '';
    return prod.toString().trim().toUpperCase().replace(/\s+/g, '').replace(/[^A-Z0-9-]/g, '');
}

const wb = XLSX.readFile('c:/Users/i0215099/Desktop/MPS_UPDATE/MPS2603-1.xlsx');
const ws = wb.Sheets['생산배포용'];
const raw = XLSX.utils.sheet_to_json(ws, {header:1});

// MPS2603-1 전용 인덱스 테스트
const monthIndices = [4, 7, 8, 9, 10, 12]; // 2월~7월 추정
let totalCount = 0;

for (let r = 6; r < raw.length; r++) {
    const row = raw[r] || [];
    const model = (row[2] || '').toString().trim();
    if (!model || model === '총합계') continue;

    monthIndices.forEach(colIdx => {
        const q = parseInt(row[colIdx]) || 0;
        if (q > 0) totalCount += q;
    });
}

console.log('Total Rows with Indices [4, 7, 8, 9, 10, 12]:', totalCount);
