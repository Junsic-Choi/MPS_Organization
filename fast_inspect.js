const xlsx = require('xlsx');

const OPT = { cellFormula: false, cellStyles: false, cellNF: false, cellDates: false };

const wb = xlsx.readFile('일반비_MPS2603-1(생산배포용).xlsx', OPT);
console.log('=== Sheets ===', wb.SheetNames);

// ── 생산배포용 탭 헤더 + 첫 5행
const prodName = wb.SheetNames.find(n => n.includes('생산배포용'));
const prodWs = wb.Sheets[prodName];
const prodData = xlsx.utils.sheet_to_json(prodWs, { header: 1, defval: '' });
console.log('\n=== 생산배포용 첫 8행 (A,B,C,D,E,H,I,J,K,M열) ===');
for (let r = 0; r < 8; r++) {
    const row = prodData[r] || [];
    console.log(`R${r+1}: A=[${row[0]}] B=[${row[1]}] C=[${row[2]}] D=[${row[3]}] E=[${row[4]}] H=[${row[7]}] I=[${row[8]}] J=[${row[9]}] K=[${row[10]}] M=[${row[12]}]`);
}

// ── MPS 탭 주요 열 (4행 헤더 + 5~8행 데이터)
const mpsName = wb.SheetNames.find(n => n === 'MPS');
const mpsWs = wb.Sheets[mpsName];
const mpsData = xlsx.utils.sheet_to_json(mpsWs, { header: 1, defval: '' });
console.log('\n=== MPS 탭 4~8행 (D,E,G,H,I,M,R,W,AC,AI 열) ===');
// 컬럼 번호 → 0-indexed: D=3, E=4, G=6, H=7, I=8, M=12, R=17, W=22, AC=28, AI=34, AO=40
const mpsCols = { D:3, E:4, G:6, H:7, I:8, M:12, R:17, W:22, AC:28, AI:34, AO:40 };
for (let r = 3; r < 9; r++) {
    const row = mpsData[r] || [];
    let out = `R${r+1}: `;
    for (const [col, idx] of Object.entries(mpsCols)) {
        out += `${col}=[${row[idx]}] `;
    }
    console.log(out);
}

// ── Site 탭
const siteName = wb.SheetNames.find(n => n === 'Site' || n.toLowerCase() === 'site');
if (siteName) {
    const siteWs = wb.Sheets[siteName];
    const siteData = xlsx.utils.sheet_to_json(siteWs, { header: 1, defval: '' });
    console.log('\n=== Site 탭 첫 15행 ===');
    for (let r = 0; r < 15; r++) {
        console.log(`R${r+1}: `, (siteData[r] || []).slice(0, 8).join(' | '));
    }
}

// ── 생산배포용 총 행 수 & 월 수량합 확인
console.log('\n=== 생산배포용 총 데이터행 수 ===', prodData.length - 1);
// E,H,I,J,K,M 합계
let totals = { E:0, H:0, I:0, J:0, K:0, M:0 };
for (let r = 1; r < prodData.length; r++) {
    const row = prodData[r] || [];
    totals.E += Number(row[4]) || 0;
    totals.H += Number(row[7]) || 0;
    totals.I += Number(row[8]) || 0;
    totals.J += Number(row[9]) || 0;
    totals.K += Number(row[10]) || 0;
    totals.M += Number(row[12]) || 0;
}
console.log('월별 합계:', totals, '총합:', Object.values(totals).reduce((a,b)=>a+b,0));

// ── MPS 탭 I4+M4+R4+W4+AC4+AI4 합계 확인
const mpsRow4 = mpsData[3] || [];
const mpsTotal = [8,12,17,22,28,34].reduce((s,c)=> s + (Number(mpsRow4[c])||0), 0);
console.log('\nMPS I4+M4+R4+W4+AC4+AI4 합계:', mpsTotal);
console.log('MPS 각열 값: I=', mpsRow4[8], 'M=', mpsRow4[12], 'R=', mpsRow4[17], 'W=', mpsRow4[22], 'AC=', mpsRow4[28], 'AI=', mpsRow4[34]);
