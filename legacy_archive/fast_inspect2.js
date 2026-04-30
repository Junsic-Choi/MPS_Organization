const xlsx = require('xlsx');
const fs = require('fs');

const OPT = { cellFormula: false, cellStyles: false, cellNF: false, cellDates: false };
const lines = [];

try {
    const wb = xlsx.readFile('일반비_MPS2603-1(생산배포용).xlsx', OPT);
    lines.push('=== Sheets === ' + wb.SheetNames.join(', '));

    // 생산배포용 탭
    const prodName = wb.SheetNames.find(n => n.includes('생산배포용'));
    const prodWs = wb.Sheets[prodName];
    const prodData = xlsx.utils.sheet_to_json(prodWs, { header: 1, defval: '' });
    lines.push('\n=== 생산배포용: ' + prodName + ' / 총 행수: ' + (prodData.length - 1));
    lines.push('=== 첫 8행 (A,B,C,D,E,H,I,J,K,M) ===');
    for (let r = 0; r < 8; r++) {
        const row = prodData[r] || [];
        lines.push(`R${r+1}: A=[${row[0]}] B=[${row[1]}] C=[${row[2]}] D=[${row[3]}] E=[${row[4]}] H=[${row[7]}] I=[${row[8]}] J=[${row[9]}] K=[${row[10]}] M=[${row[12]}]`);
    }

    // 생산배포용 E,H,I,J,K,M 합계
    let totals = { E:0, H:0, I:0, J:0, K:0, M:0 };
    let rowCount = 0;
    for (let r = 1; r < prodData.length; r++) {
        const row = prodData[r] || [];
        const e = Number(row[4])||0, h = Number(row[7])||0, i = Number(row[8])||0,
              j = Number(row[9])||0, k = Number(row[10])||0, m = Number(row[12])||0;
        if (e+h+i+j+k+m > 0) rowCount++;
        totals.E+=e; totals.H+=h; totals.I+=i; totals.J+=j; totals.K+=k; totals.M+=m;
    }
    const grandTotal = Object.values(totals).reduce((a,b)=>a+b,0);
    lines.push('월별 합계: ' + JSON.stringify(totals) + ' / 총합: ' + grandTotal + ' / 유효행: ' + rowCount);

    // MPS 탭
    const mpsName = wb.SheetNames.find(n => n === 'MPS');
    const mpsWs = wb.Sheets[mpsName];
    const mpsData = xlsx.utils.sheet_to_json(mpsWs, { header: 1, defval: '' });
    lines.push('\n=== MPS 탭 3~8행 (D,E,G,H,I,M,R,W,AC,AI,AO) ===');
    // 0-indexed: D=3,E=4,G=6,H=7,I=8,M=12,R=17,W=22,AC=28,AI=34,AO=40
    const cols = [['D',3],['E',4],['G',6],['H',7],['I',8],['M',12],['R',17],['W',22],['AC',28],['AI',34],['AO',40]];
    for (let r = 2; r < 9; r++) {
        const row = mpsData[r] || [];
        let out = `R${r+1}: `;
        cols.forEach(([c,i]) => { out += `${c}=[${row[i]}] `; });
        lines.push(out);
    }
    const mpsRow4 = mpsData[3] || [];
    const mpsTotal = [8,12,17,22,28,34].reduce((s,c)=> s+(Number(mpsRow4[c])||0), 0);
    lines.push('MPS I4+M4+R4+W4+AC4+AI4 합계: ' + mpsTotal);

    // MPS 탭 총 행수
    lines.push('MPS 총 행수: ' + mpsData.length);

    // Site 탭
    const siteName = wb.SheetNames.find(n => n.toLowerCase().includes('site'));
    if (siteName) {
        const siteData = xlsx.utils.sheet_to_json(wb.Sheets[siteName], { header: 1, defval: '' });
        lines.push('\n=== Site 탭: ' + siteName + ' / 첫 20행 ===');
        for (let r = 0; r < 20; r++) {
            lines.push(`R${r+1}: ` + (siteData[r]||[]).slice(0,10).map(v=>`[${v}]`).join(' '));
        }
    }

} catch(e) {
    lines.push('ERROR: ' + e.message);
    lines.push(e.stack);
}

fs.writeFileSync('fast_inspect_result.txt', lines.join('\n'), 'utf8');
