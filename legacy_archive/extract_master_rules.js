const XLSX = require('xlsx');
const fs = require('fs');

try {
    const filePath = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\Real site.xlsx';
    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
    
    const mapping = [];
    const seen = new Set();
    
    data.forEach((row, idx) => {
        if (idx < 1) return;
        const ver = (row[1] || '').toString().trim();
        const desc = (row[2] || '').toString().trim();
        
        // "07. 휴텍" 같은 형식이 포함되어 있는지 확인하거나, 그냥 모든 쌍을 가져옴
        if (ver && desc && !seen.has(ver)) {
            // 설명문구에서 사이트 명칭만 가급적 추출 (예: "21. 휴텍" 등)
            // 정규표현식으로 "숫자. 문자" 패턴을 찾음
            const match = desc.match(/\d{2}\.\s*[\wㄱ-ㅎ가-힣]+/);
            const siteName = match ? match[0] : desc;
            
            mapping.push(`${ver}: ${siteName}`);
            seen.add(ver);
        }
    });
    
    fs.writeFileSync('realsite_master_list.txt', mapping.join('\n'));
    console.log(`Extracted ${mapping.length} unique rules.`);
} catch (e) {
    console.error(e);
}
