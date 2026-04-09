const XLSX = require('xlsx');
const fs = require('fs');

try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    
    // 1. Inspect 생산배포용 (Sheet 0)
    const s0Name = workbook.SheetNames[0];
    const s0Data = XLSX.utils.sheet_to_json(workbook.Sheets[s0Name], { header: 1 });
    let s0Log = `--- Sheet 0: ${s0Name} ---\n`;
    for(let r=0; r<100; r++) if(s0Data[r]) s0Log += `Row ${r}: ${JSON.stringify(s0Data[r])}\n`;
    fs.writeFileSync('debug_s0_raw.txt', s0Log);

    // 2. Inspect MPS (Sheet 1)
    const s1Name = workbook.SheetNames.find(n => n.toUpperCase().includes('MPS')) || workbook.SheetNames[1];
    const s1Data = XLSX.utils.sheet_to_json(workbook.Sheets[s1Name], { header: 1 });
    let s1Log = `--- Sheet 1: ${s1Name} ---\n`;
    for(let r=0; r<200; r++) if(s1Data[r]) s1Log += `Row ${r}: ${JSON.stringify(s1Data[r])}\n`;
    fs.writeFileSync('debug_s1_raw.txt', s1Log);

    console.log("Dump done.");
} catch (e) {
    console.error(e);
}
