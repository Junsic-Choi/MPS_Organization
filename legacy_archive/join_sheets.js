// join_sheets.js
const XLSX = require('xlsx');
const fs = require('fs');
try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    
    // 1. Load MPS Mapping (Sheet 1)
    const mpsData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[1]], { header: 1 });
    const codeMap = {}; // Model Name -> Code
    for (let r = 5; r < mpsData.length; r++) {
        const row = mpsData[r] || [];
        const code = (row[3] || '').toString().trim();
        const prod = (row[4] || '').toString().trim();
        if (code && prod) {
            // Store by product ID and a normalized model name
            codeMap[prod] = code;
            const norm = prod.split('-')[0].replace(/[^A-Z0-9]/g, '');
            codeMap[norm] = code;
        }
    }

    // 2. Process 생산배포용 (Sheet 0)
    const prodData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
    let log = "Sample Matches:\n";
    let found = 0;
    for (let r = 6; r < 20; r++) {
        const row = prodData[r] || [];
        const group = row[1];
        const model = row[2];
        if (model) {
            const norm = model.replace(/[^A-Z0-9]/g, '');
            const code = codeMap[norm] || "NOT_FOUND";
            log += `Model=${model}, Norm=${norm}, Code=${code}\n`;
            if (code !== "NOT_FOUND") found++;
        }
    }
    fs.writeFileSync('join_audit.txt', log + `\nTotal Found in sample: ${found}`);
} catch (e) {
    fs.writeFileSync('join_audit.txt', 'ERROR: ' + e.message);
}
