const XLSX = require('xlsx');
const fs = require('fs');

function norm(s) {
    if (!s) return "";
    let n = s.toString().toUpperCase().replace(/[^A-Z0-9]/g, '');
    n = n.replace(/^PUMA/, 'P');
    n = n.replace(/^LYNX/, 'L');
    if (n.endsWith('II')) n = n.slice(0, -2) + '2';
    if (n.endsWith('III')) n = n.slice(0, -3) + '3';
    return n;
}

try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    const mpsSheet = workbook.SheetNames.find(n => n.toUpperCase().includes('MPS')) || workbook.SheetNames[1];
    const mpsData = XLSX.utils.sheet_to_json(workbook.Sheets[mpsSheet], { header: 1 });
    
    let inventory = "--- MPS SHEET INVENTORY ---\n";
    for(let r=5; r<mpsData.length; r++) {
        const row = mpsData[r] || [];
        const c = (row[3] || '').toString().trim();
        const p = (row[4] || '').toString().trim();
        if (c && p) inventory += `${c} | ${p} | (Norm: ${norm(p)})\n`;
    }
    fs.writeFileSync('mps_inventory.txt', inventory);

} catch (e) {
    fs.writeFileSync('mps_inventory.txt', 'ERROR: ' + e.message);
}
