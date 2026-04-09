const XLSX = require('xlsx');
const fs = require('fs');

try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    const monthNames = ["2월", "3월", "4월", "5월", "6월", "7월"];
    
    // Sheet 1 Summary
    const mpsSheet = workbook.Sheets[workbook.SheetNames[1]];
    const mpsRaw = XLSX.utils.sheet_to_json(mpsSheet, { header: 1 });
    const mpsMonthIdxs = [8, 12, 17, 22, 28, 34];
    const mpsSums = {};
    for (let r = 5; r < mpsRaw.length; r++) {
        const row = mpsRaw[r] || [];
        const pl = (row[1] || '').toString().trim();
        if (!pl) continue;
        if (!mpsSums[pl]) mpsSums[pl] = [0, 0, 0, 0, 0, 0];
        mpsMonthIdxs.forEach((idx, i) => {
            mpsSums[pl][i] += (parseInt(row[idx]) || 0);
        });
    }

    // Sheet 0 Summary
    const prodSheet = workbook.Sheets[workbook.SheetNames[0]];
    const prodRaw = XLSX.utils.sheet_to_json(prodSheet, { header: 1 });
    const prodMonthIdxs = [4, 7, 8, 9, 10, 12];
    const siteSums = {};
    let lastSite = "";
    for (let r = 6; r < prodRaw.length; r++) {
        const row = prodRaw[r] || [];
        if (row[0]) lastSite = row[0].toString().trim();
        if (!lastSite || lastSite.includes('총합계')) continue;
        if (!siteSums[lastSite]) siteSums[lastSite] = [0, 0, 0, 0, 0, 0];
        prodMonthIdxs.forEach((idx, i) => {
            siteSums[lastSite][i] += (parseInt(row[idx]) || 0);
        });
    }

    let report = "--- Sheet 1 (MPS) PL Sums ---\n";
    Object.keys(mpsSums).forEach(pl => {
        report += `${pl}: ${mpsSums[pl].join(', ')} (Total: ${mpsSums[pl].reduce((a,b)=>a+b,0)})\n`;
    });

    report += "\n--- Sheet 0 (Prod) Site Sums ---\n";
    Object.keys(siteSums).forEach(site => {
        report += `${site}: ${siteSums[site].join(', ')} (Total: ${siteSums[site].reduce((a,b)=>a+b,0)})\n`;
    });

    fs.writeFileSync('alignment_diagnostic.txt', report);

} catch (e) {
    fs.writeFileSync('alignment_diagnostic.txt', 'ERROR: ' + e.message);
}
