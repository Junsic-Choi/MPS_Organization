const XLSX = require('xlsx');
const fs = require('fs');

try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    
    // Sheet 1 (MPS) Profiles
    const mpsSheet = workbook.Sheets[workbook.SheetNames[1]];
    const mpsRaw = XLSX.utils.sheet_to_json(mpsSheet, { header: 1 });
    const mpsMonthIdxs = [8, 12, 17, 22, 28, 34];
    const mpsProfiles = {};
    for (let r = 5; r < mpsRaw.length; r++) {
        const row = mpsRaw[r] || [];
        const pl = (row[1] || '').toString().trim();
        if (!pl) continue;
        if (!mpsProfiles[pl]) mpsProfiles[pl] = [0, 0, 0, 0, 0, 0];
        mpsMonthIdxs.forEach((idx, i) => {
            mpsProfiles[pl][i] += (parseInt(row[idx]) || 0);
        });
    }

    // Sheet 0 (Prod) Profiles
    const prodSheet = workbook.Sheets[workbook.SheetNames[0]];
    const prodRaw = XLSX.utils.sheet_to_json(prodSheet, { header: 1 });
    const prodMonthIdxs = [4, 7, 8, 9, 10, 12];
    const siteProfiles = {};
    let lastSite = "";
    for (let r = 6; r < prodRaw.length; r++) {
        const row = prodRaw[r] || [];
        if (row[0]) lastSite = row[0].toString().trim();
        if (!lastSite || lastSite.includes('총합계')) continue;
        if (!siteProfiles[lastSite]) siteProfiles[lastSite] = [0, 0, 0, 0, 0, 0];
        prodMonthIdxs.forEach((idx, i) => {
            siteProfiles[lastSite][i] += (parseInt(row[idx]) || 0);
        });
    }

    let report = "--- Profile Match Analysis ---\n";
    const sites = Object.keys(siteProfiles);
    const pls = Object.keys(mpsProfiles);
    
    sites.forEach(site => {
        const sProf = siteProfiles[site].join(',');
        const matches = pls.filter(pl => mpsProfiles[pl].join(',') === sProf);
        report += `${site} [${sProf}] => ${matches.length > 0 ? matches.join(' / ') : "NONE"}\n`;
    });

    // Special check for Henex
    report += `\nSpecial Check: Henex Site [${siteProfiles["20. 헤넥스"] ? siteProfiles["20. 헤넥스"].join(',') : 'N/A'}]\n`;
    pls.forEach(pl => {
        report += `PL ${pl}: [${mpsProfiles[pl].join(',')}]\n`;
    });

    fs.writeFileSync('profile_matches.txt', report);

} catch (e) {
    fs.writeFileSync('profile_matches.txt', 'ERROR: ' + e.message);
}
