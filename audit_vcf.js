const XLSX = require('xlsx');
const fs = require('fs');

try {
    const file = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\MPS2603-1.xlsx';
    const wb = XLSX.readFile(file);
    const sheet0 = wb.Sheets[wb.SheetNames[0]]; 
    const mpsSheet = wb.Sheets[wb.SheetNames[1]];
    
    // Master Sheet Search
    const masterData = XLSX.utils.sheet_to_json(sheet0, { header: 1 });
    let masterFound = [];
    masterData.forEach((r, i) => {
        if (JSON.stringify(r).includes('VCF')) masterFound.push(`Row ${i+1}: ${JSON.stringify(r)}`);
    });

    // MPS Sheet Search
    const mpsData = XLSX.utils.sheet_to_json(mpsSheet, { header: 1 });
    let mpsFound = [];
    mpsData.forEach((r, i) => {
        if (JSON.stringify(r).includes('VCF')) mpsFound.push(`Row ${i+1}: ${JSON.stringify(r)}`);
    });

    let out = '--- Master Sheet (Sheet 0) ---\n' + masterFound.join('\n') + '\n\n';
    out += '--- MPS Sheet ---\n' + mpsFound.join('\n');

    fs.writeFileSync('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\vcf_audit.txt', out, 'utf8');
    console.log('Done');
} catch (err) {
    console.error(err);
}
