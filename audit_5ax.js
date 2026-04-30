const fs = require('fs');
const XLSX = require('xlsx');

function getMatchKey(s) {
    if (!s) return '';
    let k = s.toString().toUpperCase().trim();
    if (k.startsWith('VF8')) k = 'VCF850' + k.substring(3);
    if (k.startsWith('M') && !k.startsWith('MYNX')) k = 'MYNX' + k.substring(1);
    k = k.replace(/VIII/g, '8').replace(/VII/g, '7').replace(/VI/g, '6').replace(/IV/g, '4').replace(/IX/g, '9');
    k = k.replace(/III/g, '3').replace(/II/g, '2');
    k = k.split('-')[0];
    k = k.replace(/PUMA|LYNX/g, '').replace(/^P|^L/, '');
    // Checking the current faulty regex
    return k.replace(/[^A-Z0-9]/g, '');
}

function solve() {
    const buf = fs.readFileSync('MPS2603-1.xlsx');
    const wb = XLSX.read(buf, {type:'buffer'});
    const data = XLSX.utils.sheet_to_json(wb.Sheets['MPS'], {header:1});

    for(let r=5; r<data.length; r++) {
        const row = data[r] || [];
        const prod = (row[4]||'').toString().toUpperCase();
        if (prod.includes('5AX')) {
            console.log(`MPS Product with 5AX:`, prod, ` -> Key:`, getMatchKey(prod));
        }
    }
}
solve();
