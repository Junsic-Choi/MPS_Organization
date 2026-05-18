const XLSX = require('xlsx');
const fs = require('fs');

function getMatchKey(s) {
    if (!s) return '';
    let n = s.toString().toUpperCase().trim();
    n = n.replace(/PUMA|LYNX/g, '').replace(/\s+/g, '').trim();
    if (n.startsWith('P') || (n.startsWith('L') && !n.startsWith('LEO'))) {
        n = n.substring(1);
    }
    n = n.replace(/VIII/g, '8').replace(/VII/g, '7').replace(/VI/g, '6').replace(/IV/g, '4').replace(/IX/g, '9');
    n = n.replace(/III/g, '3').replace(/II/g, '2');
    if (n.includes('SMX')) {
        n = n.replace(/SMX2(?![0-9])/g, 'SMX21');
        n = n.replace(/2100/g, '21').replace(/3100/g, '31').replace(/5100/g, '51');
        n = n.replace(/SYYB/g, 'SYY').replace(/STB/g, 'SB');
    }
    if (n.includes('ST38GS')) return 'ST38GS2'; 
    if (n.includes('ST10GS')) return 'ST1GS2';
    if (n.includes('DST20')) return 'DST20';
    if (n.includes('LEO16')) return 'LEO16';
    let key = n.replace(/[^A-Z1-9]/g, '');
    if (key === 'DST2') key = 'DST20';
    key = key.replace(/0/g, '');
    key = key.replace(/[A-Z]+$/, '');
    return key;
}

function extractMonth(s) {
    if (!s) return null;
    const str = s.toString().trim();
    const yearMatch = str.match(/26\.(\d+)/);
    if (yearMatch) return parseInt(yearMatch[1]);
    const monthWordMatch = str.match(/(\d+)\s*월/);
    if (monthWordMatch) return parseInt(monthWordMatch[1]);
    if (/^\d+$/.test(str)) return parseInt(str);
    return null;
}

const file = 'MPS2605-1.xlsx';
const wb = XLSX.readFile(file);
const masterWs = wb.Sheets['MPS'] || wb.Sheets[wb.SheetNames[1]];
const masterRaw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });

console.log("Searching Master Plan for LEO and G-Tech...");

let monthRowIdx = -1;
let typeRowIdx = -1;
for (let r = 0; r < 50; r++) {
    const rowStr = (masterRaw[r] || []).join('|');
    if (rowStr.includes('월')) monthRowIdx = r;
    if (rowStr.includes('생산') && rowStr.includes('판매') && r > monthRowIdx) {
        typeRowIdx = r;
        break;
    }
}

const masterMonthCols = [];
if (monthRowIdx !== -1 && typeRowIdx !== -1) {
    masterRaw[typeRowIdx].forEach((cell, idx) => {
        if ((cell || '').toString().trim() === '생산') {
            for (let c = idx; c >= 0; c--) {
                const mNum = extractMonth(masterRaw[monthRowIdx][c]);
                if (mNum !== null) {
                    masterMonthCols.push({ name: mNum + '월', col: idx });
                    break;
                }
            }
        }
    });
}

masterRaw.forEach((row, idx) => {
    if (idx <= typeRowIdx) return;
    const site = (row[6] || '').toString().trim();
    const model = (row[3] || row[4] || '').toString().trim();
    const equipment = (row[2] || '').toString().trim();

    if (model.includes('LEO') || model.includes('ST38') || model.includes('DST20') || model.includes('ST10') || equipment.includes('LEO') || equipment.includes('S_turn')) {
        console.log(`Row ${idx}: Site=${site}, Equipment=${equipment}, Model=${model}`);
        masterMonthCols.forEach(mCol => {
            const q = row[mCol.col];
            if (q > 0) console.log(`  ${mCol.name}: Qty ${q}`);
        });
    }
});
