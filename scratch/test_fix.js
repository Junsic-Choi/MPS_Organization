const XLSX = require('xlsx');

function processMpsFileMock(filename) {
    const wb = XLSX.readFile(filename);
    const masterWsName = wb.SheetNames.find(name => name.includes('MPS') || name.includes('Master')) || wb.SheetNames[1];
    const masterWs = wb.Sheets[masterWsName];
    const masterRaw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });

    let monthRowIdx = -1;
    let typeRowIdx = -1;

    for (let r = 0; r < Math.min(50, masterRaw.length); r++) {
        const rowStr = (masterRaw[r] || []).join('|');
        if (rowStr.includes('월') && monthRowIdx === -1) monthRowIdx = r;
        if (rowStr.includes('생산') && rowStr.includes('판매') && r > monthRowIdx) {
            typeRowIdx = r;
            break;
        }
    }

    const masterMonthCols = [];
    if (monthRowIdx !== -1 && typeRowIdx !== -1) {
        const monthRow = masterRaw[monthRowIdx];
        const typeRow = masterRaw[typeRowIdx];

        typeRow.forEach((cell, idx) => {
            const type = (cell || '').toString().trim();
            if (type === '생산') {
                for (let c = idx; c >= 0; c--) {
                    const mNum = extractMonth(monthRow[c]);
                    if (mNum !== null) {
                        masterMonthCols.push({ name: mNum + '월', col: idx });
                        break;
                    }
                }
            }
        });
    }

    function extractMonth(s) {
        if (!s) return null;
        if (s instanceof Date) return s.getMonth() + 1;
        const str = s.toString().trim();
        const dotMatch = str.match(/(?:20)?26\.(\d+)/);
        if (dotMatch) return parseInt(dotMatch[1]);
        const monthWordMatch = str.match(/(\d+)\s*월/);
        if (monthWordMatch) return parseInt(monthWordMatch[1]);
        if (/^\d+$/.test(str)) return parseInt(str);
        return null;
    }

    console.log(`\n--- File: ${filename} ---`);
    console.log(`All detected months (${masterMonthCols.length}):`, masterMonthCols.map(m => m.name));

    // Apply the 6-month limit fix
    const slicedMonths = masterMonthCols.slice(0, 6);
    console.log(`Sliced months (first 6):`, slicedMonths.map(m => m.name));

    let totalQty = 0;
    let seongjuQty = 0;

    masterRaw.forEach((row, idx) => {
        if (idx <= typeRowIdx) return;
        const site = String(row[6] || '').trim();
        const isSeongju = (site === '1842' || site.includes('성주'));

        slicedMonths.forEach(mCol => {
            const q = parseInt(row[mCol.col]) || 0;
            totalQty += q;
            if (isSeongju) seongjuQty += q;
        });
    });

    console.log(`Total Qty: ${totalQty}`);
    console.log(`Seongju Qty: ${seongjuQty}`);
}

processMpsFileMock('MPS2603-1.xlsx');
processMpsFileMock('MPS2604-1.xlsx');
processMpsFileMock('MPS2605-1.xlsx');
