const XLSX = require('xlsx');

const files = ['MPS2603-1.xlsx', 'MPS2604-1.xlsx', 'MPS2605-1.xlsx'];

files.forEach(file => {
    try {
        const wb = XLSX.readFile(file);
        const wsName = wb.SheetNames.find(n => n.includes('MPS') || n.includes('Master')) || wb.SheetNames[0];
        const ws = wb.Sheets[wsName];
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });
        
        // Find row that contains "실적"
        let monthRowIdx = -1;
        let typeRowIdx = -1;

        for (let r = 0; r < Math.min(50, raw.length); r++) {
            const rowStr = (raw[r] || []).join('|');
            if (rowStr.includes('월') && monthRowIdx === -1) monthRowIdx = r;
            if (rowStr.includes('생산') && rowStr.includes('판매') && r > monthRowIdx) {
                typeRowIdx = r;
                break;
            }
        }

        if (monthRowIdx !== -1 && typeRowIdx !== -1) {
            const monthRow = raw[monthRowIdx];
            const typeRow = raw[typeRowIdx];
            const months = [];
            
            typeRow.forEach((cell, idx) => {
                const type = (cell || '').toString().trim();
                if (type === '생산') {
                    for (let c = idx; c >= 0; c--) {
                        const mStr = String(monthRow[c] || '').trim();
                        if (mStr.includes('월')) {
                            months.push({ name: mStr, idx });
                            break;
                        }
                    }
                }
            });
            console.log(`\nFile: ${file}`);
            console.log('Detected months:', months.map(m => m.name));
        } else {
            console.log(`\nFile: ${file} - Could not detect month structure`);
        }
    } catch (e) {
        console.log(`\nFile: ${file} - Error: ${e.message}`);
    }
});
