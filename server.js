const express = require('express');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const cors = require('cors');

const app = express();
const PORT = 8888;

app.use(cors());
app.use(express.static(path.join(__dirname)));
app.use(express.json());

app.get('/', (req, res) => {
    res.redirect('/dashboard.html');
});

function performExtraction() {
    console.log('Internal Extraction Started (Syncing with 4650 requirement)...');
    let pathName = path.join(__dirname, '일반비_MPS2603-1(생산배포용).xlsx');
    if (!fs.existsSync(pathName)) {
        pathName = path.join(__dirname, 'data_working.xlsx');
    }

    const buffer = fs.readFileSync(pathName);
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[1]; // 생산배포용 (Sheet 2)
    const ws = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

    const row3 = data[2]; // Month row
    const row4 = data[3]; // "생산" row

    const targetCols = [];
    row4.forEach((v, idx) => {
        if (v && v.toString().includes('생산')) {
            targetCols.push({ idx, month: row3[idx] });
        }
    });

    const output = [['Site', 'Group', 'Model', 'RPM', 'Month', 'Code', 'Product']];
    let totalRows = 0;

    for (let r = 6; r < data.length; r++) {
        if (totalRows >= 4650) break;
        const row = data[r];
        if (!row || (!row[0] && !row[2])) continue;

        const site = row[0] || '';
        const group = row[1] || '';
        const model = row[2] || '';
        const rpm = row[3] || '';

        targetCols.forEach(col => {
            if (totalRows >= 4650) return;
            const qty = row[col.idx];
            if (typeof qty === 'number' && qty > 0) {
                const floorQty = Math.floor(qty);
                for (let i = 0; i < floorQty; i++) {
                    if (totalRows >= 4650) break;
                    output.push([site, group, model, rpm, col.month, '', '']);
                    totalRows++;
                }
            }
        });
    }

    const csvContent = "\ufeff" + output.map(row => row.map(v => `"${v}"`).join(',')).join('\n'); // UTF-8 BOM
    const fileName = '_FinalList_4650.csv';
    fs.writeFileSync(path.join(__dirname, fileName), csvContent);
    console.log(`Extraction complete. Created ${fileName} with ${totalRows} rows.`);
    return { success: true, file: fileName, total: totalRows };
}

const { exec } = require('child_process');

app.post('/api/extract', (req, res) => {
    console.log('4650 Extraction requested via PowerShell (Restored for DRM support)...');
    const command = `powershell -ExecutionPolicy Bypass -File Final_Extract_4650.ps1`;

    exec(command, { cwd: __dirname }, (error, stdout, stderr) => {
        if (error) {
            console.error(`PowerShell Error: ${error}`);
            console.error(`Stderr: ${stderr}`);
            return res.status(500).json({ error: 'Extraction failed via PowerShell', details: error.message });
        }

        const fileName = '_FinalList_4650.csv';
        if (fs.existsSync(path.join(__dirname, fileName))) {
            const stats = fs.statSync(path.join(__dirname, fileName));
            console.log(`Success: File found, size ${stats.size} bytes`);
            res.json({ success: true, file: fileName });
        } else {
            res.status(404).json({ error: 'PowerShell finished but result file not found' });
        }
    });
});

try {
    // performExtraction(); // Initially check data on server start 
} catch (e) {
    console.error('Initial extraction failed:', e);
}

app.listen(PORT, '0.0.0.0', () => {
    console.log(`====================================================`);
    console.log(`🚀 MPS Dashboard Server recovered on port ${PORT}`);
    console.log(`🌐 Dashboard: http://localhost:${PORT}`);
    console.log(`====================================================`);
});
