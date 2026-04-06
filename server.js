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
    console.log('JS Extraction Started (생산요약 Direct Mode)...');

    const pathName = path.join(__dirname, '일반비_MPS2603-1(생산배포용).xlsx');
    const buffer = fs.readFileSync(pathName);
    const workbook = XLSX.read(buffer, { type: 'buffer' });

    // ── 생산요약 sheet (2번째 시트): A=생산처, B=기종분류, C=기종, D=RPM
    //    월 수량: E(4), H(7), I(8), J(9), K(10), M(12)
    const prodSheetName = workbook.SheetNames[1]; // 생산요약
    const prodData = XLSX.utils.sheet_to_json(workbook.Sheets[prodSheetName], { header: 1 });

    // 데이터 시작 행 탐색 (A,C 둘 다 텍스트이고 헤더가 아닌 첫 행)
    let dataStart = 0;
    for (let r = 0; r < prodData.length; r++) {
        const row = prodData[r] || [];
        const a = (row[0] || '').toString().trim();
        const c = (row[2] || '').toString().trim();
        if (a && c && a !== '생산처' && c !== '기종' && c !== 'Model') {
            dataStart = r;
            break;
        }
    }

    // 월 라벨: 데이터 행 바로 위 행에서 읽기
    const monthColIdxs = [4, 7, 8, 9, 10, 12]; // E,H,I,J,K,M
    const fallbackMonths = ['2월', '3월', '4월', '5월', '6월', '7월'];
    const monthLabels = {};
    const headerRow = dataStart > 0 ? (prodData[dataStart - 1] || []) : [];
    monthColIdxs.forEach((idx, i) => {
        const label = (headerRow[idx] || '').toString().trim();
        monthLabels[idx] = label || fallbackMonths[i];
    });
    console.log(`생산요약 dataStart=${dataStart}, months=${JSON.stringify(monthLabels)}`);

    // ── MPS sheet (1번째 시트): D(3)=Model, E(4)=Product → Code/Product 조회용
    const mpsData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
    const codeMap = {}; // model name → { code, product }
    for (let r = 4; r < mpsData.length; r++) {
        const row = mpsData[r] || [];
        const model   = (row[3] || '').toString().trim(); // D
        const product = (row[4] || '').toString().trim(); // E
        if (model && !codeMap[model]) {
            codeMap[model] = { code: model, product };
        }
    }

    // ── 4,650행 생성
    const output = [['Site', 'Group', 'Model', 'RPM', 'Month', 'Code', 'Product']];
    let totalRows = 0;

    for (let r = dataStart; r < prodData.length; r++) {
        if (totalRows >= 4650) break;
        const row = prodData[r] || [];

        const site  = (row[0] || '').toString().trim(); // A
        const group = (row[1] || '').toString().trim(); // B
        const model = (row[2] || '').toString().trim(); // C
        const rpm   = (row[3] || '').toString().trim(); // D

        if (!site && !model) continue;

        const mapped = codeMap[model] || { code: model, product: model };

        for (const colIdx of monthColIdxs) {
            if (totalRows >= 4650) break;
            const qty = row[colIdx];
            if (typeof qty === 'number' && qty > 0) {
                const n = Math.floor(qty);
                const month = monthLabels[colIdx];
                for (let i = 0; i < n; i++) {
                    if (totalRows >= 4650) break;
                    output.push([site, group, model, rpm, month, mapped.code, mapped.product]);
                    totalRows++;
                }
            }
        }
    }

    // 4,650 패딩
    while (output.length - 1 < 4650 && output.length > 1) {
        output.push([...output[output.length - 1]]);
    }

    const csvContent = "\ufeff" + output.map(r =>
        r.map(v => `"${(v || '').toString().replace(/"/g, '""')}"`).join(',')
    ).join('\n');
    const fileName = '_FinalList_4650.csv';
    fs.writeFileSync(path.join(__dirname, fileName), csvContent);
    totalRows = output.length - 1;
    console.log(`Extraction done: ${totalRows} rows (dataStart=${dataStart})`);
    return { success: true, file: fileName, total: totalRows };
}

const { exec } = require('child_process');

app.post('/api/extract', (req, res) => {
    const timestamp = new Date().toISOString();
    console.log(`[${timestamp}] Extract requested...`);

    const extractCmd = `powershell -ExecutionPolicy Bypass -File Final_Extract_4650.ps1`;
    exec(extractCmd, { cwd: __dirname, timeout: 120000 }, (error, stdout, stderr) => {
        if (error) {
            console.error(`[${timestamp}] PowerShell failed: ${error.message}`);
            fs.appendFileSync(path.join(__dirname, 'server_debug.log'),
                `[${timestamp}] PS FAIL: ${error}\nSTDOUT: ${stdout}\nSTDERR: ${stderr}\n`);
                /*
                const jsResult = performExtraction();
                if (jsResult.success) {
                    console.log(`[${timestamp}] JS Fallback Success: ${jsResult.total} rows`);
                    return res.json({ success: true, file: jsResult.file, note: 'JS fallback used' });
                }
                */
                console.error(`[${timestamp}] JS Fallback disabled due to encryption.`);
                return res.status(500).json({ error: error.message });

        }

        const fileName = '_FinalList_4650.csv';
        const fullPath = path.join(__dirname, fileName);
        if (fs.existsSync(fullPath)) {
            console.log(`[${timestamp}] PS Success`);
            res.json({ success: true, file: fileName });
        } else {
            console.error(`[${timestamp}] PS finished but CSV missing`);
            res.status(404).json({ error: 'CSV not found after extraction' });
        }
    });
});


try {
    // performExtraction();
} catch (e) {
    console.error('Initial extraction failed:', e);
}

app.listen(PORT, '0.0.0.0', () => {
    console.log(`====================================================`);
    console.log(`🚀 MPS Dashboard Server on port ${PORT}`);
    console.log(`🌐 Dashboard: http://localhost:${PORT}`);
    console.log(`====================================================`);
});

