const fs = require('fs');
const xlsx = require('xlsx');

try {
    const workbook = xlsx.readFile('test.xlsx');
    let out = 'Sheets: ' + workbook.SheetNames.join(', ') + '\n';

    for (let s of workbook.SheetNames) {
        if (s.includes('MPS') || s.includes('생산배포용')) {
            out += `\n--- Sheet: ${s} ---\n`;
            const sheet = workbook.Sheets[s];
            const json = xlsx.utils.sheet_to_json(sheet, { header: 1 });

            if (json.length >= 4) {
                out += 'Row 3:\n';
                [8, 12, 17, 22, 28, 34].forEach(i => {
                    out += `  Col ${i} : ${json[2] ? json[2][i] : 'undefined'} / ${json[3] ? json[3][i] : 'undefined'}\n`;
                });
            } else {
                out += 'Not enough rows.\n';
            }
        }
    }
    fs.writeFileSync('node_out.txt', out);
} catch (e) {
    fs.writeFileSync('node_out.txt', 'Error: ' + e.message + '\n' + e.stack);
}
