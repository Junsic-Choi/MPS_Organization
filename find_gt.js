const ExcelJS = require('exceljs');
const path = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\일반비_MPS2603-1(생산배포용).xlsx';
const wb = new ExcelJS.Workbook();
wb.xlsx.readFile(path).then(() => {
    const ws = wb.getWorksheet('생산배포용') || wb.worksheets[1];
    let found = false;
    ws.eachRow((row, rowNum) => {
        const cellVal = String(row.getCell(3).value || "");
        if (cellVal.includes('GT')) {
            console.log(`FOUND GT at row ${rowNum}: ${cellVal}`);
            found = true;
        }
    });
    if (!found) console.log("NO GT FOUND IN SHEET");
}).catch(err => {
    console.error('Error reading excel:', err.message);
});
