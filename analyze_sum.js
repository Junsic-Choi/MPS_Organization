const XLSX = require('xlsx');
const workbook = XLSX.readFile('MPS2603-1.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
const monthColIdxs = [4, 7, 8, 9, 10, 12];
let cumulativeSum = 0;
let henexSum = 0;
let henexStarted = false;

data.forEach((row, r) => {
    if (r < 6) return;
    const site = row[0] ? row[0].toString().trim() : "";
    if (site.includes('헤넥스')) henexStarted = true;
    
    let rowSum = 0;
    monthColIdxs.forEach(c => {
        const val = parseInt(row[c]) || 0;
        rowSum += val;
    });
    
    if (henexStarted && !site.includes('총합계')) {
        henexSum += rowSum;
    } else if (!site.includes('총합계')) {
        cumulativeSum += rowSum;
    }
    
    if (site.includes('헤넥스')) {
        console.log(`Row ${r}: Henex found. Current Cumulative Sum before Henex: ${cumulativeSum}`);
    }
});

console.log(`Total count before Henex: ${cumulativeSum}`);
console.log(`Henex total count: ${henexSum}`);
console.log(`Grand Total: ${cumulativeSum + henexSum}`);
