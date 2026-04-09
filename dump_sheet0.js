const XLSX = require('xlsx');
const fs = require('fs');
const workbook = XLSX.readFile('MPS2603-1.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
let output = "";
data.forEach((row, r) => {
    output += `R${r}: ` + JSON.stringify(row) + "\n";
});
fs.writeFileSync('sheet0_full_dump.txt', output);
console.log('Dumped ' + data.length + ' rows to sheet0_full_dump.txt');
