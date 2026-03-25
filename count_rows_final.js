const fs = require('fs');
const path = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\_FinalList_4650.csv';
const data = fs.readFileSync(path, 'utf8');
const lines = data.split('\n').filter(l => l.trim().length > 0);
console.log('Final Row Count (including header):', lines.length);
console.log('Sample Row:', lines[lines.length - 1]);
