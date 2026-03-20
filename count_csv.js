const fs = require('fs');
const content = fs.readFileSync('_FinalList_utf8.csv', 'utf8');
const lines = content.split('\n');
fs.writeFileSync('FINAL_COUNT.txt', lines.length.toString());
