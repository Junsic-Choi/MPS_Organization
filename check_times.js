const fs = require('fs');
const path = require('path');

const files = fs.readdirSync('.');
const log = files.map(f => {
    const stats = fs.statSync(f);
    return `${f} : ${stats.mtime}`;
}).join('\n');

fs.writeFileSync('file_times.txt', log);
