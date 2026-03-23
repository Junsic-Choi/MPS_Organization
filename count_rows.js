const fs = require('fs');
try {
    const data = fs.readFileSync('_FinalList_MPS.csv', 'utf8');
    const lines = data.split('\n').filter(l => l.trim() !== '');
    fs.writeFileSync('csv_count.txt', `TOTAL_LINES: ${lines.length}`);
} catch (e) {
    fs.writeFileSync('csv_count.txt', `ERROR: ${e.message}`);
}
