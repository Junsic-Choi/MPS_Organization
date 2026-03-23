const fs = require('fs');
try {
    const data = fs.readFileSync('_FinalList.csv', 'utf8');
    const lines = data.split('\n');
    fs.writeFileSync('csv_start_sample.txt', lines.slice(0, 51).join('\n'));
} catch (e) {
    console.error(e.message);
}
