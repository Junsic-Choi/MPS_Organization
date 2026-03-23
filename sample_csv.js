const fs = require('fs');
try {
    const data = fs.readFileSync('_FinalList.csv', 'utf8');
    const lines = data.split('\n');
    console.log('Total Lines:', lines.length);
    console.log('Sample (index 5000-5020):');
    console.log(lines.slice(5000, 5021).join('\n'));
    fs.writeFileSync('csv_sample_report.txt', lines.slice(5000, 5021).join('\n'));
} catch (e) {
    console.error(e.message);
}
