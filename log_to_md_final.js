const fs = require('fs');
try {
    const content = fs.readFileSync('final_4650_log.txt', 'utf8');
    fs.writeFileSync('final_4650_report.md', '# Final 4650 Audit Report\n\n```text\n' + content + '\n```');
} catch (e) {
    fs.writeFileSync('final_4650_report.md', 'ERROR: ' + e.message);
}
