const fs = require('fs');
try {
    const content = fs.readFileSync('category_audit_log.txt', 'utf8');
    fs.writeFileSync('audit_report.md', '# Category Audit Report\n\n```text\n' + content + '\n```');
} catch (e) {
    fs.writeFileSync('audit_report.md', 'ERROR: ' + e.message);
}
