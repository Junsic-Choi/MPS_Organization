const fs = require('fs');
try {
    fs.writeFileSync('node_test_status.txt', 'Node started at ' + new Date().toISOString() + '\n');
    const XLSX = require('xlsx');
    fs.appendFileSync('node_test_status.txt', 'XLSX loaded successfully\n');
    const multer = require('multer');
    fs.appendFileSync('node_test_status.txt', 'Multer loaded successfully\n');
} catch (e) {
    fs.writeFileSync('node_test_error.txt', e.stack);
}
process.exit(0);
