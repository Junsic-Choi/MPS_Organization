const fs = require('fs');
const path = require('path');
const p = path.join(__dirname, 'persistence_test.txt');
try {
    fs.writeFileSync(p, 'SUCCESS_WRITE_' + Date.now());
    console.log('WRITE_OK: ' + p);
} catch (e) {
    console.error('WRITE_FAIL: ' + e.message);
}
