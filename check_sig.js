const fs = require('fs');
const path = '일반비_MPS2603-1(생산배포용).xlsx';
let output = '';

try {
    const buffer = Buffer.alloc(8);
    const fd = fs.openSync(path, 'r');
    fs.readSync(fd, buffer, 0, 8, 0);
    fs.closeSync(fd);

    output += 'Hex Signature: ' + buffer.toString('hex') + '\n';
    output += 'ASCII Signature: ' + buffer.toString('ascii') + '\n';
} catch (e) {
    output += 'Error: ' + e.message + '\n';
}
fs.writeFileSync('sig_output.txt', output);
