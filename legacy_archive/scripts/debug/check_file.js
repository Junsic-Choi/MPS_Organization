const fs = require('fs');
try {
    const buf = fs.readFileSync('data_working.xlsx');
    console.log('Size:', buf.length);
    console.log('Header:', buf.slice(0, 4).toString('hex')); // PK is 504b0304
    fs.writeFileSync('node_file_check.txt', `Size: ${buf.length}, Header: ${buf.slice(0, 4).toString('hex')}`);
} catch (e) {
    fs.writeFileSync('node_file_check.txt', 'ERROR: ' + e.message);
}
