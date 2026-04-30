const fs = require('fs');
const path = require('path');

async function check() {
    let log = "--- System Check ---\n";
    log += `Time: ${new Date().toLocaleString()}\n`;
    log += `CWD: ${process.cwd()}\n`;

    try {
        require('express');
        log += "Express: OK\n";
    } catch(e) { log += "Express: MISSING\n"; }

    try {
        require('xlsx');
        log += "XLSX: OK\n";
    } catch(e) { log += "XLSX: MISSING\n"; }

    const xlsxPath = path.join(__dirname, 'MPS2603-1.xlsx');
    if (fs.existsSync(xlsxPath)) {
        log += `File Found: ${xlsxPath}\n`;
        try {
            const fd = fs.openSync(xlsxPath, 'r+');
            fs.closeSync(fd);
            log += "File Lock: NO (Accessible)\n";
        } catch(e) {
            log += "File Lock: YES (Busy or No Permission)\n";
        }
    } else {
        log += `File MISSING: ${xlsxPath}\n`;
    }

    fs.writeFileSync('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\final_diagnostic.log', log);
    console.log(log);
}

check();
