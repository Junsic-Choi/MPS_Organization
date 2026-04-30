console.log('NODE_OK');
try {
    require('xlsx');
    console.log('XLSX_OK');
} catch (e) {
    console.log('XLSX_MISSING');
}
