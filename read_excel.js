const xlsx = require('xlsx');

try {
    const workbook = xlsx.readFile('MPS2603-1(생산배포용).xlsx');
    console.log('Sheets:', workbook.SheetNames);
    
    if (workbook.SheetNames.includes('[생산배포용]')) {
        const sheet = workbook.Sheets['[생산배포용]'];
        const json = xlsx.utils.sheet_to_json(sheet, { header: 1 });
        
        console.log('First 5 rows:');
        for (let i = 0; i < 5 && i < json.length; i++) {
            console.log(`Row ${i + 1}:`, json[i]);
        }
        
        // Find header row (usually the first or second row with actual text)
        let headerRow = null;
        for (let i = 0; i < Math.min(10, json.length); i++) {
            if (json[i] && json[i].filter(x => x).length > 5) {
                headerRow = json[i];
                console.log(`Potential Header Row (Index ${i}):`, headerRow);
                break;
            }
        }
    } else {
        console.log('Sheet [생산배포용] not found.');
    }
} catch (e) {
    console.error('Error reading excel file:', e);
}
