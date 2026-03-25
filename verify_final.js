const fs = require('fs');
const file = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\_FinalList_4650_FIXED.csv';
try {
    const data = fs.readFileSync(file, 'utf8');
    const lines = data.split('\n');
    console.log('Total Rows:', lines.length - 1);
    console.log('DHF Count:', lines.filter(l => l.includes('DHF')).length);
    console.log('DHF First Match:', lines.filter(l => l.includes('DHF'))[0]);
    console.log('DBD First Match:', lines.filter(l => l.includes('DBD'))[0]);
    console.log('HFP First Match:', lines.filter(l => l.includes('HFP'))[0]);
} catch (e) {
    console.error('Verify error:', e.message);
}
