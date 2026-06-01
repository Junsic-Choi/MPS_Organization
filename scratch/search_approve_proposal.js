const fs = require('fs');

const html = fs.readFileSync('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\dashboard.html', 'utf8');
const lines = html.split('\n');

lines.forEach((line, idx) => {
    if (line.includes('approveProposal') || line.includes('prop-select') || line.includes('proposal') || line.includes('unmapped-')) {
        console.log(`Line ${idx+1}: ${line.trim()}`);
    }
});
