const fs = require('fs');

const html = fs.readFileSync('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\dashboard.html', 'utf8');
const lines = html.split('\n');

lines.forEach((line, idx) => {
    if (line.includes('body-left') || line.includes('table-left') || line.includes('panel-left') || line.includes('left')) {
        // Let's print if it has javascript logic
        if (line.includes('document.getElementById') || line.includes('innerHTML') || line.includes('createElement')) {
            console.log(`Line ${idx+1}: ${line.trim()}`);
        }
    }
});
