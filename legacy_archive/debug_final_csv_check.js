const fs = require('fs');

const content = fs.readFileSync('_MPS_Final_Data_v3.csv', 'utf8');
const lines = content.split('\n');
let out = "--- CSV Final Verification ---\n";
lines.forEach(line => {
    if (line.includes('XG800') && line.includes('휴텍')) {
        out += line + "\n";
    }
});
fs.writeFileSync('final_csv_proof.txt', out);
console.log('Done.');
