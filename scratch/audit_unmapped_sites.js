const XLSX = require('xlsx');
const { processMpsFile } = require('../extractor');
const fs = require('fs');

async function auditUnmapped() {
    console.log('--- Auditing Unmapped Items by Site ---');
    const buffer = fs.readFileSync('MPS2605-1.xlsx');
    // Basic site master rules from dashboard.html
    const rules = {
        siteMaster: {
            "I0215116": "07. 삼광", "I0205716": "09. 서진", "I0206873": "성주", "I0206954": "성주",
            "I0215001": "세양", "I0212077": "04. 성우", "I0213836": "04. 성우", "I0213835": "06. 원진",
            "I0206330": "05. 세양", "I0206329": "05. 세양", "I0206328": "05. 세양", "I0205562": "15. 신우",
            "I0205561": "15. 신우", "I0205560": "15. 신우", "I0206254": "11. 대영", "I0206253": "11. 대영",
            "9AHT": "21. 휴텍", "1840": "남산", "1842": "성주", "9ASW": "04. 성우"
        }
    };

    const result = await processMpsFile(buffer, rules);
    const unused = result.unusedData;
    
    const stats = {};
    unused.forEach(row => {
        const site = row.Site || 'Unknown';
        if (!stats[site]) stats[site] = { count: 0, models: new Set() };
        stats[site].count += (parseInt(row.Qty) || 0);
        stats[site].models.add(row.UnmappedModel || row.Model);
    });

    console.log('\n[Unmapped Stats per Site]');
    for (const site in stats) {
        console.log(`${site}: ${stats[site].count} qty (${stats[site].models.size} distinct models)`);
        console.log(`  Samples: ${Array.from(stats[site].models).slice(0, 5).join(', ')}`);
    }
}

auditUnmapped();
