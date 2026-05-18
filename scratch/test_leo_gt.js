const XLSX = require('xlsx');
const { processMpsFile } = require('../extractor');
const fs = require('fs');

async function testLeoGt() {
    const buffer = fs.readFileSync('MPS2605-1.xlsx');
    // Using the same rules as the dashboard would use
    const rules = {
        siteMaster: {
            "I0215001": "LEO",
            "9ACE": "지티테크",
            "I0169394": "지티테크"
        }
    };

    console.log('--- Testing LEO and G-Tech Matching ---');
    const result = await processMpsFile(buffer, rules);
    
    const leoResults = result.finalResults.filter(r => r.Site === 'LEO' && r.Model.includes('LEO'));
    console.log(`LEO Matches: ${leoResults.length}`);
    leoResults.forEach(m => console.log(`  Match: ${m.Month} ${m.Model} -> ${m.Product} (Qty: ${m.Qty})`));

    const gtResults = result.finalResults.filter(r => r.Site === '지티테크');
    console.log(`G-Tech Matches: ${gtResults.length}`);
    gtResults.forEach(m => console.log(`  Match: ${m.Month} ${m.Model} -> ${m.Product} (Qty: ${m.Qty})`));

    const unmapped = result.unusedData.filter(r => r.Site === 'LEO' || r.Site === '지티테크');
    console.log(`Remaining Unmapped (LEO/GT): ${unmapped.length}`);
    unmapped.forEach(u => console.log(`  Unmapped: ${u.Month} ${u.Site} ${u.Model} (Qty: ${u.Qty})`));
}

testLeoGt();
