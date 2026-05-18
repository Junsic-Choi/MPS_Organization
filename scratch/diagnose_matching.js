const XLSX = require('xlsx');
const { processMpsFile } = require('../extractor');
const fs = require('fs');

async function diagnoseMatching() {
    const buffer = fs.readFileSync('MPS2605-1.xlsx');
    const rules = {
        siteMaster: {
            "I0215116": "07. 삼광", "I0205716": "09. 서진", "I0206873": "성주", "I0206954": "성주",
            "I0215001": "LEO", "I0212077": "04. 성우", "I0213836": "04. 성우", "I0213835": "06. 원진",
            "I0206330": "05. 세양", "I0206329": "05. 세양", "I0206328": "05. 세양", "I0205562": "15. 신우",
            "I0205561": "15. 신우", "I0205560": "15. 신우", "I0206254": "11. 대영", "I0206253": "11. 대영",
            "9AHT": "21. 휴텍", "1840": "남산", "1842": "성주", "9ASW": "04. 성우",
            "9ACE": "지티테크", "I0169394": "지티테크"
        }
    };

    console.log('--- Diagnosing VTR and DBM Matching ---');
    const result = await processMpsFile(buffer, rules);
    
    const vtrMatches = result.finalResults.filter(r => r.Site === '성주' && r.Month === '6월' && r.Model.includes('VTR'));
    console.log(`VTR Matches (성주, 6월): ${vtrMatches.length}`);
    vtrMatches.forEach(m => console.log(`  Match: ${m.Model} -> ${m.Product}`));

    const vtrUnused = result.unusedData.filter(r => r.Site === '성주' && r.Month === '6월' && r.Model.includes('VTR'));
    console.log(`VTR Unmapped (성주, 6월): ${vtrUnused.length}`);
    vtrUnused.forEach(u => console.log(`  Unmapped: ${u.UnmappedModel} (Qty: ${u.Qty})`));

    const dbmUnused = result.unusedData.filter(r => r.Site === '성주' && r.Month === '9월' && r.Model.includes('DBM'));
    console.log(`DBM Unmapped (성주, 9월): ${dbmUnused.length}`);
    dbmUnused.forEach(u => console.log(`  Unmapped: ${u.UnmappedModel} (Qty: ${u.Qty})`));
}

diagnoseMatching();
