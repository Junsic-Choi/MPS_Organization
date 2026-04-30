const fs = require('fs');
const extractor = require('./extractor.js');

async function test() {
    try {
        const buf = fs.readFileSync('MPS2603-1.xlsx');
        const verCodeLookupPath = 'ver_code_lookup.json';
        const siteMasterPath = 'site_master.json';
        
        let verCodeLookup = {};
        let siteMaster = {};
        if (fs.existsSync(verCodeLookupPath)) verCodeLookup = JSON.parse(fs.readFileSync(verCodeLookupPath));
        if (fs.existsSync(siteMasterPath)) siteMaster = JSON.parse(fs.readFileSync(siteMasterPath));

        console.log("Analyzing file...");
        const result = await extractor.processMpsFile(buf, siteMaster, null, verCodeLookup);
        console.log("Success! Items:", result.length);
    } catch (e) {
        console.error("Error occurred:", e);
    }
}
test();
