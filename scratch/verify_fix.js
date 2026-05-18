const { processMpsFile } = require('../extractor');

async function run() {
    const filename = 'MPS2605-1.xlsx';
    console.log(`Testing ${filename}...`);
    // Pass the siteMaster rules from dashboard.html to match the user's environment
    const rules = {
        siteMaster: {
            "I0215116": "07. 삼광",
            "I0205716": "09. 서진",
            "I0206873": "성주",
            "I0206954": "성주",
            "I0215001": "세양",
            "I0212077": "성우",
            "9AHT": "21. 휴텍",
            "1840": "남산",
            "1842": "성주"
        }
    };
    try {
        const result = await processMpsFile(filename, rules);
        console.log('Total Matched:', result.finalResults.length);
        console.log('Total Unmatched:', result.unusedData.length);
        if (result.finalResults.length > 0) {
            console.log('Sample Match:', result.finalResults[0]);
        }
    } catch (err) {
        console.error(err);
    }
}

run();
