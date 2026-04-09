const fs = require('fs');
const readline = require('readline');

async function checkMapping() {
    const fileStream = fs.createReadStream('_MPS_Final_Data_v3.csv');
    const rl = readline.createInterface({
        input: fileStream,
        crlfDelay: Infinity
    });

    console.log('--- Current Mapping Results for Hutec / LYNX XG ---');
    for await (const line of rl) {
        if (line.includes('휴텍') || line.includes('LYNX XG')) {
            console.log(line);
        }
    }
}

checkMapping();
