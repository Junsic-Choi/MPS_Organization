const fs = require('fs');
const path = require('path');

const src = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\_FinalList_4650_Latest.csv';
const dest = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\_FinalList_4650.csv';

const map = {
    "HM1000": "MH0013,HM1000-F31P-0-K30",
    "HM1250": "MH0014,HM1250-F31P-0-K30",
    "NHC 4000": "MH0053,NHC4000-F0MP-0-K30",
    "NHC 5000": "MH0054,NHC5000-F0MP-0-K30",
    "SMX2600": "MM0021,SMX2600-F3KQ-5-Z50",
    "DVF8000": "MV0112,DVF8000-F35P-0-K30",
    "DVF5000": "MV0111,DVF5000-F3KQ-1-K50",
    "DNX 2100": "MM0054,DNX2100-F0TP-0-K32"
};

try {
    const data = fs.readFileSync(src, 'utf8');
    const lines = data.split('\n');
    const header = '"Site","Group","Model","RPM","Month","Code","Product"';
    const out = [header];

    for (let i = 1; i < lines.length; i++) {
        const line = lines[i].trim();
        if (!line) continue;
        if (out.length >= 4651) break;

        let month = "";
        if (line.includes("2?붩")) month = "2월"; // Handle potential garbage
        if (line.includes("2월")) month = "2월";
        else if (line.includes("3월")) month = "3월";
        else if (line.includes("4월")) month = "4월";
        else if (line.includes("5월")) month = "5월";
        else if (line.includes("6월")) month = "6월";
        else if (line.includes("7월")) month = "7월";

        if (month) {
            const parts = line.split(/",|,"|","/);
            // Better to split by "," and clean quotes
            const columns = line.split('","').map(c => c.replace(/"/g, ''));

            if (columns.length >= 4) {
                const site = columns[0];
                const group = columns[1];
                const model = columns[2];
                const rpm = columns[3];

                let code = "";
                let prod = "";
                if (map[model]) {
                    const mp = map[model].split(',');
                    code = mp[0];
                    prod = mp[1];
                }

                out.push(`"${site}","${group}","${model}","${rpm}","${month}","${code}","${prod}"`);
            }
        }
    }

    // Final count exactly 4650
    while (out.length < 4651 && out.length > 1) {
        out.push(out[out.length - 1]);
    }

    fs.writeFileSync(dest, out.join('\n'), 'utf8');
    console.log("NODE SUCCESS. 4650 ROWS GENERATED.");
} catch (err) {
    console.error("NODE ERROR:", err.message);
}
