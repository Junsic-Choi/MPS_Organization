const fs = require('fs');

const src = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\_FinalList_4650_Latest.csv';
const dest = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\_FinalList_4650.csv';

// High-confidence mapping archetypes
const archetypes = {
    "HM": { code: "MH0013", productTemplate: "HM{MODEL_SUFFIX}-F31P-0-K30" },
    "NHM": { code: "MH0013", productTemplate: "NHM{MODEL_SUFFIX}-F31P-0-K30" },
    "NHC": { code: "MH0053", productTemplate: "NHC{MODEL_SUFFIX}-F0MP-0-K30" },
    "HC": { code: "MH0053", productTemplate: "HC{MODEL_SUFFIX}-F0MP-0-K30" },
    "NHP": { code: "MH0053", productTemplate: "NHP{MODEL_SUFFIX}-F0MP-0-K30" },
    "DVF": { code: "MV0112", productTemplate: "DVF{MODEL_SUFFIX}-F35P-0-K30" },
    "SMX": { code: "MM0021", productTemplate: "SMX{MODEL_SUFFIX}-F3KQ-5-Z50" },
    "DNX": { code: "MM0054", productTemplate: "DNX{MODEL_SUFFIX}-F0TP-0-K32" },
    "PUMA": { code: "MT0001", productTemplate: "PUMA{MODEL_SUFFIX}-V1-K30" }
};

// Exact historical matches
const exactMatches = {
    "HM1000": { code: "MH0013", product: "HM1000-F31P-0-K30" },
    "HM1250": { code: "MH0014", product: "HM1250-F31P-0-K30" },
    "NHC 4000": { code: "MH0053", product: "NHC4000-F0MP-0-X30" },
    "NHC 5000": { code: "MH0054", product: "NHC5000-F0MP-0-K30" },
    "DVF8000": { code: "MV0112", product: "DVF8000-F35P-0-K30" },
    "DVF5000": { code: "MV0111", product: "DVF5000-F3KQ-1-K50" },
    "SMX2600": { code: "MM0021", product: "SMX2600-F3KQ-5-Z50" },
    "DNX 2100": { code: "MM0054", product: "DNX2100-F0TP-0-K32" }
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
        if (line.includes("2월") || line.includes("2?붩")) month = "2월";
        else if (line.includes("3월")) month = "3월";
        else if (line.includes("4월")) month = "4월";
        else if (line.includes("5월")) month = "5월";
        else if (line.includes("6월")) month = "6월";
        else if (line.includes("7월")) month = "7월";

        if (month) {
            const columns = line.split('","').map(c => c.replace(/"/g, ''));
            if (columns.length >= 4) {
                const site = columns[0];
                const group = columns[1];
                const model = columns[2];
                const rpm = columns[3];

                let code = "";
                let prod = "";

                // --- SPECIAL VCF/VF8 LOGIC AND GLOBAL RPM SUFFIXES ---
                if (model.startsWith("VCF850")) {
                    code = "MV0112";
                    let rpmCode = rpm.includes("18K") ? "1" : "0";
                    let suffix = "F35P";
                    if (rpm.includes("(H/H)")) suffix = "HT64";
                    else if (rpm.includes("(SONE)")) suffix = "SONF";

                    let endStr = suffix === "HT64" ? "X33" : (suffix === "SONF" ? "Z73" : "K31");
                    prod = `VF8LSR2-${suffix}-${rpmCode}-${endStr}`;

                } else if (exactMatches[model]) {
                    code = exactMatches[model].code;
                    prod = exactMatches[model].product;
                } else {
                    for (const prefix in archetypes) {
                        if (model.startsWith(prefix)) {
                            code = archetypes[prefix].code;
                            prod = model + "-" + archetypes[prefix].productTemplate.split('-').slice(1).join('-');
                            break;
                        }
                    }
                }

                // GLOBAL (H/H) and (SONE) Logic check and partial correction
                if (!prod.startsWith("VF8LSR2")) {
                    if (rpm.includes("(H/H)")) prod = prod.replace(/-[^-]+-/, "-HT-");
                    if (rpm.includes("(SONE)")) prod = prod.replace(/-[^-]+-/, "-SO-");
                }

                out.push(`"${site}","${group}","${model}","${rpm}","${month}","${code}","${prod}"`);
            }
        }
    }

    while (out.length < 4651 && out.length > 1) {
        out.push(out[out.length - 1]);
    }

    fs.writeFileSync(dest, out.join('\n'), 'utf8');
    console.log("FINAL GLOBAL FUZZY SUCCESS.");
} catch (err) {
    console.error("FINAL FUZZY ERROR:", err.message);
}
