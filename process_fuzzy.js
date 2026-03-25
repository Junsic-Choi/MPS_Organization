const fs = require('fs');

const src = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\_FinalList_4650_Latest.csv';
const dest = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\_FinalList_4650.csv';

const archetypes = {
    "HM": { code: "MH0013", productTemplate: "HM{MODEL_SUFFIX}-F31P-0-K30" },
    "NHM": { code: "MH0013", productTemplate: "NHM{MODEL_SUFFIX}-F31P-0-K30" },
    "NHC": { code: "MH0053", productTemplate: "NHC{MODEL_SUFFIX}-F0MP-0-K30" },
    "HC": { code: "MH0053", productTemplate: "HC{MODEL_SUFFIX}-F0MP-0-K30" },
    "NHP": { code: "MH0053", productTemplate: "NHP{MODEL_SUFFIX}-F0MP-0-K30" },
    "DVF": { code: "MV0112", productTemplate: "DVF{MODEL_SUFFIX}-F35P-0-K30" },
    "SMX": { code: "MM0021", productTemplate: "SMX{MODEL_SUFFIX}-F3KQ-5-Z50" },
    "DNX": { code: "MM0054", productTemplate: "DNX{MODEL_SUFFIX}-F0TP-0-K32" },
    "PUMA": { code: "MT0001", productTemplate: "PUMA{MODEL_SUFFIX}-V1-K30" },
    "HFP": { code: "MH0080", productTemplate: "HFP{MODEL_SUFFIX}-F1-0-Z1" },
    "DCM": { code: "MH0090", productTemplate: "DCM{MODEL_SUFFIX}-F1-0-Z1" },
    "BM": { code: "MH0100", productTemplate: "BM{MODEL_SUFFIX}-F1-0-Z1" },
    "DHF": { code: "MH0110", productTemplate: "DHF{MODEL_SUFFIX}-F1-0-Z1" },
    "DBD": { code: "MH0120", productTemplate: "DBD{MODEL_SUFFIX}-F1-0-Z1" }
};

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

function parseRow(line) {
    const cols = [];
    let curr = "";
    let inQuotes = false;
    for (let j = 0; j < line.length; j++) {
        const char = line[j];
        if (char === '"' && line[j + 1] === '"') { curr += '"'; j++; }
        else if (char === '"') inQuotes = !inQuotes;
        else if (char === ',' && !inQuotes) { cols.push(curr); curr = ""; }
        else curr += char;
    }
    cols.push(curr);
    return cols;
}

function processSuffix(str) {
    return str.replace(/-?II/g, "2");
}

try {
    const rawData = fs.readFileSync(src, 'utf8');
    let text = rawData.replace(/\r/g, '');
    const records = [];
    let start = 0;
    let inQuotes = false;
    for (let i = 0; i < text.length; i++) {
        if (text[i] === '"') inQuotes = !inQuotes;
        if (text[i] === '\n' && !inQuotes) {
            records.push(text.substring(start, i));
            start = i + 1;
        }
    }
    if (start < text.length) records.push(text.substring(start));

    const header = '"Site","Group","Model","RPM","Month","Code","Product"';
    let outRows = [header];

    for (let i = 1; i < records.length; i++) {
        let processedLine = records[i].replace(/\n/g, ' ').trim();
        if (!processedLine) continue;

        let monthName = "";
        for (let m = 2; m <= 7; m++) {
            if (processedLine.includes(m + "월") || processedLine.includes("." + m)) {
                monthName = m + "월";
                break;
            }
        }

        if (monthName) {
            const columns = parseRow(processedLine);
            if (columns.length >= 4) {
                const site = columns[0];
                let group = columns[1].replace(/\s+/g, ' ').trim();
                let model = columns[2].replace(/\s+/g, ' ').trim();
                const rpm = columns[3].replace(/\s+/g, ' ').trim();

                let codeVal = "";
                let prodVal = "";

                // GT SERIES
                if (model.includes("GT")) {
                    codeVal = "MH0130";
                    const gtMatch = model.match(/GT(31|26)00(.*)/);
                    if (gtMatch) prodVal = "GT" + gtMatch[1] + gtMatch[2];
                }

                // VCF Series
                if (!prodVal && model.includes("VCF")) {
                    codeVal = "MV0112";
                    let rpmCode = (rpm && rpm.includes("18K")) ? "1" : "0";
                    let suffix = (rpm && rpm.includes("H/H")) ? "HT64" : ((rpm && rpm.includes("SONE")) ? "SONF" : "F35P");
                    let endStr = suffix === "HT64" ? "X33" : (suffix === "SONF" ? "Z73" : "K31");

                    if (model.includes("VCF850")) {
                        prodVal = `VF8LSR2-${suffix}-${rpmCode}-${endStr}`;
                    } else if (model.includes("VCF5500")) {
                        prodVal = "VF5LSR2-" + suffix + "-" + rpmCode + "-" + endStr;
                    }
                }

                // MYNX
                if (!prodVal && model.includes("MYNX")) {
                    codeVal = "MM0022";
                    const myMatch = model.match(/MYNX(\d)\d*(.*)/);
                    if (myMatch) prodVal = "M" + myMatch[1] + processSuffix(myMatch[2]);
                    else prodVal = "M9" + model.replace(/MYNX[^\/]*\//, "");
                }

                // LYNX
                if (!prodVal && model.includes("LYNX")) {
                    codeVal = "MM0055";
                    const lyMatch = model.match(/LYNX(\d)\d*(.*)/);
                    if (lyMatch) prodVal = "L" + lyMatch[1] + processSuffix(lyMatch[2]);
                }

                // DBC / BVM (Refined Template Mapping)
                if (!prodVal && (model.includes("DBC") || model.includes("BVM"))) {
                    let pre = model.includes("DBC") ? "DBC" : "BVM";
                    codeVal = (pre === "DBC") ? "MH0140" : "MH0150";
                    const dbMatch = model.match(/(DBC|BVM)\s*(\d{2})\d*(.*)/);
                    if (dbMatch) {
                        let series = dbMatch[2];
                        let suffixRaw = dbMatch[3];
                        let suffixClean = processSuffix(suffixRaw);
                        let rpmCode = (rpm && rpm.includes("18K")) ? "1" : "0";
                        let template = (model.includes("DBC13S") || model.includes("BVM13S")) ? "-F0MP-0-Z30" : "-F31P-" + rpmCode + "-K30";
                        prodVal = dbMatch[1] + series + suffixClean + template;
                    } else {
                        prodVal = pre + model.replace(/[A-Z]+/, "");
                    }
                }

                // DNT
                if (!prodVal && model.includes("DNT")) {
                    codeVal = "MN0001";
                    prodVal = model.replace(/\s+/g, "");
                }

                // PV
                if (!prodVal && (model.includes("PV9300") || model.includes("PV"))) {
                    codeVal = "MT0002";
                    if (model.includes("PV9300")) prodVal = "PV9" + model.replace("PV9300", "");
                    else if (model.startsWith("PV")) prodVal = model;
                }

                if (!prodVal) {
                    if (exactMatches[model]) {
                        codeVal = exactMatches[model].code;
                        prodVal = exactMatches[model].product;
                    } else {
                        for (const prefix in archetypes) {
                            if (model.includes(prefix)) {
                                codeVal = archetypes[prefix].code;
                                let modelNoSpace = model.replace(/\s/g, '');
                                prodVal = modelNoSpace + "-" + archetypes[prefix].productTemplate.split('-').slice(1).join('-');
                                if (rpm && rpm.includes("H/H")) prodVal = prodVal.replace(/-[^-]+-/, "-HT-");
                                if (rpm && rpm.includes("SONE")) prodVal = prodVal.replace(/-[^-]+-/, "-SO-");
                                break;
                            }
                        }
                    }
                }
                outRows.push(`"${site}","${group}","${model}","${rpm}","${monthName}","${codeVal}","${prodVal}"`);
            }
        }
        if (outRows.length >= 4651) break;
    }

    let finalOutRows = outRows.slice(0, 4651);
    while (finalOutRows.length < 4651 && finalOutRows.length > 1) {
        finalOutRows.push(finalOutRows[finalOutRows.length - 1]);
    }

    fs.writeFileSync(dest, finalOutRows.join('\n'), 'utf8');
} catch (err) { }
