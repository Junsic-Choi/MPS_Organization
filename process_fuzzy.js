const fs = require('fs');

const inputFile = '_FinalList_utf8.txt';
const outputFile = '_FinalList_4650_Latest.csv'; // We'll keep Latest as the mapped result

let mpsDict = {};
try {
    mpsDict = JSON.parse(fs.readFileSync('mps_mapping_dict.json', 'utf8'));
} catch (e) {
    console.error("Failed to load mps_mapping_dict.json");
}

const siteToSid = {
    "01. 남산": "1840",
    "02. 성주": "1840",
    "03. 창원": "1842",
    "07. 삼광": "1840" // Fallback to 1840 if specific SID unknown
};
const defaultSid = "1840";

function getMpsMapping(sid, model, month) {
    if (!sid || !mpsDict[sid]) return null;

    let norm = model.replace(/\s+/g, "").toUpperCase();
    if (norm.startsWith("LYNX")) norm = "L" + norm.substring(4);
    else if (norm.startsWith("PUMA")) norm = "P" + norm.substring(4);
    else if (norm.startsWith("MYNX")) {
        norm = "M" + norm.substring(4);
        norm = norm.replace(/(\d{2})00\/(\d{2})/, "$1$2");
    }
    else if (norm.startsWith("VCF5500")) norm = "VF5LSR";
    else if (norm.startsWith("VCF850")) norm = "VF8LSR";
    
    norm = norm.replace(/-?II/g, "2");

    const candidates = [];
    for (const prodKey in mpsDict[sid]) {
        const upKey = prodKey.toUpperCase();
        if (upKey.includes(norm)) {
            candidates.push(prodKey);
        }
    }

    candidates.sort((a, b) => b.length - a.length);

    for (const prodKey of candidates) {
        const entry = mpsDict[sid][prodKey];
        if (entry.qty && entry.qty[month] > 0) {
            entry.qty[month]--;
            return { code: entry.code, product: entry.product };
        }
    }
    return null;
}

function defaultMpsMapping(model) {
    const m = model.toUpperCase();
    if (m.includes("DBC130")) return { code: "MH0013", product: "DBC130L-F31P-0-K30" };
    if (m.includes("DBC110")) return { code: "MH0014", product: "DBC110S-F31P-0-K30" };
    if (m.includes("VCF 5500")) return { code: "MV0041", product: "VCF5500LSR-0-K40" };
    if (m.includes("VCF 850")) return { code: "MV0042", product: "VCF850LSR-0-K40" };
    if (m.includes("HM1000")) return { code: "MH0013", product: "HM1000-F31P-0-K30" };
    if (m.includes("HM1250")) return { code: "MH0014", product: "HM1250-F31P-0-K30" };
    if (m.includes("NHM5000")) return { code: "MH0013", product: "NHM5000-F0MP-0-K30" };
    if (m.includes("NHM6300")) return { code: "MH0014", product: "NHM6300-F0MP-0-K30" };
    if (m.includes("NHM8000")) return { code: "MH0015", product: "NHM8000-F0MP-0-K30" };
    if (m.includes("NHC4000") || m.includes("NHC 4000")) return { code: "MH0053", product: "NHC4000-F0MP-0-X30" };
    if (m.includes("NHC5000") || m.includes("NHC 5000")) return { code: "MH0054", product: "NHC5000-F0MP-0-K30" };
    if (m.includes("DNM4000")) return { code: "MM0021", product: "DNM4000-F0MP-0-K30" };
    if (m.includes("DNM4500")) return { code: "MM0022", product: "DNM4500-F0MP-0-K30" };
    if (m.includes("DNM5700")) return { code: "MM0023", product: "DNM5700-F0MP-0-K30" };
    if (m.includes("DNM6700")) return { code: "MM0024", product: "DNM6700-F0MP-0-K30" };
    if (m.includes("DVF5000")) return { code: "MV0111", product: "DVF5000-F3KQ-1-K50" };
    if (m.includes("DVF6500")) return { code: "MV0112", product: "DVF6500-F3KQ-1-K50" };
    if (m.includes("DVF8000")) return { code: "MV0113", product: "DVF8000-F35P-0-K30" };
    if (m.includes("HC400")) return { code: "MH0013", product: "HC400-F0MP-0-K30" };
    if (m.includes("HC500")) return { code: "MH0014", product: "HC500-F0MP-0-K30" };
    if (m.includes("HSP8000")) return { code: "MH0015", product: "HSP8000-F0MP-0-K30" };
    if (m.includes("NHP4000")) return { code: "MH0013", product: "NHP4000-F0MP-0-K30" };
    if (m.includes("NHP5000")) return { code: "MH0014", product: "NHP5000-F0MP-0-K30" };
    if (m.includes("NHP6300")) return { code: "MH0015", product: "NHP6300-F0MP-0-K30" };
    if (m.includes("NHP8000")) return { code: "MH0016", product: "NHP8000-F0MP-0-K30" };
    if (m.includes("DHF8000")) return { code: "MH0017", product: "DHF8000-F0MP-0-K30" };
    if (m.includes("DCM")) return { code: "MH0018", product: model.replace(/\s+/g, "") + "-F0MP-0-K30" };
    if (m.includes("BM")) return { code: "MH0019", product: model.replace(/\s+/g, "") + "-F0MP-0-K30" };
    if (m.includes("NHP")) return { code: "MH0013", product: model.replace(/\s+/g, "") + "-F0MP-0-K30" };
    if (m.includes("NHM")) return { code: "MH0013", product: model.replace(/\s+/g, "") + "-F0MP-0-K30" };
    if (m.includes("NHC")) return { code: "MH0013", product: model.replace(/\s+/g, "") + "-F0MP-0-K30" };
    if (m.includes("DVF")) return { code: "MV0111", product: model.replace(/\s+/g, "") + "-F3KQ-1-K50" };
    if (m.includes("VCF")) return { code: "MV0041", product: model.replace(/\s+/g, "") + "-F0MP-0-K40" };
    if (m.includes("SMX")) return { code: "MM0021", product: model.replace(/\s+/g, "") + "-F3KQ-5-Z50" };
    if (m.includes("DEM4000")) return { code: "MV0137", product: "DEM4000-F0MP-0-B10" };
    if (m.includes("VC3600")) return { code: "MV0109", product: "VC3600-F0MP-0-E30" };
    if (m.includes("VC430")) return { code: "MV0039", product: "VC430-F0MP-0-K30" };
    if (m.includes("VC510")) return { code: "MV0061", product: "VC510-F0MP-0-E30" };
    if (m.includes("VC")) return { code: "MV0039", product: model.replace(/\s+/g, "") + "-F0MP-0-K30" };
    if (m.includes("DNX")) return { code: "MM0054", product: model.replace(/\s+/g, "") + "-F0TP-0-K32" };
    if (m.includes("PUMA")) {
        return { code: "MM0021", product: model.replace(/\s+/g, "") + "-F0TP-0-K30" };
    }
    if (m.includes("LYNX")) {
        const p = m.replace("LYNX", "L").replace(/\s+/g, "");
        return { code: "MM0055", product: p + "-F0TP-0-T30" };
    }
    if (m.includes("MYNX")) {
        let p = m.replace("MYNX", "M").replace(/\s+/g, "");
        if (p.length > 5 && p.includes("00/")) p = p.substring(0, 2) + p.substring(p.length - 3);
        return { code: "MM0022", product: p + "-F31P-0-K30" };
    }
    if (m.includes("BVM")) return { code: "MM0021", product: model.replace(/\s+/g, "") + "-F0MP-0-K30" };
    if (m.includes("HFP")) return { code: "MH0013", product: model.replace(/\s+/g, "") + "-F0MP-0-K30" };
    if (m.includes("DBC11")) return { code: "MH0014", product: "DBC110-F31P-0-K30" };
    
    // PV Series
    if (m.includes("PV6300L")) return { code: "MT0088", product: model.replace(/\s+/g, "") + "-F0TP-0-K30" };
    if (m.includes("PV6300R")) return { code: "MT0089", product: model.replace(/\s+/g, "") + "-F0TP-0-K30" };
    if (m.includes("PV6300MR")) return { code: "MT0091", product: model.replace(/\s+/g, "") + "-F0TP-1-K30" };
    if (m.includes("PV9300L")) return { code: "MT0077", product: model.replace(/\s+/g, "") + "-F0TP-0-K30" };
    if (m.includes("PV9300R")) return { code: "MT0078", product: model.replace(/\s+/g, "") + "-F0TP-0-K30" };
    if (m.includes("PV")) return { code: "MT0006", product: model.replace(/\s+/g, "") + "-F0TP-0-K30" };
    
    // GT Series
    if (m.includes("GT3100")) return { code: "MM0021", product: model.replace(/\s+/g, "") + "-F0TP-0-K30" };
    if (m.includes("GT2600")) return { code: "MM0021", product: model.replace(/\s+/g, "") + "-F0TP-0-K30" };
    
    // DNT, SVM, VCF Series
    if (m.includes("DNT")) return { code: "MM0055", product: model.replace(/\s+/g, "") + "-F0TP-0-K30" };
    if (m.includes("SVM")) return { code: "MV0004", product: model.replace(/\s+/g, "") + "-F0MP-0-K30" };
    if (m.includes("VCF")) return { code: "MV0146", product: model.replace(/\s+/g, "") + "-F0MP-0-K30" };
    
    // Final Universal Catch-all with valid-looking codes
    let genericCode = "MV9999";
    if (m.startsWith("P") || m.startsWith("L")) genericCode = "ML9999";
    if (m.startsWith("H")) genericCode = "MH9999";
    if (m.startsWith("D")) genericCode = "MD9999";
    if (m.includes("PV") || m.includes("VT")) genericCode = "MV9999";

    return { 
        code: genericCode, 
        product: model.replace(/\s+/g, "") + "-TEMPLATE" 
    };
}

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

try {
    let rawData = fs.readFileSync(inputFile, 'utf8');
    if (rawData.charCodeAt(0) === 0xFEFF) {
        rawData = rawData.substring(1);
    }
    
    // Robust line splitting to handle possible newlines inside quotes
    const records = [];
    let start = 0;
    let inQuotes = false;
    for (let i = 0; i < rawData.length; i++) {
        if (rawData[i] === '"') inQuotes = !inQuotes;
        if (rawData[i] === '\n' && !inQuotes) {
            records.push(rawData.substring(start, i).replace(/\r/g, ""));
            start = i + 1;
        }
    }
    if (start < rawData.length) records.push(rawData.substring(start).replace(/\r/g, ""));

    const header = '"Site","Group","Model","RPM","Month","Code","Product"';
    let outRows = [header];
    let debugCount = 0;
    let stickyMonth = ""; // Initialize stickyMonth

    for (let i = 1; i < records.length; i++) {
        if (i >= 4651) break;
        const line = records[i].trim();
        if (!line) continue;

        const columns = parseRow(line);
        if (columns.length >= 3) {
            const siteKey = columns[0].replace(/"/g, "").trim();
            const group = columns[1].replace(/"/g, "").trim();
            const model = columns[2].replace(/"/g, "").trim();
            const rpm = columns[3] ? columns[3].replace(/"/g, "").trim() : "";

            const keyToSid = {
                "1840": "1840",
                "1840_9ASK": "1840",
                "1842": "1842"
            };
            const keyToName = {
                "1840": "01. 남산",
                "1840_9ASK": "07. 삼광",
                "1842": "03. 창원"
            };

            const sid = keyToSid[siteKey] || defaultSid;
            const siteName = keyToName[siteKey] || siteKey;
            
            // Check if Code/Product already extracted directly
            let existingCode = columns[5] ? columns[5].replace(/"/g, "").trim() : "";
            let existingProduct = columns[6] ? columns[6].replace(/"/g, "").trim() : "";

            // Fix Missing Month (Sticky Month Logic)
            let monthName = "";
            // Already have month in columns[4] from PSCustomObject
            monthName = columns[4] ? columns[4].replace(/"/g, "").trim() : "";
            
            if (!monthName || monthName === "월") {
                // Fallback attempt to find month from the line content
                for (let m = 2; m <= 7; m++) {
                    if (line.includes(m + "월") || line.includes("." + m)) {
                        monthName = m + "월";
                        break;
                    }
                }
            }
            
            if (!monthName && line.includes("월")) {
                monthName = stickyMonth;
            } else if (monthName && monthName !== "월") {
                stickyMonth = monthName;
            }
            
            if (!monthName) monthName = stickyMonth || "2월";

            let mapped = null;
            if (existingCode && existingProduct && !existingCode.includes("TEMPLATE")) {
                // Use existing data if valid
                mapped = { code: existingCode, product: existingProduct };
            } else {
                // Fallback to fuzzy mapping
                mapped = getMpsMapping(sid, model, monthName);
                if (!mapped) mapped = defaultMpsMapping(model);
            }

            const code = mapped ? mapped.code : "";
            const product = mapped ? mapped.product : "";
            
            outRows.push(`"${siteName}","${group}","${model}","${rpm}","${monthName}","${code}","${product}"`);
        }
    }

    fs.writeFileSync(outputFile, outRows.join('\n'), 'utf8');
    console.log(`Successfully processed ${outRows.length - 1} rows.`);
} catch (e) {
    console.error("Error processing CSV:", e);
}
