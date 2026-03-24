const fs = require('fs');
const jsonSite = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\site_data.json';
const outMap = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\site_map_all.json';

try {
    const raw = fs.readFileSync(jsonSite);
    const data = raw.toString('utf8');
    // If it's not valid JSON, try to clean it
    let json = [];
    try {
        json = JSON.parse(data.replace(/^\uFEFF/, ''));
    } catch (e) {
        // Fallback: search for patterns if JSON fails
        const matches = data.match(/"Prod. Ver":"([^"]+)","Prod. Ver Description":"([^"]+)"/g);
        if (matches) {
            matches.forEach(m => {
                const sub = m.match(/"Prod. Ver":"([^"]+)","Prod. Ver Description":"([^"]+)"/);
                if (sub) json.push({ "Prod. Ver": sub[1], "Prod. Ver Description": sub[2] });
            });
        }
    }

    const map = {};
    json.forEach(item => {
        const d = item["Prod. Ver Description"];
        const v = item["Prod. Ver"];
        if (d && v) map[d] = v;
    });

    fs.writeFileSync(outMap, JSON.stringify(map, null, 2), 'utf8');
    console.log(`Mapped ${Object.keys(map).length} items from site data.`);
} catch (err) {
    console.error("Error:", err.message);
}
