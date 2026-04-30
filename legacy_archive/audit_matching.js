const fs = require('fs');
const csv = fs.readFileSync('_MPS_Final_Data_v3.csv', 'utf8');
const lines = csv.split('\n');

const results = {
    LYNX: { total: 0, matched: 0, unmapped: 0, stolenByOther: 0 },
    PUMA: { total: 0, matched: 0, unmapped: 0 },
    DVF: { total: 0, matched: 0, unmapped: 0 },
    'DNM/DNX': { total: 0, matched: 0, unmapped: 0 },
    VCF: { total: 0, matched: 0, unmapped: 0 },
    HORIZ: { total: 0, matched: 0, unmapped: 0 },
    OTHERS: { total: 0, matched: 0, unmapped: 0 }
};

function getCategory(name) {
    if (!name) return "ETC";
    const n = name.toUpperCase().trim();

    if (n.startsWith('VCF') || n.includes(' VCF') || n.startsWith('VF') || n.includes(' VF')) return 'VCF';
    if (n.startsWith('DVF') || n.includes(' DVF')) return 'DVF';
    if (n.startsWith('TT') || n.includes(' TT') || n.startsWith('TL') || n.includes(' TL')) return 'T-LATHE';
    if ((n.startsWith('M') && /^[0-9]/.test(n.substring(1))) || n.startsWith('MYNX') || n.includes(' DNM') || n.includes(' VC')) return 'DNM/DNX';
    if ((n.startsWith('V') && /^[0-9]/.test(n.substring(1))) || n.startsWith('PV') || n.startsWith('VT') || n.startsWith('VTR') || n.startsWith('VAW') || n.includes(' V-')) return 'V-LATHE';
    if (n.startsWith('T') && /^[0-9]/.test(n.substring(1))) return 'T-LATHE';
    if (n.startsWith('GT') || n.includes(' GT')) return 'PUMA-GT'; 
    if (n.startsWith('ST') || n.includes(' ST')) return 'LYNX';    
    if (n.startsWith('DNM') || n.startsWith('DNX') || n.startsWith('DNC') || n.includes(' DNM') || n.includes(' DNX')) return 'DNM/DNX';
    if (n.startsWith('HM') || n.startsWith('NHP') || n.startsWith('NHM') || n.startsWith('HC') || n.startsWith('HP') || n.includes(' NHP') || n.includes(' NHM')) return 'HORIZ';
    if (n.startsWith('LYNX') || (n.startsWith('L') && /^[0-9]/.test(n.substring(1)))) return 'LYNX';
    if (n.startsWith('PUMA') || (n.startsWith('P') && /^[0-9]/.test(n.substring(1)))) return 'PUMA';
    const pure = n.replace(/^[^0-9]+/, '');
    if (pure.startsWith('26') || pure.startsWith('31') || pure.startsWith('41') || pure.startsWith('51') || pure.startsWith('60') || pure.startsWith('70') || pure.startsWith('80') || pure.startsWith('1000')) return 'PUMA';
    if (pure.startsWith('21') || pure.startsWith('20') || pure.startsWith('16')) return 'LYNX';
    if (n.startsWith('ML') || n.startsWith('LS') || n.startsWith('MS') || n.includes('LYNX') || n.includes('SMX')) return 'LYNX';
    if (n.includes('PUMA') || n.startsWith('P')) return 'PUMA';
    return "OTHERS";
}

function norm(s) {
    if (!s) return "";
    let res = s.toString().toUpperCase().trim();
    if (res.startsWith('DCM')) res = 'DC' + res.substring(3);
    if (res.startsWith('PUMA')) res = 'P' + res.substring(4);
    if (res.startsWith('LYNX')) res = res.substring(4);
    if (res.startsWith('VCF')) res = 'VF' + res.substring(3);

    // [ADD] 숫자 시리즈 압축 (P2600 -> P26)
    res = res.replace(/([A-Z])([0-9]{2})00/g, '$1$2');

    if (res.startsWith('P51')) res = res.replace(/^P510+/, 'P51');
    return res.replace(/[^A-Z0-9]/g, '');
}


lines.slice(1).forEach(line => {
    if (!line.trim()) return;
    const parts = line.split(',').map(p => p.trim().replace(/^"|"$/g, ''));
    if (parts.length < 7) return;
    
    const model = parts[2];
    const product = parts[6];
    const cat = getCategory(model);
    
    if (!results[cat]) results[cat] = { total: 0, matched: 0, unmapped: 0 };
    results[cat].total++;
    
    if (product === 'UNMAPPED' || !product) {
        results[cat].unmapped++;
    } else {
        results[cat].matched++;
        const prodCat = getCategory(product);
        if (prodCat !== cat) {
            if (!results[cat].stolenFrom) results[cat].stolenFrom = {};
            results[cat].stolenFrom[prodCat] = (results[cat].stolenFrom[prodCat] || 0) + 1;
        }
    }
});

console.log('--- Matching Quality Audit Results ---');
console.table(results);
fs.writeFileSync('audit_results.json', JSON.stringify(results, null, 2));
