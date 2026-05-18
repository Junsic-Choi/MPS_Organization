function getMatchKey(s) {
    if (!s) return '';
    let n = s.toString().toUpperCase().trim();
    
    // [1. ROMAN TO NUMBER]
    n = n.replace(/VIII/g, '8').replace(/VII/g, '7').replace(/VI/g, '6').replace(/IV/g, '4').replace(/IX/g, '9');
    n = n.replace(/III/g, '3').replace(/II/g, '2');

    // [2. SPECIALIZED SERIES RULES]
    if (n.includes('SMX')) {
        n = n.replace(/SMX2(?![0-9])/g, 'SMX21');
        n = n.replace(/2100/g, '21').replace(/3100/g, '31').replace(/5100/g, '51');
        n = n.replace(/SYYB/g, 'SYY').replace(/STB/g, 'SB');
    }

    if (n.startsWith('VCF')) {
        n = n.replace('VCF', 'VF');
        n = n.replace(/(\d\d)0/, '$1');
    }

    if (n.startsWith('MYNX')) {
        let taperMatch = n.match(/6500\/(\d)0/);
        if (taperMatch) n = 'M65' + taperMatch[1];
        else n = 'M' + n.substring(4);
    }
    
    if (n.startsWith('VM')) {
        if (n.startsWith('VMX')) n = 'M' + n.substring(3);
        else n = 'V' + n.substring(2);
    }

    if (n.includes('DNM')) {
        n = n.replace(/DNM(\d+)0\/(\d+)/, 'DNM$1$2');
    }

    // [3. GENERIC NORMALIZATION]
    n = n.replace(/PUMA|LYNX/g, '').replace(/^P|^L/, '').replace(/\s+/g, '').trim();

    if (n.startsWith('VTR')) {
        let numMatch = n.match(/VTR(\d+)/);
        if (numMatch) {
            let num = numMatch[1];
            if (num === '1620' || num === '162') num = '162';
            else if (num === '1216' || num === '121') num = '121';
            else if (num === '2025' || num === '202') num = '202';
            n = 'VTR' + num;
        }
    }

    let key = n.replace(/[^A-Z1-9]/g, '');
    key = key.replace(/0/g, '');

    if (key.length >= 5) {
        key = key.replace(/[2-9]$/, '');
    }

    key = key.replace(/[A-Z]+$/, '');
    
    if (key.startsWith('DC')) {
        key = key.replace(/[A-Z]$/, '');
    }

    return key;
}

const tests = [
    'VCF850SR', 'VF85SR2',
    'DBM2540U', 'DBM254U',
    'VCF5500SL', 'VCF55SL'
];

tests.forEach(t => console.log(`${t} -> ${getMatchKey(t)}`));
