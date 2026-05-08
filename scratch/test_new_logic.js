function getMatchKey_New(s) {
    if (!s) return '';
    let n = s.toString().toUpperCase().trim();
    
    // [1. ROMAN TO NUMBER]
    n = n.replace(/VIII/g, '8').replace(/VII/g, '7').replace(/VI/g, '6').replace(/IV/g, '4').replace(/IX/g, '9');
    n = n.replace(/III/g, '3').replace(/II/g, '2');

    // [2. SPECIALIZED SERIES RULES]
    if (n.includes('SMX')) {
        n = n.replace(/2100/g, '21').replace(/3100/g, '31').replace(/5100/g, '51');
        n = n.replace(/SYYB/g, 'SYY').replace(/STB/g, 'SB');
    }

    if (n.startsWith('VCF')) {
        n = n.replace('VCF', 'VF');
        n = n.replace(/(\d)\d0/, '$1');
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

    if (n.startsWith('VTR')) {
        let numMatch = n.match(/VTR(\d+)/);
        if (numMatch) {
            let num = numMatch[1];
            if (num === '1620') num = '162';
            else if (num === '1216') num = '121';
            n = 'VTR' + num;
        }
    }

    // [3. GENERIC NORMALIZATION]
    n = n.replace(/PUMA|LYNX/g, '').replace(/^P|^L/, '').replace(/\s+/g, '').trim();

    if (n.startsWith('MYNX')) n = 'M' + n.substring(4);
    else if (n.startsWith('VMX')) n = 'M' + n.substring(3);
    else if (n.startsWith('VM') && !n.startsWith('VMX')) n = 'M' + n.substring(2);
    else if (n.startsWith('MP')) n = 'M' + n.substring(2);
    else if (n.startsWith('DNM')) n = 'D' + n.substring(3);
    else if (n.startsWith('DCM')) n = 'DC' + n.substring(3);
    else if (n.startsWith('DVF')) n = 'V' + n.substring(3);
    else if (n.startsWith('VCF')) n = 'V' + n.substring(3);
    else if (n.startsWith('VT') && !n.startsWith('VTR')) n = 'V' + n.substring(2);
    else if (n.startsWith('TT') && !n.startsWith('TTR')) n = 'T' + n.substring(2);
    
    if (n.startsWith('V') && !n.startsWith('VTR') && !n.startsWith('VFC')) {
        let digits = n.match(/\d{2,3}/);
        if (digits) n = 'V' + digits[0].substring(0, 2);
    }
    if (n.startsWith('TW')) {
        n = n.replace(/(\d+)(?:MZ|WB|W|B|Z|M)+\d*$/g, '$1');
        let base = n.match(/TW\d+/);
        if (base) n = base[0];
    }

    if (n.startsWith('GT2600')) {
        n = n.replace('XLMB', 'XLB').replace('XLMA', 'XLA').replace('XMB', 'XB').replace('XMA', 'XA');
    }

    let key = n.replace(/[^A-Z1-9]/g, '');
    key = key.replace(/0/g, '');

    // [MODIFIED VERSION STRIP]: Remove trailing version digit (2-9) if length is enough
    if (key.length >= 4) {
        key = key.replace(/[2-9]$/, '');
    }

    // [GLOBAL SUFFIX STRIP]
    key = key.replace(/[A-Z]+$/, '');
    
    return key;
}

const testModels = [
    'VCF850LSR', 
    'VF8LSR2',
    'VCF850SR',
    'VF85SR2',
    'SMX2100STB',
    'SMX21ST',
    'SMX2100S',
    'SMX210S'
];

testModels.forEach(m => {
    console.log(`"${m}" -> "${getMatchKey_New(m)}"`);
});
