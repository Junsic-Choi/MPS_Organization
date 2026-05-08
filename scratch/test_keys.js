const getMatchKey = (s) => {
    if (!s) return '';
    let n = s.toString().toUpperCase().trim();
    
    // [BRAND STRIP]
    n = n.replace(/PUMA|LYNX/g, '').trim();

    // [PV/VT SPECIAL]
    if (n.startsWith('PV') || (n.startsWith('VT') && !n.startsWith('VTR')) || (n.startsWith('V') && /\d/.test(n[1]) && !n.startsWith('VCF'))) {
        n = n.replace(/MR$/, 'R').replace(/ML$/, 'L').replace(/M$/, '');
        n = n.replace(/R$/, '').replace(/L$/, '');
    }

    // [PREFIX STRIP]
    n = n.replace(/^P|^L/, '').trim();

    // [SMART SPLIT]
    let parts = n.split('-');
    if (parts.length > 1 && /^[FSHM][0-9]/.test(parts[1])) {
        n = parts[0];
    }

    // [POLISHED]
    n = n.replace(/\/\d{2,}[A-Z]?/g, '').replace(/\s\d+K/g, '');

    // [SAFE ENHANCEMENT]
    if (n.startsWith('VCF')) n = 'V' + n.substring(3);
    if (n.startsWith('VF') && !n.startsWith('VFC')) n = 'V' + n.substring(2);
    if (n.startsWith('DNM')) n = 'D' + n.substring(3);
    if (n.startsWith('DCM')) n = 'DC' + n.substring(3);
    if (n.startsWith('TT') && !n.startsWith('TTR')) n = 'T' + n.substring(2);
    if (n.startsWith('LL')) n = 'L' + n.substring(2);
    
    // 로마자 -> 숫자 변환
    n = n.replace(/VIII/g, '8').replace(/VII/g, '7').replace(/VI/g, '6').replace(/IV/g, '4').replace(/IX/g, '9');
    n = n.replace(/III/g, '3').replace(/II/g, '2');
    
    // [V-SERIES SPECIAL]
    if (n.startsWith('VCF') || (n.startsWith('VF') && !n.startsWith('VFC'))) {
        let digits = n.match(/\d{2}/);
        if (digits) n = 'V' + digits[0];
    }
    
    // [MYNX SPECIAL]
    if (n.startsWith('MYNX')) n = 'M' + n.substring(4);
    else if (n.startsWith('M') && n.length >= 5) {
        if (n.startsWith('M954') || n.startsWith('M955')) {
            n = 'M95' + n.substring(4);
        }
        else if ((n[2] === '4' || n[2] === '5') && n[2] === n[3]) {
            n = n.substring(0, 2) + n.substring(3);
        }
    }

    // [TT/SMX SPECIAL]
    n = n.replace(/SYYB/g, 'SYY').replace(/STB/g, 'SB');

    // [VTR SPECIAL]
    if (n.startsWith('VTR')) {
        if (n.startsWith('VTR1012')) n = 'VTR10' + n.substring(7);
        else if (n.startsWith('VTR1216')) n = 'VTR12' + n.substring(7);
        else if (n.startsWith('VTR1620')) n = 'VTR16' + n.substring(7);
        else if (n.startsWith('VTR2025')) n = 'VTR20' + n.substring(7);
    }

    // [TW SPECIAL]
    if (n.startsWith('TW')) {
        n = n.replace(/([A-Z])\d+$/, '$1');
    }

    // [5AX SPECIAL]
    if (n.includes('DNM') && n.includes('5A')) {
        n = 'DNM355A';
    }

    // [GT SPECIAL]
    if (n.startsWith('GT2600')) {
        n = n.replace('XLMB', 'XLB').replace('XLMA', 'XLA').replace('XMB', 'XB').replace('XMA', 'XA');
    }

    let key = n.replace(/[^A-Z1-9]/g, '');
    if (key.startsWith('DC')) {
        key = key.replace(/[A-Z]$/, '');
    }
    return key;
};

const testCases = [
    "VCF850LSR", "VF85SR2",
    "PUMA VTR1216FC", "VTR12FC",
    "DNM350/5AX", "DNM355A"
];

testCases.forEach(s => {
    console.log(`"${s}" -> "${getMatchKey(s)}"`);
});
