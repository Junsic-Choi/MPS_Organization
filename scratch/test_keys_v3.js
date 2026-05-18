function getMatchKey(s) {
    if (!s) return '';
    let n = s.toString().toUpperCase().trim();
    
    n = n.replace(/VIII/g, '8').replace(/VII/g, '7').replace(/VI/g, '6').replace(/IV/g, '4').replace(/IX/g, '9');
    n = n.replace(/III/g, '3').replace(/II/g, '2');

    if (n.includes('SMX')) {
        n = n.replace(/SMX2(?![0-9])/g, 'SMX21');
        n = n.replace(/2100/g, '21').replace(/3100/g, '31').replace(/5100/g, '51');
        n = n.replace(/SYYB/g, 'SYY').replace(/STB/g, 'SB');
    }

    if (n.startsWith('VCF')) {
        n = n.replace('VCF', 'VF');
        // VCF850 -> VF85 -> VF8
        let m = n.match(/VF(\d+)/);
        if (m) n = 'VF' + m[1].substring(0, 1);
    } else if (n.startsWith('VF')) {
        let m = n.match(/VF(\d+)/);
        if (m) n = 'VF' + m[1].substring(0, 1);
    }

    if (n.startsWith('VTR')) {
        let numMatch = n.match(/VTR(\d+)/);
        if (numMatch) {
            let num = numMatch[1];
            // VTR1216 -> VTR12, VTR1012 -> VTR10
            n = 'VTR' + num.substring(0, 2);
        }
    }

    if (n.startsWith('MYNX')) n = 'M' + n.substring(4);
    n = n.replace(/PUMA|LYNX/g, '').replace(/^P|^L/, '').replace(/\s+/g, '').trim();

    let key = n.replace(/[^A-Z1-9]/g, '');
    key = key.replace(/0/g, '');
    
    // Final grouping for VF if not already handled
    if (key.startsWith('VF')) {
        key = 'VF' + key.replace(/[^1-9]/g, '').substring(0, 1);
    }
    
    return key;
}

const testCases = [
    'PUMA VTR1216FC', 'VTR12FC',
    'PUMA VTR1012FC', 'VTR10FC',
    'VCF850LSR', 'VF8LSR2',
    'VCF5500L', 'VF5LSR'
];

testCases.forEach(t => console.log(`${t} -> ${getMatchKey(t)}`));
