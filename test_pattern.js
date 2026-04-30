function getPatternGroup(p) {
    const up = (p || '').toString().toUpperCase();
    if (up.includes('DBC') || up.includes('DBD') || up.includes('DBM')) return '34. BORING';
    if (up.includes('VCF') || up.includes('VF8')) return '17. VCF 850 Series';
    if (up.includes('VTR')) return '33. VTR Series';
    if (up.includes('VT') || up.includes('VTL')) return '30. PV/VT Series';
    if (up.includes('MYNX') || up.includes('M65') || up.includes('M75')) return '12. MYNX Series';
    if (up.includes('DNM') || up.includes('DEM')) return '13. DNM/DEM Series';
    if (up.includes('DVF')) return '11. DVF Series';
    if (up.includes('SMX')) return '21. SMX Series';
    if (up.includes('DC325') || up.startsWith('DC')) return '35. DC Series';
    if (up.includes('PUMA')) return 'PUMA Series';
    if (up.includes('LYNX')) return 'LYNX Series';
    return '';
}

console.log('Result for DNM755L:', getPatternGroup('DNM755L-FOMP-0-E30'));
console.log('Result for DNM355A:', getPatternGroup('DNM355A-FOMP-0-U30'));
