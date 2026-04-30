const fs = require('fs');
const XLSX = require('xlsx');

const MASTER_GROUPS = [
  '01. HM Series', '02. HFP Series', '02. NHM Series', '03. HC/HP Series',
  '03. HSP Series', '04-1. DHF Series', '04. NHC/NHP Series', '05. 문형MC(W/F TYPE)',
  '06. BM Series', '07. DBM Series', '08. VM Series', '10. DNM 5AX Series',
  '11. DVF Series', '11. NX / MP / VM 소형 Series', '12. MD Series',
  '12. MYNX Series', '13-1. AV4 Series', '13-1. New DNM Series', '13-3. DEM4000',
  '13-4. SVM Series', '13-5. BVM5700', '13. DNM Series', '15. VC430/500/510 Series',
  '16. T4000/T3600D  Series', '17. VCF 5500 Series', '17. VCF 850 Series',
  '18. S_turn', '19. GT Series', '20. DNT Series', '20. LYNX Series',
  '20. New LYNX Series', '20.1 DNC Series', '20.1 LYNX XG Series', '21. P2100 Series',
  '21.1. P2100 Series II', '22. P2600 Series', '22.1. P2600 Series II', '23. P3100 Series',
  '25. P4100 Series', '25. P5100 Series', '26. P500/600/700/800 Series',
  '27. P600/700/800 XLY/XLM Series', '28. SMX Series', '29. DNX Series',
  '29. TL/TT Series', '30. PV/VT Series', '31. IV/VAW Series', '31. PVX/AW  Series',
  '32. TW Series', '33. VTR Series', '34. BORING', '38. GVX Series', '90. LEO Series'
];

function getPatternGroup(p) {
    const up = (p || '').toString().toUpperCase();
    if (up.includes('DBC') || up.includes('DBD') || up.includes('DBM')) return '34. BORING';
    if (up.includes('VCF') || up.includes('VF8')) return '17. VCF 850 Series';
    if (up.includes('DEM')) return '13-3. DEM4000';
    if (up.includes('DNM 5AX')) return '10. DNM 5AX Series';
    if (up.includes('DNM')) return '13. DNM Series';
    if (up.includes('VTR')) return '33. VTR Series';
    if (up.includes('VT') || up.includes('VTL')) return '30. PV/VT Series';
    if (up.includes('MYNX') || up.includes('M65') || up.includes('M75')) return '12. MYNX Series';
    if (up.includes('DVF')) return '11. DVF Series';
    if (up.includes('SMX')) return '21. SMX Series';
    if (up.includes('P41') || up.includes('PUMA 4100')) return '25. P4100 Series';
    if (up.includes('DC325') || up.startsWith('DC')) return '35. DC Series';
    if (up.includes('PUMA')) return 'PUMA Series';
    if (up.includes('LYNX')) return 'LYNX Series';
    return '';
}

const mismatch = [];
const patterns = ['DBC', 'VCF', 'DEM', 'DNM 5AX', 'DNM', 'VTR', 'VTL', 'MYNX', 'DVF', 'SMX', 'P4100', 'DC325', 'PUMA', 'LYNX'];

patterns.forEach(p => {
    let gp = getPatternGroup(p);
    if (gp && !MASTER_GROUPS.includes(gp)) {
        mismatch.push(`${p} => ${gp}`);
    }
});

console.log('--- Mismatched Pattern Groups with Master Sheet ---');
console.log(mismatch.join('\n'));
