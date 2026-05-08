const s = 'MYNX7500/50 II';
let n = s.toString().toUpperCase().trim();
console.log('1. Start:', n);

n = n.replace(/\/\d{2,}[A-Z]?/g, '').replace(/\s\d+K/g, '');
console.log('2. Stripped spec:', n);

if (n.startsWith('MYNX')) n = 'M' + n.substring(4);
console.log('3. Normalize prefix:', n);

n = n.replace(/PUMA|LYNX/g, '').replace(/^P|^L/, '').trim();
console.log('4. Strip branding:', n);

n = n.replace(/VIII/g, '8').replace(/VII/g, '7').replace(/VI/g, '6').replace(/IV/g, '4').replace(/IX/g, '9');
n = n.replace(/III/g, '3').replace(/II/g, '2');
console.log('5. Roman to Num:', n);

if (n.startsWith('M') && n.length >= 4) {
    if (n.startsWith('M954') || n.startsWith('M955')) {
        n = 'M95' + n.substring(4);
    } else if (n.length >= 5 && (n[2] === '4' || n[2] === '5') && /\d/.test(n[3])) {
        n = n.substring(0, 2) + n.substring(3);
    }
}
console.log('6. MYNX Special:', n);

let key = n.replace(/[^A-Z1-9]/g, '');
console.log('7. Final Key (pre-strip):', key);

key = key.replace(/([A-Z])[2-9]$/, '$1');
console.log('8. Final Key (post-strip):', key);
