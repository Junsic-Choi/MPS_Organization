function getCategory(name) {
    if (!name) return "ETC";
    const n = name.toUpperCase().trim();
    const isLynx = n.startsWith('LYNX') || n.startsWith('L21') || n.startsWith('L26') || n.startsWith('L32') || 
                   n.startsWith('L20') || n.startsWith('L16') || n.startsWith('ML') || n.startsWith('LS');
    if (isLynx) return 'LYNX';
    if (n.startsWith('PUMA') || n.startsWith('P') || n.startsWith('VT') || n.startsWith('TT')) return 'PUMA';
    if (n.startsWith('VCF') || n.startsWith('VF')) return 'VCF';
    return "OTHERS";
}

console.log('Production: PUMA 5100LB -> ', getCategory('PUMA 5100LB'));
console.log('MPS Code: ML0413 -> ', getCategory('ML0413'));
console.log('MPS Product: P51XLYB-F0TP-0-E30 -> ', getCategory('P51XLYB-F0TP-0-E30'));
