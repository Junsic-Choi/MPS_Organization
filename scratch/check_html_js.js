const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf8');
const scriptMatch = html.match(/<script>([\s\S]*?)<\/script>/);
if (scriptMatch) {
    const js = scriptMatch[1];
    try {
        new Function(js);
        console.log('JS syntax is valid');
    } catch (e) {
        console.log('JS syntax error:', e.message);
        // Find the line number
        const lines = js.split('\n');
        // Eval line by line to find error
        let current = '';
        for (let i = 0; i < lines.length; i++) {
            current += lines[i] + '\n';
            try {
                new Function(current + '}'); // Try to close any blocks
            } catch (err) {
                // If it's a syntax error that's not just "unexpected end of input"
                if (!err.message.includes('Unexpected end of input')) {
                    // This might be it, but it's tricky with brackets.
                }
            }
        }
    }
} else {
    console.log('No script tag found');
}
