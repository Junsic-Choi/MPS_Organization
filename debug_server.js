const fs = require('fs');
console.log('Starting debug wrapper...');
try {
    process.on('uncaughtException', (err) => {
        fs.writeFileSync('server_crash.txt', 'Uncaught Exception:\n' + err.stack);
        process.exit(1);
    });
    require('./server.js');
} catch (e) {
    fs.writeFileSync('server_crash.txt', 'Require Error:\n' + e.stack);
    process.exit(1);
}
