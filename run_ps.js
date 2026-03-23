const { spawn } = require('child_process');
const fs = require('fs');

const logFile = fs.createWriteStream('extraction_full_log.txt');

const ps = spawn('powershell.exe', [
    '-ExecutionPolicy', 'Bypass',
    '-File', 'Auto_Extract_Final.ps1'
]);

ps.stdout.on('data', (data) => {
    console.log(data.toString());
    logFile.write(data);
});

ps.stderr.on('data', (data) => {
    console.error(data.toString());
    logFile.write('ERROR: ' + data);
});

ps.on('close', (code) => {
    logFile.write(`\nProcess exited with code ${code}`);
    logFile.end();
    console.log(`Exit code: ${code}`);
});
