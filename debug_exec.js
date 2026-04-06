const { exec } = require('child_process');
const path = require('path');

const scriptPath = path.join(__dirname, 'Final_Extract_4650.ps1');
const command = `powershell -ExecutionPolicy Bypass -File "${scriptPath}"`;

console.log('Executing:', command);
exec(command, (error, stdout, stderr) => {
    console.log('--- STDOUT ---');
    console.log(stdout);
    console.log('--- STDERR ---');
    console.log(stderr);
    if (error) {
        console.log('--- ERROR ---');
        console.error(error);
    }
});
