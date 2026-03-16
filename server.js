const express = require('express');
const { exec } = require('child_process');
const path = require('path');
const fs = require('fs');
const cors = require('cors');

const app = express();
const PORT = 8888; // Default port

app.use(cors());
app.use(express.static(path.join(__dirname)));
app.use(express.json());

// Redirect root to dashboard.html
app.get('/', (req, res) => {
    res.redirect('/dashboard.html');
});

// Endpoint to trigger the PowerShell extraction script
app.post('/api/extract', (req, res) => {
    console.log('Extraction requested from client.');
    const scriptPath = path.join(__dirname, 'Auto_Make_CSV.ps1');

    // Execute the PowerShell script
    const command = `powershell.exe -ExecutionPolicy Bypass -NoProfile -File "${scriptPath}"`;

    exec(command, (error, stdout, stderr) => {
        if (error) {
            console.error(`exec error: ${error}`);
            // Log it but we might still have a partial success, we'll return error status
            return res.status(500).json({ error: 'Extraction failed', details: error.message });
        }

        console.log(`Extraction stdout: ${stdout}`);
        if (stderr) {
            console.error(`Extraction stderr: ${stderr}`);
        }

        // Let's find the newest generated CSV file to return to the client
        const files = fs.readdirSync(__dirname);
        const csvFiles = files.filter(f => f.endsWith('_FinalList.csv'));

        if (csvFiles.length === 0) {
            return res.status(404).json({ error: 'Extraction completed, but no FinalList.csv found.' });
        }

        // Sort by modified time descending to get the newest file (in case of multiple)
        csvFiles.sort((a, b) => {
            return fs.statSync(path.join(__dirname, b)).mtime.getTime() - fs.statSync(path.join(__dirname, a)).mtime.getTime();
        });

        const latestCsv = csvFiles[0];
        console.log(`Extraction successful. Found file: ${latestCsv}`);
        res.json({ success: true, file: latestCsv });
    });
});

app.listen(PORT, '0.0.0.0', () => {
    console.log(`====================================================`);
    console.log(`🚀 MPS Dashboard Server is running!`);
    console.log(`🌐 Local access URL: http://localhost:${PORT}`);
    console.log(`📡 Network access URL (from other PCs): Check your IP and use port ${PORT}`);
    console.log(`====================================================`);
});
