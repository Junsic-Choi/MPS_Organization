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

// Endpoint to trigger the extraction (Now using Node-only libraries)
app.post('/api/extract', (req, res) => {
    console.log('Extraction requested from client using Node-XLSX.');

    // Execute the PowerShell extraction script (handles DRM files via COM)
    const command = `powershell -ExecutionPolicy Bypass -NoProfile -File Auto_Make_CSV.ps1`;

    exec(command, (error, stdout, stderr) => {
        if (error) {
            console.error(`Extraction error: ${error}`);
            return res.status(500).json({ error: 'Extraction failed', details: error.message });
        }

        console.log(`Extraction stdout: ${stdout}`);
        if (stderr) {
            console.error(`Extraction stderr: ${stderr}`);
        }

        // Find the newest CSV file
        try {
            const files = fs.readdirSync(__dirname);
            const csvFiles = files.filter(f => f.endsWith('_FinalList.csv'));

            if (csvFiles.length === 0) {
                return res.status(404).json({ error: 'Extraction completed, but no FinalList.csv found.' });
            }

            csvFiles.sort((a, b) => {
                return fs.statSync(path.join(__dirname, b)).mtime.getTime() - fs.statSync(path.join(__dirname, a)).mtime.getTime();
            });

            const latestCsv = csvFiles[0];
            console.log(`Extraction successful. Found file: ${latestCsv}`);
            res.json({ success: true, file: latestCsv });
        } catch (readError) {
            console.error(`Error reading output directory: ${readError}`);
            res.status(500).json({ error: 'Failed to read extraction result' });
        }
    });
});

app.listen(PORT, '0.0.0.0', () => {
    console.log(`====================================================`);
    console.log(`🚀 MPS Dashboard Server is running!`);
    console.log(`🌐 Local access URL: http://localhost:${PORT}`);
    console.log(`📡 Network access URL (from other PCs): Check your IP and use port ${PORT}`);
    console.log(`====================================================`);
});
