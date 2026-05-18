const http = require('http');

const data = JSON.stringify({
    filename: 'MPS2605-1.xlsx'
});

const options = {
    hostname: 'localhost',
    port: 8890,
    path: '/api/extract-live',
    method: 'POST',
    headers: {
        'Content-Type': 'application/json',
        'Content-Length': data.length
    }
};

const req = http.request(options, (res) => {
    console.log(`STATUS: ${res.statusCode}`);
    res.setEncoding('utf8');
    res.on('data', (chunk) => {
        console.log(`BODY: ${chunk.substring(0, 100)}...`);
    });
});

req.on('error', (e) => {
    console.error(`problem with request: ${e.message}`);
});

req.write(data);
req.end();
