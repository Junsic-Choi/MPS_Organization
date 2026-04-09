const express = require('express');
const cors = require('cors');
const app = express();
const PORT = 8889;
app.use(cors());
app.use(express.static(__dirname));
app.get('/health', (req, res) => res.send('OK'));
app.listen(PORT, () => console.log(`Test server: http://localhost:${PORT}/health`));
