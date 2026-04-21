const express = require('express');
const cors = require('cors');
const apiRoutes = require('./routes/api');
require('dotenv').config();
const path = require('path');

const app = express();
app.use('/uploads', express.static(path.join(__dirname, '../uploads')));
const PORT = 3022;

app.use(cors());
app.use(express.json());

// Rotaları bağla
app.use('/api', apiRoutes);

app.listen(PORT, () => {
    console.log(`🚀 Sunucu ${PORT} portunda çalışıyor...`);
    console.log(`📡 Endpoint: http://localhost:${PORT}/api/convert`);
});



