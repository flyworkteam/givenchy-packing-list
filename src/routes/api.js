const express = require('express');
const router = express.Router();
const multer = require('multer');
const packingController = require('../controllers/packingController');

// Dosyayı bellekte (RAM) tut, diske yazma
const upload = multer({ storage: multer.memoryStorage() });

// POST /api/convert
router.post('/convert', upload.single('file'), packingController.convertPackingList);

module.exports = router;