const fs = require('fs');
const path = require('path');

// Oluşturulacak Dosyalar ve İçerikleri
const files = {
    // 1. Package.json: Gerekli kütüphaneler (AWS SDK kaldırıldı)
    'package.json': JSON.stringify({
        name: "dior-packing-backend",
        version: "2.0.0",
        description: "n8n packing list parser with BunnyCDN",
        main: "src/app.js",
        scripts: {
            "start": "node src/app.js",
            "dev": "nodemon src/app.js"
        },
        dependencies: {
            "express": "^4.18.2",
            "multer": "^1.4.5-lts.1",
            "axios": "^1.6.2",
            "dotenv": "^16.3.1",
            "cors": "^2.8.5"
        }
    }, null, 2),

    // 2. .env: BunnyCDN ve n8n Ayarları
    '.env': `PORT=3000

# BunnyCDN Ayarları
# Storage Host: Genelde "storage.bunnycdn.com" veya "ny.storage.bunnycdn.com" (Zone panelinde yazar)
BUNNY_STORAGE_HOST=storage.bunnycdn.com

# Storage Zone Name: Bunny panelinde oluşturduğun Storage Zone adı
BUNNY_STORAGE_ZONE=dior-packing-storage

# Access Key: Storage Zone > FTP & API Access kısmındaki şifre (Password)
BUNNY_ACCESS_KEY=buraya_bunny_sifreni_yaz

# Pull Zone URL: Dosyaları okumak için bağlı olan CDN URL'i (sonunda / olmasın)
BUNNY_PULL_ZONE_URL=https://senin-cdn-adresin.b-cdn.net

# n8n Webhook URL (POST metodlu olan)
N8N_WEBHOOK_URL=https://senin-n8n-adresin.com/webhook/process-packing-list
`,

    // 3. Service: BunnyCDN Yükleme İşlemleri
    'src/services/storageService.js': `
const axios = require('axios');
require('dotenv').config();

exports.uploadToCDN = async (fileBuffer, fileName) => {
    // Dosya ismini temizle ve benzersiz yap
    const cleanName = fileName.replace(/[^a-zA-Z0-9.]/g, '_');
    const uniqueFileName = \`\${Date.now()}_\${cleanName}\`;
    
    // BunnyCDN Storage Yolu: /uploads/ klasörü altına koyalım
    const filePath = \`uploads/\${uniqueFileName}\`;

    // 1. Storage API Endpoint'i (Dosya Yükleme Yeri)
    // Format: https://{StorageHost}/{StorageZoneName}/{FilePath}
    const storageUrl = \`https://\${process.env.BUNNY_STORAGE_HOST}/\${process.env.BUNNY_STORAGE_ZONE}/\${filePath}\`;

    try {
        // BunnyCDN'e dosyayı PUT metoduyla gönderiyoruz
        await axios.put(storageUrl, fileBuffer, {
            headers: {
                'AccessKey': process.env.BUNNY_ACCESS_KEY,
                'Content-Type': 'application/octet-stream' // Bunny otomatik algılar
            },
            maxBodyLength: Infinity,
            maxContentLength: Infinity
        });

        // 2. Public CDN URL'ini oluştur (İnsanların göreceği link)
        // Format: https://{PullZoneUrl}/{FilePath}
        const publicUrl = \`\${process.env.BUNNY_PULL_ZONE_URL}/\${filePath}\`;
        
        return publicUrl;

    } catch (error) {
        // Hata detayını logla
        console.error("BunnyCDN Upload Hatası:", error.response ? error.response.data : error.message);
        throw new Error("CDN Yükleme başarısız");
    }
};
`,

    // 4. Service: n8n ile İletişim
    'src/services/n8nService.js': `
const axios = require('axios');

exports.processFileWithN8N = async (fileUrl) => {
    try {
        console.log(">> n8n servisine istek atılıyor...");
        
        // n8n Webhook'una POST isteği (Binary cevap bekliyoruz)
        const response = await axios.post(
            process.env.N8N_WEBHOOK_URL,
            { file_url: fileUrl },
            {
                responseType: 'arraybuffer', // ÖNEMLİ: CSV dosyası buffer olarak gelecek
                timeout: 300000 // 5 Dakika timeout (Gemini uzun sürebilir)
            }
        );

        console.log(">> n8n'den başarılı yanıt alındı!");
        return response.data;
    } catch (error) {
        console.error("n8n İletişim Hatası:", error.message);
        throw new Error("n8n işlemi başarısız veya zaman aşımı");
    }
};
`,

    // 5. Controller: Tüm Mantığı Yöneten Beyin
    'src/controllers/packingController.js': `
const storageService = require('../services/storageService');
const n8nService = require('../services/n8nService');

exports.convertPackingList = async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: "Lütfen bir dosya yükleyin" });
        }

        console.log("1. İstek alındı, dosya BunnyCDN'e yükleniyor...");

        // ADIM 1: Gelen ham dosyayı BunnyCDN'e yükle
        const inputCdnUrl = await storageService.uploadToCDN(
            req.file.buffer,
            req.file.originalname
        );

        console.log("-> Ham Dosya Yüklendi:", inputCdnUrl);
        console.log("2. n8n'e gönderiliyor ve cevap bekleniyor...");

        // ADIM 2: n8n'e gönder ve CSV cevabını bekle (Await)
        const csvBuffer = await n8nService.processFileWithN8N(inputCdnUrl);

        console.log("-> n8n cevabı (CSV) alındı, BunnyCDN'e yükleniyor...");

        // ADIM 3: Gelen CSV dosyasını BunnyCDN'e yükle
        const outputFileName = 'result_packing_list.csv';
        const outputCdnUrl = await storageService.uploadToCDN(
            csvBuffer,
            outputFileName
        );

        console.log("-> İşlem Tamamlandı! CSV Linki:", outputCdnUrl);

        // ADIM 4: Kullanıcıya linkleri dön
        res.status(200).json({
            success: true,
            message: "Dosya başarıyla işlendi",
            input_url: inputCdnUrl,
            csv_url: outputCdnUrl
        });

    } catch (error) {
        console.error("Controller Hatası:", error);
        res.status(500).json({ 
            success: false, 
            error: error.message 
        });
    }
};
`,

    // 6. Routes: API Endpoint Tanımı
    'src/routes/api.js': `
const express = require('express');
const router = express.Router();
const multer = require('multer');
const packingController = require('../controllers/packingController');

// Dosyayı bellekte (RAM) tut, diske yazma
const upload = multer({ storage: multer.memoryStorage() });

// POST /api/convert
router.post('/convert', upload.single('file'), packingController.convertPackingList);

module.exports = router;
`,

    // 7. App: Sunucu Giriş Noktası
    'src/app.js': `
const express = require('express');
const cors = require('cors');
const apiRoutes = require('./routes/api');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

// Rotaları bağla
app.use('/api', apiRoutes);

app.listen(PORT, () => {
    console.log(\`🚀 Sunucu \${PORT} portunda çalışıyor...\`);
    console.log(\`📡 Endpoint: http://localhost:\${PORT}/api/convert\`);
});
`
};

// Klasör ve Dosya Oluşturma Mantığı
Object.entries(files).forEach(([filePath, content]) => {
    const fullPath = path.join(__dirname, filePath);
    const dirName = path.dirname(fullPath);

    // Klasör yoksa oluştur
    if (!fs.existsSync(dirName)) {
        fs.mkdirSync(dirName, { recursive: true });
    }

    // Dosyayı yaz
    fs.writeFileSync(fullPath, content.trim());
    console.log(`✅ Oluşturuldu: ${filePath}`);
});

console.log("\n🐰 BunnyCDN Kurulumu Tamamlandı!");
console.log("Şimdi şu adımları yap:");
console.log("1. npm install (Paketleri yükle)");
console.log("2. .env dosyasını aç ve BunnyCDN bilgilerini gir");
console.log("3. npm start");