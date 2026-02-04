const axios = require('axios');
const fs = require('fs');
const path = require('path');
require('dotenv').config();

exports.uploadToCDN = async (fileBuffer, fileName) => {
    // Dosya ismini temizle
    const cleanName = fileName.replace(/[^a-zA-Z0-9.]/g, '_');
    const uniqueFileName = `${Date.now()}_${cleanName}`;

    // --- BYPASS MODU KONTROLÜ ---
    if (process.env.USE_LOCAL_STORAGE === 'true') {
        console.log("🚧 DEV MODU: Dosya yerel diske kaydediliyor...");

        // 1. Klasör yoksa oluştur
        const uploadDir = path.join(__dirname, '../../uploads');
        if (!fs.existsSync(uploadDir)) {
            fs.mkdirSync(uploadDir, { recursive: true });
        }

        // 2. Dosyayı diske yaz
        const filePath = path.join(uploadDir, uniqueFileName);
        fs.writeFileSync(filePath, fileBuffer);

        // 3. Ngrok veya Local URL döndür
        // Eğer .env dosyasında NGROK_URL varsa onu kullan, yoksa localhost dön
        const baseUrl = process.env.NGROK_URL || `http://localhost:${process.env.PORT}`;
        return `${baseUrl}/uploads/${uniqueFileName}`;
    }
    // -----------------------------

    // AŞAĞISI NORMAL BUNNYCDN KODU (Aynen kalıyor)
    const filePath = `uploads/${uniqueFileName}`;
    const storageUrl = `https://${process.env.BUNNY_STORAGE_HOST}/${process.env.BUNNY_STORAGE_ZONE}/${filePath}`;

    try {
        await axios.put(storageUrl, fileBuffer, {
            headers: {
                'AccessKey': process.env.BUNNY_ACCESS_KEY,
                'Content-Type': 'application/octet-stream'
            },
            maxBodyLength: Infinity,
            maxContentLength: Infinity
        });

        return `${process.env.BUNNY_PULL_ZONE_URL}/${filePath}`;

    } catch (error) {
        console.error("BunnyCDN Upload Hatası:", error.response ? error.response.data : error.message);
        throw new Error("CDN Yükleme başarısız");
    }
};




