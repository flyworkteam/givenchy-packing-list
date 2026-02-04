const path = require('path');
const XLSX = require('xlsx'); // xlsx kütüphanesi
const storageService = require('../services/storageService');
const n8nService = require('../services/n8nService');

exports.convertPackingList = async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: "Lütfen bir dosya yükleyin" });
        }

        console.log("🚀 İşlem Başlıyor...");
        const originalName = req.file.originalname;

        // 1. BunnyCDN'e Yükle
        const inputCdnUrl = await storageService.uploadToCDN(req.file.buffer, originalName);

        // 2. n8n'den CSV Buffer al
        const n8nResponseBuffer = await n8nService.processFileWithN8N(inputCdnUrl);

        console.log("3. Excel dönüşümü ve Hücre Tipi Ayarları yapılıyor...");

        // --- KRİTİK DEĞİŞİKLİK BAŞLANGICI ---

        // A) Okurken 'raw: true' kullanıyoruz. 
        // Bu sayede Excel uzun numaraları otomatik yuvarlayıp bozamaz. Hepsi String gelir.
        const workbook = XLSX.read(n8nResponseBuffer, { type: 'buffer', raw: true });

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // B) Hücreleri tek tek gezip SADECE sayı olması gerekenleri sayıya çeviriyoruz.
        // GUMRUK_NO (A sütunu) ve PARCEL_NUMBER (B sütunu) -> METİN kalmalı.
        // QUANTITY (E sütunu) ve AU (F sütunu) -> SAYI olmalı.

        const range = XLSX.utils.decode_range(worksheet['!ref']); // Dolu alanın sınırlarını al

        // Sütun İndeksleri (0'dan başlar): A=0, B=1, C=2, D=3, E=4, F=5
        const quantityColIndex = 3; // E Sütunu
        const auColIndex = 4;       // F Sütunu

        for (let R = range.s.r + 1; R <= range.e.r; ++R) { // Başlık (0. satır) hariç döngü
            // Quantity (E Sütunu) Dönüştürme
            const quantityCellAddress = XLSX.utils.encode_cell({ c: quantityColIndex, r: R });
            const quantityCell = worksheet[quantityCellAddress];
            if (quantityCell && quantityCell.v) {
                quantityCell.t = 'n'; // Type: Number yap
                quantityCell.v = Number(quantityCell.v); // Değeri sayıya çevir
            }

            // AU (F Sütunu) Dönüştürme
            const auCellAddress = XLSX.utils.encode_cell({ c: auColIndex, r: R });
            const auCell = worksheet[auCellAddress];
            if (auCell && auCell.v) {
                auCell.t = 'n'; // Type: Number yap
                auCell.v = Number(auCell.v); // Değeri sayıya çevir
            }
        }

        // C) Dosyayı buffer'a yaz
        const xlsxBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });

        // --- KRİTİK DEĞİŞİKLİK SONU ---

        const outputFileName = `${path.parse(originalName).name}_analysis.xlsx`;
        console.log(`4. Sonuç dosyası (${outputFileName}) BunnyCDN'e yükleniyor...`);
        const outputCdnUrl = await storageService.uploadToCDN(xlsxBuffer, outputFileName);
        
        console.log("-> İşlem Tamam! Excel Linki:", outputCdnUrl);

        res.status(200).json({
            success: true,
            message: "Dönüştürme başarılı",
            input_url: inputCdnUrl,
            xlsx_url: outputCdnUrl,
            file_name: outputFileName
        });

    } catch (error) {
        console.error("❌ Controller Hatası:", error.message);
        res.status(500).json({ success: false, error: error.message });
    }
};