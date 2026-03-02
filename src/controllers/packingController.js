const path = require('path');
const os = require('os');
const fs = require('fs').promises;
const { exec } = require('child_process');
const util = require('util');
const execPromise = util.promisify(exec);
const xlsx = require('xlsx'); // EXCEL KÜTÜPHANESİ

const storageService = require('../services/storageService');

/**
 * Metin haline gelmiş PDF'i kurallara göre JSON formatına dönüştürür.
 */
function parsePackingList(text) {
    const lines = text.split('\n');
    const products = [];

    let currentGumrukNo = null;
    let currentProduct = null;

    // KURAL DEĞİŞİKLİĞİ: P ile başlayan 7 hane olacak AMA içinde en az 1 tane rakam (?=.*\d) barındıracak.
    // Bu sayede "PARCELS", "PACKAGE" gibi sadece harften oluşan kelimeler elenmiş olur.
    const productRegex = /(P(?=.*\d)[A-Za-z0-9]{6})\s+(.*)\s+(\d+)(?:\s|$)/i;
    const auRegex = /(\d+)\s+(?:COLIS IDENTIQUES|IDENTICAL PARCELS)/i;

    for (let i = 0; i < lines.length; i++) {
        const rawLine = lines[i];
        const trimmedLine = rawLine.trim();

        if (!trimmedLine) continue;

        // KURAL 1: Gümrük numarası tam 9 hane aranıyor.
        const gumrukMatch = rawLine.match(/^\s{0,5}(\d{9})\b/);

        if (gumrukMatch) {
            currentGumrukNo = gumrukMatch[1];
        }
        else if (/(?:HANDLING UNIT N°|UNITE DE MANUTENTION)\s*(\d{9})/i.test(trimmedLine)) {
            const match = trimmedLine.match(/(?:HANDLING UNIT N°|UNITE DE MANUTENTION)\s*(\d{9})/i);
            if (match) currentGumrukNo = match[1];
        }

        // KURAL 2: Ürün Satırını Yakalama
        const prodMatch = trimmedLine.match(productRegex);

        // Ekstra Güvenlik: Yakalanan referans her ihtimale karşı 'PARCELS' ise es geç
        if (prodMatch && prodMatch[1].toUpperCase() !== 'PARCELS') {
            if (currentProduct) {
                products.push(currentProduct);
            }

            currentProduct = {
                gumruk_no: currentGumrukNo,
                reference: prodMatch[1].toUpperCase(),    // Harfleri standart büyük harf yapar
                description: prodMatch[2].trim(),         // Açıklama kısmı
                quantity: prodMatch[3],                   // Miktar
                au: "1"
            };
        }

        // KURAL 3: AU (Koli Sayısı) Yakalama
        const auMatch = trimmedLine.match(auRegex);
        if (auMatch && currentProduct) {
            currentProduct.au = auMatch[1];
        }
    }

    if (currentProduct) {
        products.push(currentProduct);
    }

    // Fallback: Herhangi bir aksilikte Gümrük Numarası boş kalanlara en son geçerli olanı ata
    let lastValidGumruk = null;
    for (let p of products) {
        if (p.gumruk_no) lastValidGumruk = p.gumruk_no;
        else if (lastValidGumruk) p.gumruk_no = lastValidGumruk;
    }

    return products;
}

exports.convertPackingList = async (req, res) => {
    let tempPdfPath = null;
    let tempTxtPath = null;

    try {
        if (!req.file) {
            return res.status(400).json({ error: "Lütfen bir dosya yükleyin" });
        }

        console.log("🚀 İşlem Başlıyor...");
        const originalName = req.file.originalname;

        // ADIM 1: Gelen PDF'i BunnyCDN'e Yükle
        console.log(`1. ${originalName} BunnyCDN'e yükleniyor...`);
        const inputCdnUrl = await storageService.uploadToCDN(
            req.file.buffer,
            originalName
        );
        console.log("-> PDF Linki:", inputCdnUrl);

        // ADIM 2: PDF'i yerel diske geçici olarak yaz
        console.log("2. PDF analiz için yerel diske hazırlanıyor...");
        const tempFileName = `packing_${Date.now()}_${Math.floor(Math.random() * 1000)}`;
        tempPdfPath = path.join(os.tmpdir(), `${tempFileName}.pdf`);
        tempTxtPath = path.join(os.tmpdir(), `${tempFileName}.txt`);

        await fs.writeFile(tempPdfPath, req.file.buffer);

        // ADIM 3: pdftotext ile PDF'i metne çevir
        console.log("3. PDF verileri metne dönüştürülüyor...");
        await execPromise(`pdftotext -layout "${tempPdfPath}" "${tempTxtPath}"`);

        // ADIM 4: Oluşan metni oku ve JSON'a çevir
        console.log("4. Metin verileri analiz ediliyor...");
        const pdfText = await fs.readFile(tempTxtPath, 'utf8');
        const parsedJsonData = parsePackingList(pdfText);

        // --- ADIM 5: JSON'U ÖZEL BAŞLIKLI EXCEL'E (XLSX) ÇEVİRME ---
        console.log("5. JSON verisi Excel (XLSX) formatına dönüştürülüyor...");

        // Veriyi istediğin sütun başlıklarıyla eşliyoruz (Parcel No kaldırıldı)
        const formattedDataForExcel = parsedJsonData.map(item => ({
            "Gumruk No": item.gumruk_no,
            "Reference": item.reference,
            "Description": item.description,
            "Quantity": item.quantity,
            "AU": item.au
        }));

        // Başlıkların sırasını belirleyerek sayfayı oluşturuyoruz
        const worksheet = xlsx.utils.json_to_sheet(formattedDataForExcel, {
            header: ["Gumruk No", "Reference", "Description", "Quantity", "AU"]
        });

        const workbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(workbook, worksheet, "Paket_Listesi");

        // Excel'i Buffer olarak oluştur
        const excelBuffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
        // ---------------------------------------------------------

        // ADIM 6: Excel dosyasını BunnyCDN'e Yükle
        const outputFileName = `${path.parse(originalName).name}_analysis.xlsx`;
        console.log(`6. Sonuç dosyası (${outputFileName}) BunnyCDN'e yükleniyor...`);

        const outputCdnUrl = await storageService.uploadToCDN(
            excelBuffer,
            outputFileName
        );
        console.log("-> İşlem Tamam! XLSX Linki:", outputCdnUrl);

        // ADIM 7: Başarılı Dönüş
        res.status(200).json({
            success: true,
            message: "Dönüştürme başarılı",
            input_url: inputCdnUrl,
            xlsx_url: outputCdnUrl,
            file_name: outputFileName,
            data: parsedJsonData
        });

    } catch (error) {
        console.error("❌ Controller Hatası:", error.message);
        const errorMessage = error.response?.data?.message || error.message;

        res.status(500).json({
            success: false,
            error: "İşlem sırasında bir hata oluştu: " + errorMessage
        });
    } finally {
        // Geçici dosyaları temizle
        if (tempPdfPath) await fs.unlink(tempPdfPath).catch(() => { });
        if (tempTxtPath) await fs.unlink(tempTxtPath).catch(() => { });
    }
};