const axios = require('axios');
const https = require('https');

exports.processFileWithN8N = async (fileUrl) => {
    try {
        console.log(">> n8n servisine istek atılıyor...");

        const agent = new https.Agent({
            rejectUnauthorized: false
        });

        const response = await axios.post(
            process.env.N8N_WEBHOOK_URL,
            { file_url: fileUrl },
            {
                responseType: 'arraybuffer',
                timeout: 600000,
            }
        );

        console.log(">> n8n'den başarılı yanıt alındı!");
        return response.data;
    } catch (error) {
        // Hatayı daha detaylı görelim
        if (error.response) {
            console.error(`n8n Hatası: ${error.response.status} - ${error.response.statusText}`);
            console.error("Veri:", error.response.data.toString()); // Hata mesajını oku
        } else {
            console.error("n8n İletişim Hatası:", error.message);
        }

        throw new Error("n8n işlemi başarısız");
    }
};