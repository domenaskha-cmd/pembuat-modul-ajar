const express = require('express');
const app = express();
const path = require('path');

// Middleware untuk membaca JSON dan form data
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// SINKRONISASI ROUTE: Ganti/sesuaikan rute di bawah ini dengan rute asli aplikasi Anda
// Menyediakan folder utama agar file CSS, gambar, atau JS pendukung Anda terbaca
app.use(express.static(path.join(__dirname)));

// Mengarahkan rute utama langsung ke file tampilan index.html
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Contoh rute API jika Anda memilikinya
app.get('/api/test', (req, res) => {
    res.json({ message: "Koneksi API aman dan lancar!" });
});

// Baris paling penting untuk Vercel Serverless
module.exports = app;

// Tetap jalankan listen hanya untuk local development
const PORT = process.env.PORT || 3000;
if (process.env.NODE_ENV !== 'production') {
    app.listen(PORT, () => {
        console.log(`Server running on port ${PORT}`);
    });
}
