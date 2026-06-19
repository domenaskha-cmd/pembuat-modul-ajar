const express = require('express');
const app = express();
const path = require('path');

// Middleware untuk membaca JSON dan form data
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// SINKRONISASI ROUTE: Ganti/sesuaikan rute di bawah ini dengan rute asli aplikasi Anda
app.get('/', (req, res) => {
    res.send('<h1>Server Backend Berhasil Berjalan!</h1><p>Jika aplikasi Anda menggunakan HTML statis, pastikan file tersebut ditaruh di folder public.</p>');
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
