const express = require('express');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle } = require('docx');
const path = require('path');
const app = express();

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

app.post('/generate-modul', async (req, res) => {
    try {
        const { kurikulum, mapel, kelas, materi, alokasi_waktu } = req.body;

        // Pengaturan garis batas (Border) resmi & Estetik
        const thinBorder = { style: BorderStyle.SINGLE, size: 6, color: "CBD5E0" };
        const thickTopBorder = { style: BorderStyle.SINGLE, size: 18, color: "1A365D" };

        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } // Standar Margin 2.54 cm
                    }
                },
                children: [
                    // ==================== JUDUL UTAMA MODUL ====================
                    new Paragraph({
                        text: `MODUL AJAR KE-1: ${materi.toUpperCase()}`,
                        heading: HeadingLevel.HEADING_1,
                        alignment: AlignmentType.CENTER,
                    }),
                    new Paragraph({
                        text: `IMPLEMENTASI ${kurikulum.toUpperCase()} - TAHUN AJARAN 2026/2027`,
                        heading: HeadingLevel.HEADING_3,
                        alignment: AlignmentType.CENTER,
                    }),
                    new Paragraph({ text: "" }),

                    // ==================== 1. INFORMASI UMUM ====================
                    new Paragraph({ text: "1. INFORMASI UMUM", heading: HeadingLevel.HEADING_2 }),
                    
                    new Paragraph({ text: "A. IDENTITAS MODUL", heading: HeadingLevel.HEADING_3 }),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder, insideHorizontal: thinBorder },
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({ width: { size: 35, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: " Nama Penyusun", bold: true })] })] }),
                                    new TableCell({ width: { size: 65, type: WidthType.PERCENTAGE }, children: [new Paragraph({ text: " Tim Pengembang Perangkat Ajar Utama" })] }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: " Institusi / Madrasah", bold: true })] })] }),
                                    new TableCell({ children: [new Paragraph({ text: " Satuan Pendidikan Setempat" })] }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: " Mata Pelajaran / Elemen", bold: true })] })] }),
                                    new TableCell({ children: [new Paragraph({ text: ` ${mapel} / Bidang Kajian Utama` })] }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: " Kelas / Jenjang Sekolah", bold: true })] })] }),
                                    new TableCell({ children: [new Paragraph({ text: ` ${kelas} / Tingkat Menengah` })] }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: " Alokasi Waktu", bold: true })] })] }),
                                    new TableCell({ children: [new Paragraph({ text: ` ${alokasi_waktu} (Sesuai jam pelajaran unit kerja)` })] }),
                                ]
                            }),
                        ],
                    }),
                    new Paragraph({ text: "" }),

                    new Paragraph({ text: "B. KOMPETENSI AWAL", heading: HeadingLevel.HEADING_3 }),
                    new Paragraph({ text: `Peserta didik telah memiliki pengetahuan dasar/prasyarat umum serta ketrampilan awal terkait asas operasional sebelum mendalami bab ${materi}.` }),
                    new Paragraph({ text: "" }),

                    new Paragraph({ text: "C. PROFIL PELAJAR PANCASILA & PPRA", heading: HeadingLevel.HEADING_3 }),
                    new Paragraph({ children: [new TextRun({ text: "• Beriman dan bertaqwa kepada TYME serta berakhlak mulia; ", bold: true }), new TextRun(`Peserta didik menunjukkan perilaku bersyukur dengan mengucap Hamdalah karena bisa belajar topik ${materi}.`)] }),
                    new Paragraph({ children: [new TextRun({ text: "• Gotong Royong; ", bold: true }), new TextRun("Peserta didik menunjukkan perilaku kerjasama yang harmonis selama aktivitas diskusi kelompok.")] }),
                    new Paragraph({ children: [new TextRun({ text: "• Mandiri; ", bold: true }), new TextRun("Peserta didik menunjukkan perilaku percaya diri serta bertanggung jawab atas proses belajarnya.")] }),
                    new Paragraph({ children: [new TextRun({ text: "• Bernalar Kritis; ", bold: true }), new TextRun("Peserta didik menunjukkan perilaku berani bertanya dan berani memberi masukan konstruktif.")] }),
                    new Paragraph({ children: [new TextRun({ text: "• PPRA - Berkeadaban (ta'addub); ", bold: true }), new TextRun("Peserta didik menunjukkan perilaku santun dan tidak sombong jika sudah menguasai materi pembelajaran.")] }),
                    new Paragraph({ children: [new TextRun({ text: "• PPRA - Keteladanan (qudwah); ", bold: true }), new TextRun("Peserta didik menunjukkan perilaku yang baik dan layak dijadikan contoh bagi rekan lainnya.")] }),
                    new Paragraph({ text: "" }),

                    new Paragraph({ text: "D. SARANA DAN PRASARANA", heading: HeadingLevel.HEADING_3 }),
                    new Paragraph({ text: "• MEDIA: Slides Presentasi Interaktif, Video Dokumenter Kontekstual, Simulator Sistem Digital.\n• ALAT/BAHAN: Laptop, LCD Proyektor, Papan Tulis, Kertas Plano, dan Spidol Warna.\n• SUMBER BELAJAR: Buku Panduan Guru, Buku Siswa Utama Kurikulum Reguler, dan Lembar Kasus Riil Lingkungan Sekitar." }),
                    new Paragraph({ text: "" }),

                    new Paragraph({ text: "E. TARGET PESERTA DIDIK", heading: HeadingLevel.HEADING_3 }),
                    new Paragraph({ text: "• Peserta didik umum/reguler: Umum, tidak ada kesulitan dalam mencerna dan memahami materi ajar.\n• Peserta didik berkemampuan cepat: Mencerna dengan sangat cepat dan mampu mencapai keterampilan berpikir tingkat tinggi (HOTS)." }),
                    new Paragraph({ text: "" }),

                    new Paragraph({ text: "F. MODEL DAN METODE PEMBELAJARAN", heading: HeadingLevel.HEADING_3 }),
                    new Paragraph({ text: "Pendekatan Saintifik dengan model Blended Learning menggunakan kombinasi metode ceramah, tanya jawab, diskusi terstruktur, inkuiri (pencarian mandiri), resitasi (penugasan), dan simulasi aksi." }),
                    new Paragraph({ text: "" }),

                    // ==================== 2. KOMPONEN INTI ====================
                    new Paragraph({ text: "2. KOMPONEN INTI", heading: HeadingLevel.HEADING_2 }),
                    
                    new Paragraph({ text: "A. TUJUAN PEMBELAJARAN (Unsur ABCD)", heading: HeadingLevel.HEADING_3 }),
                    new Paragraph({ text: `Melalui penerapan pendekatan saintifik dan metode diskusi terstruktur (Condition), peserta didik reguler maupun berkemampuan cepat (Audience) dapat mengidentifikasi, menguraikan struktur, serta memecahkan problem aktual terkait topik ${materi} (Behaviour) dengan baik, benar, dan penuh rasa percaya diri (Degree).` }),
                    
                    new Paragraph({ text: "B. PEMAHAMAN BERMAKNA", heading: HeadingLevel.HEADING_3 }),
                    new Paragraph({ text: `Peserta didik memperoleh manfaat nyata berupa kemampuan berpikir logis-struktural untuk menganalisis fenomena ${materi} di dunia nyata, serta mampu mengoptimalkan komponen tersebut guna memecahkan masalah sistemik di kehidupan sehari-hari.` }),
                    
                    new Paragraph({ text: "C. PERTANYAAN PEMANTIK", heading: HeadingLevel.HEADING_3 }),
                    new Paragraph({ text: `1. Apa yang akan terjadi di lingkungan sekitar kita apabila aspek fungsional dari ${materi} ini mengalami kegagalan total?\n2. Bagaimana cara paling efisien yang bisa kita tempuh untuk mengoptimalkan kinerja materi ini dengan keterbatasan sarana yang ada?` }),
                    new Paragraph({ text: "" }),

                    new Paragraph({ text: "D. KEGIATAN PEMBELAJARAN DETAIL", heading: HeadingLevel.HEADING_3 }),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: { top: thickTopBorder, bottom: thinBorder, left: thinBorder, right: thinBorder, insideHorizontal: thinBorder },
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({ width: { size: 25, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: "Fase Aktivitas", bold: true })] })] }),
                                    new TableCell({ width: { size: 75, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: "Langkah Instruksional Operasional (Berbasis Saintifik & Aktif)", bold: true })] })] }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({ children: [new Paragraph({ text: "Pendahuluan\n(15 Menit)" })] }),
                                    new TableCell({ children: [new Paragraph({ text: "• Guru membuka kelas dengan salam hangat, berdoa bersama, dan melakukan presensi berkala.\n• Guru menyiapkan fisik-psikis siswa melalui kegiatan yel-yel motivasi atau nyanyian penyemangat.\n• Guru melakukan kegiatan Asesmen Diagnostik Kognitif & Non-Kognitif singkat.\n• Guru memberikan pertanyaan pemantik dan menjelaskan kompetensi akhir yang wajib dicapai bersama." })] }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({ children: [new Paragraph({ text: "Inti\n(60 Menit)" })] }),
                                    new TableCell({ children: [new Paragraph({ text: `• Mengamati: Peserta didik mengamati paparan media visual dan lembaran data kasus kontekstual terkait topik ${materi}.\n• Menanya: Peserta didik mengajukan pertanyaan kritis terkait poin-poin permasalahan yang belum dipahami.\n• Mengeksplorasi: Peserta didik membentuk kelompok, mengumpulkan fakta, dan meneliti alternatif pemecahan masalah.\n• Mengasosiasi: Peserta didik berdiskusi kelompok mengolah data, menganalisis hubungan sebab-akibat, dan merumuskan draf solusi.\n• Mengomunikasikan: Peserta didik mempresentasikan hasil diskusi di depan kelas dan menerima tanggapan kelompok lain.` })] }),
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({ children: [new Paragraph({ text: "Penutup\n(15 Menit)" })] }),
                                    new TableCell({ children: [new Paragraph({ text: "• Guru bersama peserta didik melakukan refleksi dan menarik kesimpulan komprehensif atas hasil belajar.\n• Guru mengadakan aktivitas Asesmen Formatif akhir proses.\n• Guru menginformasikan materi tindak lanjut untuk pertemuan pekan depan, lalu menutup pembelajaran dengan doa dan salam." })] }),
                                ]
                            }),
                        ],
                    }),
                    new Paragraph({ text: "" }),

                    new Paragraph({ text: "E. ASESMEN (INSTRUMEN LENGKAP)", heading: HeadingLevel.HEADING_3 }),
                    new Paragraph({ text: "1. Asesmen Diagnostik: Menjawab 3 pertanyaan kunci penilai kesiapan mental dan prasyarat dasar sebelum bab dimulai.\n2. Asesmen Formatif: Pengamatan keaktifan performa diskusi, penilaian lisan di tengah kelas, dan tes tertulis akhir sesi.\n3. Asesmen Ranah Sikap: Menggunakan lembar observasi jurnal harian pencatatan Profil Pelajar Pancasila.\n4. Asesmen Sumatif: Pemberian Tugas Rumah (PR) terstruktur berbasis proyek/studi analisis kasus mandiri.\n5. Penilaian Keterampilan: Unjuk kerja instrumen performa presentasi dan penyusunan draf lembar kerja." }),
                    new Paragraph({ text: "" }),

                    new Paragraph({ text: "F. PENGAYAAN DAN REMEDIAL", heading: HeadingLevel.HEADING_3 }),
                    new Paragraph({ text: "• Remedial: Diberikan bimbingan khusus individual atau tutor sebaya bagi peserta didik yang belum mencapai target ketuntasan kompetensi.\n• Pengayaan: Diberikan kepada peserta didik berkemampuan cepat berupa pengerjaan latihan soal-soal penalaran analisis tingkat tinggi (AKM/HOTS)." }),
                    new Paragraph({ text: "" }),

                    // ==================== 3. LAMPIRAN ====================
                    new Paragraph({ text: "3. LAMPIRAN PENDUKUNG", heading: HeadingLevel.HEADING_2 }),
                    
                    new Paragraph({ text: "A. LEMBAR KERJA PESERTA DIDIK (LKPD)", heading: HeadingLevel.HEADING_3 }),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: { top: thickTopBorder, bottom: thickTopBorder, left: thickTopBorder, right: thickTopBorder },
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({
                                        margin: { top: 200, bottom: 200, left: 200, right: 200 },
                                        children: [
                                            new Paragraph({ text: "LEMBAR KERJA SISWA MANDIRI & KELOMPOK", heading: HeadingLevel.HEADING_4, alignment: AlignmentType.CENTER }),
                                            new Paragraph({ text: "Petunjuk Kerja:\n1. Bacalah doa setiap mengawali dan mengakhiri belajar!\n2. Cermatilah rangkuman materi pendukung yang tersedia pada lampiran.\n3. Kerjakan tugas, 10 soal pilihan ganda, dan 5 soal uraian di bawah ini secara teliti dan kolaboratif!\n4. Kumpulkan hasil lembar kerja tepat waktu sebelum durasi habis!" }),
                                            new Paragraph({ text: "" }),
                                            
                                            new Paragraph({ children: [new TextRun({ text: "BANK SOAL PILIHAN GANDA (10 BUTIR HOTS)", bold: true })] }),
                                            new Paragraph({ text: `1. Mengapa penguasaan terhadap konsep ${materi} dipandang sangat vital dalam aplikasi di kehidupan nyata?\n   a. Membatasi kreativitas berpikir\n   b. Mengoptimalkan efisiensi pemecahan masalah secara logis-struktural\n   c. Mempercepat tugas tanpa peduli prosedur\n   d. Mengurangi interaksi antar komponen` }),
                                            new Paragraph({ text: `2. Komponen manakah yang bertindak sebagai fondasi utama penopang sistem ${materi}?\n   a. Komponen pelengkap luar\n   b. Struktur inti yang terintegrasi penuh\n   c. Dokumentasi sekunder\n   d. Opini subjektif` }),
                                            new Paragraph({ text: "3. Jika ditemukan kendala operasional pada sistem, langkah saintifik pertama yang harus diambil adalah...\n   a. Mengabaikan masalah\n   b. Mengidentifikasi akar penyebab secara objektif\n   c. Mengubah total tanpa rencana\n   d. Menutup seluruh sistem" }),
                                            new Paragraph({ text: "4. Karakteristik utama dari output pemecahan masalah yang baik dan berhasil guna tinggi adalah...\n   a. Kaku dan permanen\n   b. Fleksibel dan adaptif terhadap dinamika riil\n   c. Berbiaya mahal tanpa hasil\n   d. Bergantung pada manual satu arah" }),
                                            new Paragraph({ text: `5. Penerapan bab materi ${materi} secara kontekstual bertujuan utama untuk melatih siswa agar...\n   a. Berpikir instan dan subjektif\n   b. Berpikir analitis, kritis, serta berbasis data sahih\n   c. Menghafal seluruh isi teks\n   d. Mengabaikan kerja kelompok` }),
                                            new Paragraph({ text: "6. Instrumen penunjang keadilan objektivitas penilaian guru di dalam kelas adalah...\n   a. Catatan tebakan informal\n   b. Indikator ketercapaian dan rubrik skala performa\n   c. Soal ujian tanpa kunci jawaban\n   d. Kedekatan personal murid" }),
                                            new Paragraph({ text: "7. Mengapa materi ajar wajib dikaitkan dengan fenomena riil lingkungan sekitar?\n   a. Agar proses belajar terkesan rumit\n   b. Agar keterkaitan konsep teori dengan dunia nyata terlihat jelas\n   c. Untuk membuang alokasi waktu\n   d. Mengurangi keaktifan murid" }),
                                            new Paragraph({ text: "8. Tantangan global terbesar yang memengaruhi dinamika perkembangan materi ini di era modern adalah...\n   a. Minimnya buku cetak lama\n   b. Laju konvergensi teknologi baru yang menuntut pembaruan konseptual\n   c. Pengurangan jam sekolah\n   d. Perubahan kalender libur" }),
                                            new Paragraph({ text: "9. Dimensi Profil Pancasila yang dilatih saat siswa melakukan debat pemecahan masalah kelompok adalah...\n   a. Egoisme individu\n   b. Gotong royong, kolaborasi, dan menghargai perbedaan\n   c. Ketergantungan penuh\n   d. Penerimaan pasif tekstual" }),
                                            new Paragraph({ text: "10. Seseorang dikatakan telah mencapai ketuntasan kompetensi belajar apabila...\n    a. Menyerahkan lembar jawaban paling cepat\n    b. Menunjukkan pencapaian utuh pada asesmen formatif, diskusi, dan performa makna\n    c. Menyalin penuh draf teman\n    d. Memiliki nilai sama dengan rata-rata lama" }),
                                            new Paragraph({ text: "" }),

                                            new Paragraph({ children: [new TextRun({ text: "BANK SOAL URAIAN (5 BUTIR ANALITIS)", bold: true })] }),
                                            new Paragraph({ text: `1. Uraikan analisis kritis Anda mengenai urgensi penguasaan konsep ${materi} di era kemajuan modern saat ini!\nJawab:\n\n` }),
                                            new Paragraph({ text: "2. Jabarkan 3 dampak negatif konkret yang berpotensi muncul jika pilar keilmuan ini tidak diterapkan secara konsisten!\nJawab:\n\n" }),
                                            new Paragraph({ text: "3. Rumuskan strategi pemecahan masalah kreatif versi tim Anda jika unit sekolah mengalami keterbatasan sarana laboratorium!\nJawab:\n\n" }),
                                            new Paragraph({ text: "4. Mengapa variasi instrumen evaluasi bernuansa HOTS wajib diintegrasikan dalam modul ajar ini? Hubungkan dengan daya saing global!\nJawab:\n\n" }),
                                            new Paragraph({ text: "5. Telaah dan rumuskan hubungan timbal balik yang ideal antara penguasaan aspek teori dengan kemahiran praktik nyata di lapangan!\nJawab:\n\n" }),
                                        ]
                                    })
                                ]
                            })
                        ]
                    }),
                    new Paragraph({ text: "" }),

                    new Paragraph({ text: "B. BAHAN BACAAN GURU & PESERTA DIDIK (RANGKUMAN)", heading: HeadingLevel.HEADING_3 }),
                    new Paragraph({ text: `Kajian keilmuan mengenai ${materi} merupakan pilar instruksional penting yang mengatur integrasi fungsional sistem belajarnya. Pemahaman konseptual yang kokoh memandu pendidik dan siswa untuk mampu merancang solusi rekayasa, mitigasi risiko galat, serta melakukan akselerasi inovasi secara efektif. Di era modern, penguasaan materi ini dikolaborasikan langsung dengan teknologi digital demi melahirkan output yang kompetitif.` }),
                    new Paragraph({ text: "" }),

                    new Paragraph({ text: "C. GLOSARIUM (ALFABETIKAL)", heading: HeadingLevel.HEADING_3 }),
                    new Paragraph({ text: `• ABCD = Audience (Siswa), Behaviour (Perilaku), Condition (Kondisi), Degree (Tingkat Ketepatan).\n• Blended Learning = Model pembelajaran kombinasi tatap muka langsung dan daring.\n• HOTS = Higher Order Thinking Skills (Kemampuan berpikir aras tinggi meliputi analisis, evaluasi, kreativitas).\n• ${materi.toUpperCase()} = Domain topik bahasan instruksional utama pada perangkat modul ini.\n• PPRA = Profil Pelajar Rahmatan Lil Alamin (Nilai pembentukan karakter penunjang kesantunan umat).` }),
                    new Paragraph({ text: "" }),

                    new Paragraph({ text: "D. DAFTAR PUSTAKA (Format NATABUKOPEN)", heading: HeadingLevel.HEADING_3 }),
                    new Paragraph({ text: "Kementerian Pendidikan, Kebudayaan, Riset, dan Teknologi Republik Indonesia. (2026). Buku Panduan Guru dan Buku Siswa Utama Kurikulum Nasional. Jakarta: Pusat Kurikulum dan Perbukuan." }),
                    new Paragraph({ text: "" }),
                    
                    // KUNCI JAWABAN RESMI
                    new Paragraph({ text: "KUNCI JAWABAN RESMI (PEGANGAN GURU)", heading: HeadingLevel.HEADING_4 }),
                    new Paragraph({ text: "• Pilihan Ganda: 1.B | 2.B | 3.B | 4.B | 5.B | 6.B | 7.B | 8.B | 9.B | 10.B" }),
                    new Paragraph({ text: "• Uraian: 1. Melatih penalaran objektif berbasis fakta nyata. 2. Penurunan produktivitas, kegagalan sistemik, ketidakmampuan adaptasi. 3. Substitusi media digital interaktif atau pemanfaatan barang lingkungan sekitar. 4. Menyiapkan SDM kritis berskala daya saing global. 5. Teori yang matang memandu ketepatan praktik, dan kendala praktik menyempurnakan khazanah teori." }),
                    new Paragraph({ text: "" }),

                    // Tanda Tangan Pengesahan Resmi
                    new Paragraph({ text: "Mengetahui,\nKepala Madrasah / Sekolah                    Guru Mata Pelajaran\n\n\n\n_____________________                    _____________________\nNIP.                                         NIP.", alignment: AlignmentType.LEFT })
                ],
            }],
        });

        const buffer = await Packer.toBuffer(doc);
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename=Modul_Ajar_Sistematika_Lengkap_${mapel.replace(/\s+/g, '_')}.docx`);
        res.send(buffer);

    } catch (error) {
        console.error(error);
        res.status(500).send("Gagal memproses dokumen.");
    }
});

app.listen(3000, () => {
    console.log('=== SERVER MODUL AJAR AKTIF ===');
    console.log('Akses aplikasi web via browser di alamat: http://localhost:3000');
});