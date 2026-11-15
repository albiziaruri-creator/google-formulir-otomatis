# Google Formulir Otomatis

Sistem generator soal dan ujian otomatis berbasis Google Sheets, Google Forms, dan Google Apps Script.

## Deskripsi

Google Formulir Otomatis merupakan aplikasi berbasis Google Workspace untuk membantu guru mengelola bank soal, menyusun kisi-kisi, membuat Google Form ujian secara otomatis, mengelola data ujian, serta menganalisis hasil ujian dan kualitas butir soal.

Sistem dirancang untuk mengurangi pekerjaan administratif guru dan meningkatkan efektivitas evaluasi pembelajaran.

## Fitur Utama

### 1. Manajemen Bank Soal
- Menyimpan dan mengelola bank soal secara terstruktur.
- Mendukung berbagai jenis soal.
- Menyimpan kunci jawaban, poin, pembahasan, dan rubrik penilaian.
- Mendukung deskripsi dan stimulus soal.

### 2. Generator Google Form Otomatis
- Membuat Google Form langsung dari BANK_SOAL.
- Mengatur soal dan pilihan jawaban secara otomatis.
- Mengatur kunci jawaban dan bobot nilai.
- Mendukung impor soal dari dokumen Word (.docx).

### 3. Jenis Soal yang Didukung

| Kode | Jenis Soal |
|---|---|
| PG | Pilihan Ganda |
| PGK | Pilihan Ganda Kompleks |
| MJ | Menjodohkan |
| IS | Isian Singkat |
| UR | Uraian |

### 4. Manajemen Ujian
- Mencatat daftar ujian yang dibuat.
- Menyimpan ID dan URL Google Form.
- Mengelola pemetaan soal terhadap ujian.
- Menghubungkan ujian dengan spreadsheet respons.
- Menyinkronkan bobot soal dengan Google Form.

### 5. Analisis Hasil Peserta
- Mengambil respons langsung dari Google Form.
- Menghitung nilai peserta berdasarkan skor aktual setiap soal.
- Menghitung jumlah jawaban benar dan salah.
- Mengidentifikasi soal yang belum dinilai.
- Menentukan status ketuntasan berdasarkan KKM.

### 6. Analisis Kualitas Butir Soal
- Menghitung tingkat kesukaran soal.
- Menghitung daya pembeda soal.
- Mengelompokkan kualitas soal berdasarkan hasil analisis.
- Mengidentifikasi soal yang perlu ditinjau kembali.

## Struktur Spreadsheet

Sistem menggunakan beberapa sheet utama:

| Sheet | Fungsi |
|---|---|
| BANK_SOAL | Database soal |
| PENGATURAN | Konfigurasi sistem |
| PANDUAN | Petunjuk penggunaan |
| KISI_KISI | Data kisi-kisi soal |
| LOG_SOAL_FORM | Pencatatan proses generator |
| DAFTAR_UJIAN | Database ujian |
| PEMETAAN_SOAL_UJIAN | Pemetaan soal ke Google Form |
| ANALISIS_UJIAN | Data analisis ujian |
| ANALISIS_PESERTA | Hasil analisis peserta |
| ANALISIS_PER_SOAL | Statistik kualitas setiap soal |
| DASHBOARD_UJIAN | Ringkasan hasil ujian |

## Teknologi

- Google Sheets
- Google Forms
- Google Apps Script
- JavaScript
- Google Drive

## Cara Penggunaan

1. Siapkan spreadsheet utama.
2. Isi BANK_SOAL dengan soal yang akan digunakan.
3. Atur konfigurasi pada PENGATURAN.
4. Jalankan generator untuk membuat Google Form.
5. Distribusikan formulir kepada peserta.
6. Setelah ujian selesai, jalankan analisis hasil peserta.
7. Jalankan analisis per soal untuk mengevaluasi kualitas butir soal.

## Status Pengembangan

**Versi 2.0**

Fitur yang telah berhasil dikembangkan dan diuji:

- [x] Manajemen bank soal.
- [x] Impor soal dari dokumen Word.
- [x] Generator Google Form otomatis.
- [x] Dukungan PG, PGK, MJ, IS, dan UR.
- [x] Pemetaan soal ke ujian.
- [x] Sinkronisasi bobot soal.
- [x] Analisis hasil peserta.
- [x] Perhitungan nilai berdasarkan skor Google Form.
- [x] Analisis tingkat kesukaran.
- [x] Analisis daya pembeda.

## Catatan

Aplikasi ini dikembangkan menggunakan Google Apps Script dan berjalan di lingkungan Google Workspace. Pengguna perlu memiliki akses ke Google Sheets, Google Forms, dan Google Drive.

## Pengembang

Google Formulir Otomatis

Proyek pengembangan sistem evaluasi pembelajaran berbasis Google Workspace.
