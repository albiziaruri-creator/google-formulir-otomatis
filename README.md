# SISTEM GENERATOR SOAL UJIAN

Sistem pembuatan dan pengelolaan Google Formulir ujian secara otomatis berbasis Google Sheets dan Google Apps Script.

## Fitur

### Pengelolaan Bank Soal
- Penyimpanan bank soal secara terpusat menggunakan Google Sheets.
- Mendukung berbagai jenis soal:
  - Pilihan Ganda (PG).
  - Pilihan Ganda Kompleks (PGK).
  - Menjodohkan (MJ).
  - Isian Singkat (IS).
  - Uraian (UR).
- Import soal dari dokumen Microsoft Word.

### Pembuatan Google Formulir
- Pembuatan Google Formulir ujian secara otomatis.
- Pemilihan soal berdasarkan ID soal.
- Pemilihan soal secara acak.
- Pemilihan soal berdasarkan kriteria:
  - Mata pelajaran.
  - Kelas.
  - Materi.
  - Level kognitif.
  - Level Bloom.
- Pengacakan urutan soal.
- Pembuatan kisi-kisi soal secara otomatis.

### Penambahan Soal ke Formulir yang Sudah Ada
- Menambahkan soal dari bank soal ke Google Formulir yang sudah dibuat.
- Mendukung jenis soal PG, PGK, MJ, IS, dan UR.
- Pemilihan soal berdasarkan ID, secara acak, atau berdasarkan kriteria.
- Mempertahankan soal dan respons yang sudah ada pada formulir tujuan.
- Pencatatan riwayat penambahan soal secara otomatis pada sheet `LOG_SOAL_FORM`.
- Pemeriksaan riwayat berdasarkan ID soal dan ID formulir untuk membantu mencegah penambahan soal yang sama secara berulang.

## Teknologi

- Google Sheets.
- Google Apps Script.
- Google Forms.
- Google Drive API.

## Struktur Proyek

```text
SISTEM-GENERATOR-SOAL-UJIAN/
│
├── apps-script/
│   └── Code.gs
│
├── spreadsheet-template/
│   └── Template Spreadsheet
│
├── dokumentasi/
│   └── Panduan Penggunaan
│
└── README.md
```

## Struktur Spreadsheet

Spreadsheet utama menggunakan beberapa sheet berikut:

| Nama Sheet | Fungsi |
|---|---|
| `BANK_SOAL` | Menyimpan kumpulan soal |
| `PENGATURAN` | Menyimpan pengaturan sistem |
| `PANDUAN` | Petunjuk penggunaan |
| `KISI_KISI` | Menyimpan hasil kisi-kisi soal |
| `LOG_SOAL_FORM` | Mencatat riwayat penambahan soal ke Google Formulir |

## Catatan Penggunaan

- Sistem dijalankan melalui Google Sheets yang terhubung dengan Google Apps Script.
- Import dokumen Word memerlukan Google Drive API.
- Fitur penambahan soal menggunakan Google Formulir yang sudah diaktifkan sebagai kuis.
- Pastikan ID soal pada bank soal bersifat unik agar pencatatan dan pemeriksaan duplikasi dapat berjalan dengan baik.

## Catatan Repository

Repository ini digunakan sebagai cadangan kode program, template spreadsheet, dokumentasi, dan riwayat pengembangan proyek.

Jangan menyimpan kata sandi, token akses, kredensial, atau data pribadi siswa di dalam repository.
