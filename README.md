# Word Merger Python

Program berbasis Python untuk menggabungkan banyak file Microsoft Word (`.doc` dan `.docx`) menjadi satu dokumen secara otomatis berdasarkan urutan nomor pada nama file.

Program menggunakan **Microsoft Word Automation melalui `pywin32`**, sehingga proses penggabungan dilakukan langsung oleh Microsoft Word dan lebih baik dalam mempertahankan format dokumen dibandingkan menyalin teks/paragraf secara manual.

## Fitur

- Menggabungkan file `.doc` dan `.docx`.
- Mengurutkan dokumen berdasarkan angka pada nama file.
- Mendukung urutan numerik seperti `1, 2, 3, ..., 9, 10`.
- Menggunakan Microsoft Word secara langsung melalui `pywin32`.
- Mempertahankan format dokumen semaksimal mungkin.
- Mendukung dokumen dengan:
  - Font dan ukuran font
  - Bold, italic, underline
  - Tabel
  - Gambar
  - Kop surat
  - Header dan footer
  - Margin dan layout halaman
  - Section dan format dokumen
- Menambahkan page break antar dokumen.
- Tidak mengubah file sumber.
- Menghasilkan satu file `HASIL_GABUNGAN.docx`.

## Cara Kerja

```text
File Word
   │
   ├── 1.doc
   ├── 2.doc
   ├── 3.docx
   ├── ...
   └── 10.doc
          │
          ▼
   Baca nama file
          │
          ▼
   Urutkan secara numerik
          │
          ▼
   Microsoft Word Automation
          │
          ▼
   InsertFile()
          │
          ▼
   Page Break
          │
          ▼
   HASIL_GABUNGAN.docx
```

## Persyaratan

- Windows
- Python 3.9 atau lebih baru
- Microsoft Word Desktop terinstall
- Microsoft Word dapat dijalankan secara normal
- `pywin32`

> **Catatan:** Program ini tidak menggunakan LibreOffice. Microsoft Word diperlukan karena proses penggabungan dilakukan melalui Word Automation.

## Instalasi

### 1. Clone repository

```bash
git clone [https://github.com/USERNAME/word-merger-python.git](https://github.com/H4nk/Gabung-File-Word.git)
cd word-merger-python
```
### 2. Buat virtual environment

```powershell
python -m venv .venv
```

Tidak wajib melakukan `activate`. Untuk menghindari masalah PowerShell Execution Policy, Python dari virtual environment dapat dipanggil langsung.

### 3. Install dependency

```powershell
.\.venv\Scripts\python.exe -m pip install pywin32
```

Atau jika menggunakan `requirements.txt`:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Struktur Folder

Contoh:

```text
word-merger-python/
│
├── gabung_word_msword.py
├── requirements.txt
├── README.md
│
├── 1.doc
├── 2.doc
├── 3.doc
├── 4.docx
├── 5.doc
├── ...
└── 10.doc
```

File Word dapat menggunakan nama yang lebih deskriptif, misalnya:

```text
1 Surat Permohonan.doc
2 Surat Pernyataan.doc
3 Surat Tugas.docx
4 Daftar Peserta.doc
5 Lampiran.docx
```

Program mengambil **angka pertama** dari nama file sebagai dasar pengurutan.

## Menjalankan Program

Jalankan:

```powershell
.\.venv\Scripts\python.exe gabung_word_msword.py
```

Atau jika Python tersedia secara global:

```powershell
python gabung_word_msword.py
```

Program akan menghasilkan:

```text
HASIL_GABUNGAN.docx
```

## Contoh Output

Jika folder berisi:

```text
1.doc
2.doc
3.docx
4.doc
5.docx
10.doc
```

program akan memproses:

```text
1.doc
2.doc
3.docx
4.doc
5.docx
10.doc
```

Bukan:

```text
1.doc
10.doc
2.doc
3.docx
...
```

## Contoh Tampilan Terminal

```text
========================================================================
GABUNG FILE WORD MENGGUNAKAN MICROSOFT WORD
========================================================================
Folder : D:\GabungWord
Output : D:\GabungWord\HASIL_GABUNGAN.docx

Urutan dokumen:
------------------------------------------------------------------------
001.     1  1.doc
002.     2  2.doc
003.     3  3.docx
004.     4  4.doc
005.     5  5.docx
006.    10  10.doc

[1/6] Membuka: 1.doc
[2/6] Menggabungkan: 2.doc
[3/6] Menggabungkan: 3.docx
[4/6] Menggabungkan: 4.doc
[5/6] Menggabungkan: 5.docx
[6/6] Menggabungkan: 10.doc

Menyimpan hasil...

========================================================================
SELESAI
========================================================================
Jumlah file : 6
Hasil       : D:\GabungWord\HASIL_GABUNGAN.docx

File asli tidak diubah.
Penggabungan dilakukan langsung oleh Microsoft Word.
========================================================================
```

## Mengapa Menggunakan Microsoft Word Automation?

Metode sederhana menggunakan `python-docx` biasanya membuat dokumen baru dan menyalin paragraf, sehingga beberapa properti dokumen dapat berubah.

Project ini menggunakan:

```python
import win32com.client
```

dan Microsoft Word:

```python
insertion_range.InsertFile(
    str(file.resolve())
)
```

Dengan demikian Microsoft Word sendiri yang menangani struktur dokumen saat proses penggabungan.

## Catatan Format

Walaupun program menggunakan Microsoft Word untuk mempertahankan format, hasil akhir tetap bergantung pada:

1. Font yang tersedia di komputer.
2. Versi Microsoft Word.
3. Struktur section dokumen.
4. Format dokumen sumber.
5. Link atau objek eksternal yang terdapat dalam dokumen.

Untuk hasil paling konsisten, gunakan font yang sama dengan dokumen sumber dan pastikan font tersebut terinstall pada Windows.

## Keamanan File

Program tidak menghapus atau mengubah file Word sumber.

File sumber:

```text
1.doc
2.doc
3.docx
```

tetap berada di folder asal.

Program hanya membuat:

```text
HASIL_GABUNGAN.docx
```

## Troubleshooting

### `ModuleNotFoundError: No module named 'win32com'`

Install:

```powershell
.\.venv\Scripts\python.exe -m pip install pywin32
```

### Microsoft Word tidak ditemukan

Pastikan Microsoft Word Desktop terinstall dan dapat dibuka secara normal.

Program ini **tidak menggunakan Word Online**.

### File tidak sesuai urutan

Pastikan nama file memiliki angka sebagai nomor urutan:

```text
1.doc
2.doc
3.doc
10.doc
```

Hindari nama seperti:

```text
Surat A.doc
Surat B.doc
```

jika urutan numerik diperlukan.

### Word terbuka tetapi proses gagal

Tutup seluruh dokumen Microsoft Word yang sedang terbuka dan jalankan kembali program.

## Use Case

Project ini cocok untuk:

- Penggabungan surat administrasi.
- Penggabungan dokumen peserta.
- Penggabungan laporan.
- Penggabungan lampiran.
- Penggabungan berkas kegiatan.
- Penggabungan dokumen akademik.
- Penggabungan dokumen administrasi pemerintahan.
- Penggabungan dokumen berdasarkan nomor urut.

## Lisensi

Silakan tentukan lisensi project sesuai kebutuhan. Untuk project open-source umum, Anda dapat menggunakan **MIT License**.

## Author

**Harry Setya Hadi**

Python Programmer & Researcher

**Email:** xmoensen@gmail.com
**youtube:** https://www.youtube.com/@HarrySetyaHadi

---

Jika project ini bermanfaat, silakan berikan ⭐ pada repository GitHub.
