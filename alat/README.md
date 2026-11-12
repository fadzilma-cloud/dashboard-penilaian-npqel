# Alat PDF ke Excel (npqel_pdf_ke_xlsx.py)

Menukar laporan PDF i-Penilaian bagi satu kumpulan kepada **satu fail xlsx** dalam format muat naik dashboard.

**Keperluan:** Python 3.9+, `pip install pdfplumber openpyxl`

## Antara muka web setempat (muat naik pukal)
```
python npqel_pdf_ke_xlsx.py --web
```
Buka http://127.0.0.1:8765, seret semua PDF (atau folder/ZIP), semak fasa program yang dikesan, isi kumpulan dan tahun, kemudian muat turun xlsx. Pelayan hanya mendengar pada komputer sendiri; PDF tidak disimpan.

## Baris arahan
```
python npqel_pdf_ke_xlsx.py folder_pdf/ -k SRGH6 -t 2027 -o Data_SRGH6_2027.xlsx
python npqel_pdf_ke_xlsx.py data.zip -k SRGH6 -t 2027 --fasa 2027-05-11=2
```

## Borang dikenali (dikesan daripada kandungan PDF)
BK15a, BK15b, BK15c, BK15d, BK30, BK38, BK64. PDF lain dilangkau dan dilaporkan.

## Cara fasa ditentukan
1. Pilihan pengguna (`--fasa TARIKH_MULA=N` atau pilihan pada jadual program).
2. Teks "Fasa N" dalam BK38.
3. "Fasa N" pada nama folder/fail.
4. Jika tiada, susunan tarikh hanya dicadangkan dan **fail tidak dijana sehingga disahkan**.

## Had dan semakan
- Satu fail muat naik = satu kumpulan = satu tahun, maksimum tiga program (Fasa 1-3).
- BK64: hanya jawatan, gred, negeri dan jenis tempat bertugas dikira secara agregat. Nama, No. KP, telefon dan e-mel tidak dibaca ke dalam fail hasil.
- Lajur `modul` dan `tema_v13` dikosongkan; dashboard menerbitkannya sendiri.
- Fail BK15a yang kandungannya sama dengan fail lain dianggap pendua dan dilaporkan.
- Format PDF mesti serupa dengan laporan i-Penilaian 2026. Jika bahagian tidak dapat dibaca, alat melaporkannya (bukan meneka). Semak amaran dan bandingkan beberapa nilai dengan PDF asal sebelum digunakan.
- Diuji: 106 PDF SRGH5 2026 menghasilkan data yang sepadan dengan dashboard sedia ada pada semua lembaran.
