# Dashboard Penilaian NPQEL Ambilan 1/2026

Dashboard interaktif satu halaman (HTML statik) bagi laporan i-Penilaian IAB: skor fasa, sesi latihan, pensyarah, khidmat sokongan, maklum balas terbuka dan profil peserta.

- Data ialah salinan tetap yang diekstrak daripada 106 laporan PDF (BK15a/b/c/d, BK30, BK38, BK-64) pada 12/11/2026.
- Tema V13 dan purata sub-aspek ialah nilai terbitan; lihat tab "Nota data".
- Paparan: https://fadzilma-cloud.github.io/dashboard-penilaian-npqel/

## Muat naik data kumpulan lain

1. Muat turun `templat/Templat_Muat_Naik_NPQEL.xlsx` dan isi mengikut arahan dalam lembaran `arahan` (satu fail = satu kumpulan dan satu tahun).
2. Pada dashboard, tekan **Muat naik data**, pilih fail dan semak ringkasan pengesahan.
3. Tekan **Simpan dan papar**. Tajuk dan semua paparan bertukar kepada kumpulan itu; tukar semula melalui penapis **Kumpulan**.

Fail dibaca dalam pelayar dan disimpan dalam pelayar pengguna sahaja (IndexedDB); tiada data dihantar ke pelayan. `templat/Contoh_Data_SRGH5_2026.xlsx` ialah data SRGH5 dalam format yang sama.

## Alat menukar PDF kepada Excel
Lihat [`alat/README.md`](alat/README.md). Alat ini menukar PDF i-Penilaian secara pukal kepada satu fail xlsx yang boleh dimuat naik ke dashboard.
