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

## Muat naik terus daripada PDF
Dalam modal **Muat naik data**, bahagian "Atau tukar terus daripada PDF i-Penilaian": pilih semua PDF (atau folder) satu kumpulan, semak fasa program, isi kumpulan dan tahun, kemudian **Simpan dan papar**. PDF dibaca dalam pelayar (pdf.js, folder `pdfjs/`) dan tidak dihantar ke mana-mana. Fasa dikesan daripada kandungan PDF atau nama folder; jika tidak pasti, anda perlu memilihnya.

## Alat Python (pilihan)
Lihat [`alat/README.md`](alat/README.md) untuk menukar PDF kepada fail xlsx di luar dashboard.

## Log masuk Pengurus (Supabase Auth)
Butang **Muat naik data** disembunyikan secara lalai dan hanya muncul selepas Pengurus log masuk. Terdapat **Log masuk**, **Log keluar** dan **Lupa kata laluan** (e-mel pautan tetapan semula).

### Persediaan (sekali sahaja)
1. Cipta projek percuma di https://supabase.com.
2. **Authentication > Sign In / Providers > Email**: aktifkan Email; matikan **Allow new users to sign up** (supaya hanya akaun yang anda cipta boleh log masuk).
3. **Authentication > Users > Add user**: cipta akaun setiap Pengurus (e-mel dan kata laluan awal; pilih *Auto Confirm User*). Pengurus tambahan ditambah di sini juga.
4. **Authentication > URL Configuration**: tetapkan *Site URL* kepada alamat dashboard (contoh `https://fadzilma-cloud.github.io/dashboard-penilaian-npqel/`) dan tambah alamat yang sama pada *Redirect URLs*. Tanpa ini pautan Lupa kata laluan tidak akan kembali ke dashboard.
5. **Project Settings > API**: salin *Project URL* dan kunci **anon public**. Kunci ini memang awam; jangan sekali-kali masukkan kunci `service_role`.
6. Dalam `index.html`, isi `window.NPQ_AUTH_CONFIG = {url: '...', anonKey: '...'}` (cari `NPQ_AUTH_CONFIG`).

### Had keselamatan
Dashboard ialah laman statik dan data paparan sedia ada dalam fail HTML, jadi log masuk ini mengawal akses kepada fungsi muat naik, bukan kerahsiaan data paparan. Data yang dimuat naik hanya disimpan dalam pelayar pengurus itu sendiri (IndexedDB), bukan dikongsi dengan pelawat lain.
