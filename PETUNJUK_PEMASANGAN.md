# Paket Integrasi Portal PKPPS Hidayatullah Karimun V1.41

## Isi paket
- `index.html` dan `portal.html`: antarmuka Portal berdasarkan sumber V1.39 yang diberikan, diarahkan ke API Vercel.
- `Code.gs`: backend Apps Script terbaru yang Anda konfirmasi aktif; dispatcher `doPost` dan daftar aksi dipertahankan.
- `api/portal.js`: penghubung aman sisi server Vercel ke Apps Script.
- `ppdb.html`, `generator.html`: halaman pendamping dari paket integrasi sebelumnya.
- `vercel.json`: rute portal, PPDB, dan generator.

## Konfigurasi Vercel
1. Unggah isi folder ini ke repository Vercel.
2. Tambahkan Environment Variable `GAS_WEB_APP_URL` berisi URL Web App Apps Script aktif yang berakhiran `/exec`.
3. Deploy Preview terlebih dahulu. Uji login, hak akses, absensi santri, rapor, PPDB, dan modul lainnya.
4. Setelah lulus uji, hubungkan domain `pkpps.hidayatullahkarimun.com` pada proyek Vercel yang sama. Jangan mengubah domain utama.

## Apps Script
`Code.gs` di paket ini sama dengan file aktif yang Anda konfirmasi. Fungsi `doPost(e)` sudah tersedia dan membatasi aksi melalui `allowedActions`. Karena itu tidak perlu menambahkan `doPost` kedua. Jangan jalankan setup yang mengubah data hanya untuk integrasi ini. Jika kode Apps Script belum sama dengan file ini, cadangkan kode lama, cocokkan, lalu deploy New version dengan URL Web App yang sama.

## Pemeriksaan lokal
- Pemanggilan `google.script.run` pada halaman Portal diganti ke `/api/portal`.
- Fungsi API yang dipanggil dari Portal: 55.
- Fungsi yang tidak ditemukan dalam Code.gs: Tidak ada.
- Aksi yang tidak tercantum pada allowlist doPost: Tidak ada.

## Batas pengujian
Pemeriksaan ini bersifat statis/lokal. Koneksi langsung ke akun Apps Script, Google Sheets, autentikasi, dan deployment Vercel belum diuji dari sini. Lakukan uji Preview sebelum mengganti produksi.
