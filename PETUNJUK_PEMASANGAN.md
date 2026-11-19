# Portal PKPPS Hidayatullah Karimun — V1.42 Diagnostik & Optimasi Aman

Paket ini merupakan turunan V1.41. Tujuan rilis pertama adalah mengukur sumber keterlambatan sebelum mengubah query, cache, atau alur penyimpanan.

## Perubahan V1.42
- Menambahkan pencatatan durasi setiap pemanggilan API di browser.
- Menyimpan maksimal 200 pengukuran terakhir pada memori halaman.
- Menampilkan log diagnostik di Developer Tools → Console dengan awalan `[Portal API V1.42]`.
- Menambahkan header `Server-Timing` dan `X-Portal-Proxy-Duration-Ms` pada API Vercel untuk membantu membedakan waktu proxy/backend.
- Tidak mengubah format Google Sheets, fungsi simpan absensi, hak akses, atau payload API.

## Cara melihat hasil ukur
1. Buka portal dan login.
2. Tekan F12 → Console.
3. Buka beberapa menu yang terasa lambat.
4. Jalankan `getPortalApiMetrics()` pada Console untuk melihat daftar pemanggilan dan durasinya.
5. Pilih satu request `/api/portal` di tab Network untuk melihat Timing dan response headers.

## Pemasangan aman
- Jangan langsung menimpa produksi. Buat branch uji atau deployment Preview Vercel.
- Unggah isi paket ini ke branch uji; pertahankan `GAS_WEB_APP_URL` yang sama sebagai environment variable Preview.
- Uji login, terutama Absensi KBM: buka jadwal, isi, simpan, muat ulang, dan pastikan data tetap benar di portal/Sheets.
- Jika lolos, baru merge/deploy ke Production. Simpan V1.41 sebagai rollback.

## Batasan
V1.42 adalah rilis diagnostik non-invasif, bukan klaim bahwa seluruh delay sudah diperbaiki. Pengukuran diperlukan untuk memutuskan apakah bottleneck ada pada Vercel, Apps Script, Sheets, atau rendering browser. Cache dan perubahan query belum diterapkan agar tidak menampilkan data absensi/nilai yang kedaluwarsa.
