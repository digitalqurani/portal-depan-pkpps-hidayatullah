# Integrasi AI Video Studio – Portal V1.42

## Audit source upload
- `/portal` diarahkan oleh `vercel.json` ke `portal.html`.
- `index.html` adalah halaman root/landing; file ini tidak diubah.
- `index.html` dan `portal.html` pada ZIP upload identik byte-per-byte.
- Frontend memanggil backend melalui `POST /api/portal`.
- `api/portal.js` meneruskan action/args ke Google Apps Script melalui `GAS_WEB_APP_URL`.
- Login: `loginPortal(username,pin)`. Token sesi disimpan pada variabel halaman `TOKEN`.
- Dashboard memakai `getPortalDashboard(TOKEN)`, menu dibentuk dari `DATA.menus` di `render()`,
  dan modul dipilih melalui `showModule()`.
- `Code.gs` memiliki satu dispatcher `doPost(e)` dan allowlist `allowedActions`.

## Perubahan di salinan ini
- GAS: tambah `getVideoStudioIdentity(token)` dan allowlist action-nya.
- Portal: menu AI Video Studio untuk ADMIN/TU/GURU, form prompt, durasi, rasio, status job.
- Vercel: `api/video.js` sebagai proxy same-origin.
- Cloudflare: Worker memvalidasi token Portal lewat GAS, submit fal.ai queue, menyimpan job ke D1.
- D1: tabel job, project, asset, wallet, transaksi kredit, katalog model, audit.

## Model uji yang ditetapkan
- Endpoint: `fal-ai/wan/v2.7/text-to-video`
- Resolusi uji: `720p` untuk menekan biaya awal.
- Durasi antarmuka: 5 atau 10 detik.
- Rasio: 16:9, 9:16, atau 1:1.
- Safety checker: aktif.
- Payload disesuaikan dengan schema resmi Wan 2.7. Periksa harga di halaman model sebelum top-up/generasi.

## Konfigurasi
Vercel environment:
- `AI_VIDEO_WORKER_URL` = URL base Worker Cloudflare.

Cloudflare Worker secrets:
- `GAS_WEB_APP_URL` = URL Web App Apps Script aktif.
- `FAL_KEY` = API Key fal.ai.

Cloudflare Worker variable:
- `FAL_MODEL_ID` = model endpoint fal.ai yang telah dipilih dan schema input-nya diverifikasi.

## Setup D1
```bash
cd worker
npx wrangler d1 create pkpps-ai-video
```
Salin `database_id` ke `worker/wrangler.toml`, lalu:
```bash
npx wrangler d1 execute pkpps-ai-video --remote --file=migrations/0001_init.sql
```

## Deploy Worker
```bash
cd worker
npm install
npx wrangler secret put GAS_WEB_APP_URL
npx wrangler secret put FAL_KEY
npx wrangler deploy
```
Uji `/health`.

## Deploy Preview Vercel
1. Cadangkan Apps Script dan Spreadsheet.
2. Deploy `Code.gs` hasil paket sebagai New version (atau salin fungsi bridge + item allowlist).
3. Pastikan `GAS_WEB_APP_URL` menunjuk deployment GAS terbaru.
4. Tambahkan `api/video.js` dan perubahan `portal.html` ke branch Preview.
5. Isi `AI_VIDEO_WORKER_URL` pada Environment Preview.
6. Deploy Preview, login Admin, buka AI Video Studio, uji satu prompt sederhana durasi 5 detik.
7. Jangan merge Production sebelum login, role, create job, dan status job lolos uji.

## Batasan sebelum produksi
- Payload `prompt/duration/aspect_ratio` perlu dicocokkan dengan schema model fal.ai yang dipilih.
- Status masih diperiksa dengan polling manual; webhook belum dipasang.
- Output masih URL fal.ai, belum disalin ke Cloudflare R2.
- Tabel wallet/ledger baru skema; debit/reservasi kredit belum diimplementasikan.
- Rate limit dan batas pengeluaran belum diterapkan. Batasi penggunaan pada akun uji saja.
- Paket ini adalah source Preview hasil modifikasi, bukan bukti bahwa deployment produksi sudah berubah.
