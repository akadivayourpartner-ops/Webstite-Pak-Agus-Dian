# Portal Guru Wali (GitHub Pages + Google Sheets)

Frontend statis di GitHub Pages; data tersimpan di Google Sheets lewat API Google Apps Script.

## Langkah 1 - Backend (Google Sheets + Apps Script)
1. Buka Google Sheets (unggah `Template_Portal_Guru_Wali.xlsx` atau pakai sheet kosong).
2. Ekstensi > Apps Script, tempel isi `Code.gs` (versi folder `github`).
3. Jalankan `setupPortalDatabase`, setujui izin, lalu isi sheet Config & Master Siswa.
4. Deploy > Deployment baru > Aplikasi web. Jalankan sebagai: Saya. Akses: Siapa saja. Salin URL `/exec`.

## Langkah 2 - Frontend (GitHub Pages)
1. Buat repository baru di GitHub (mis. `portal-guru-wali`), unggah `index.html`.
2. Edit `index.html`, ganti `GANTI_DENGAN_URL_WEB_APP_ANDA` pada `API_URL` dengan URL `/exec`.
3. Settings > Pages > Source: Deploy from a branch > `main` / root > Save.
4. Buka `https://USERNAME.github.io/portal-guru-wali/`.

## Catatan
- Setiap mengubah `Code.gs`, buat versi baru: Deploy > Kelola deployment > Edit > Versi baru.
- URL API terlihat di kode publik. Hanya Kotak Curhat yang diproteksi password; gunakan repo private jika perlu (Pages private butuh paket berbayar).
