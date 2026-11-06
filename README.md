# Service Report Teknisi

Aplikasi laporan maintenance teknisi CV Diarana Sejahtera.

- `index.html` dan `config.js`: tampilan, di-host di GitHub Pages.
- `apps-script/Code.gs`: server data di Google Apps Script, terhubung ke Google Sheet.

Tidak ada password atau kode akses yang disimpan di repo ini. Keduanya disimpan di Script Properties Apps Script.

## 1. Siapkan Apps Script

1. Buka project Apps Script, ganti isi `Kode.gs` dengan `apps-script/Code.gs`, lalu simpan.
2. Hapus file `Index.html` dari project Apps Script. Tampilan sekarang ada di GitHub.
3. Buka **Setelan project (ikon roda gigi) > Properti skrip**, lalu tambahkan:
   - `APP_TOKEN`: kode akses tim, misalnya `diarana-servis-2026`. Kode ini dimasukkan teknisi sekali di setiap perangkat.
   - `APP_PASSWORD`: password admin untuk edit dan hapus laporan.
4. Buka **Terapkan > Kelola penerapan > pensil**, lalu atur:
   - Versi: **Versi baru**
   - Jalankan sebagai: **Saya**
   - Yang memiliki akses: **Siapa saja**

   Simpan dengan **Terapkan**. Kalau Google meminta izin akses, setujui.
5. Buka link `/exec` di browser. Kalau muncul `{"ok":true,"app":"Maintenance History Report API",...}`, server sudah siap.

## 2. Upload ke GitHub

1. Buat repo baru di GitHub, misalnya `service-report`.
2. Upload semua file di folder ini (`index.html`, `config.js`, `README.md`, folder `apps-script`).
3. Buka **Settings > Pages**. Di bagian *Build and deployment*, pilih **Deploy from a branch**, branch **main**, folder **/ (root)**, lalu **Save**.
4. Tunggu 1 sampai 2 menit. Aplikasi bisa dibuka di `https://NAMA-AKUN.github.io/service-report/`.

## 3. Bagikan ke teknisi

Kirim link GitHub Pages dan kode akses tim ke teknisi. Saat pertama kali membuka link, teknisi diminta mengisi kode akses. Setelah itu kode tersimpan di perangkat masing-masing.

Kalau kode akses perlu diganti (misalnya ada teknisi keluar), ubah `APP_TOKEN` di Script Properties. Semua perangkat akan diminta memasukkan kode yang baru. Mengganti kode akses tidak perlu deploy ulang.

## Catatan

- Link `/exec` disimpan di `config.js`. Kalau suatu saat membuat **deployment baru** dan linknya berubah, cukup ganti link di file itu.
- Setiap perubahan `Code.gs` perlu disimpan lalu dideploy ulang dengan **Versi baru**. Perubahan `index.html` cukup di-commit ke GitHub.
- Cetak service report dan export Excel berjalan langsung dari browser.
