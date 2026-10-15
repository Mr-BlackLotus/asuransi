# Report Asuransi

Aplikasi web (Google Apps Script) untuk mencatat report asuransi ke Google Sheet dan Google Drive.

Kolom yang diinput: **Foto**, **Nominal**, **No Pesanan MP**, **No Invoice EK**.

## Isi repo

- `Code.gs` : backend. Menyimpan baris baru ke Sheet, foto ke folder Drive "Foto Report Asuransi", cek data ganda, dan menampilkan data terakhir.
- `index.html` : tampilan form (ramah HP).

## Cara pasang

1. Buka Google Sheet tujuan, lalu **Ekstensi → Apps Script**.
2. Isi `Code.gs` dengan file `Code.gs` dari repo ini.
3. Tambah file HTML bernama `index`, isi dengan `index.html` dari repo ini.
4. Di `Code.gs`, pastikan `SHEET_ID` dan `SHEET_NAME` sesuai Sheet Anda.
5. **Deploy → Deployment baru → Aplikasi web**. Jalankan sebagai: Saya. Akses: Hanya saya.
6. Setujui izin Google, lalu buka link aplikasi web yang diberikan.

Setiap kali kode diubah, buat Deployment baru agar perubahan berlaku.

## Catatan

Foto di Drive diatur "siapa saja dengan link bisa melihat" agar gambar muncul di Sheet.
