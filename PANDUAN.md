# Panduan Pemasangan — Aplikasi Laporan Part Bekas

Aplikasi web mobile (dibuka dari browser HP) untuk **teknisi melaporkan part bekas kepada PIC**, lengkap dengan
foto SPK (data terisi otomatis lewat OCR), foto setiap part, indikator kelengkapan, dan serah terima ber-PIN + tanda tangan.
Setelah diterima, PIC membuat **tanggal planning** antar ke asuransi, lalu mencatat **tanggal aktual** saat part benar-benar diantar
(otomatis dinilai *Tepat Waktu / Lebih Cepat / Terlambat*).
Semua data otomatis tersimpan ke **Google Spreadsheet**, foto tersimpan ke **Google Drive**.

---

## 1. Isi paket

| File | Fungsi |
|---|---|
| `Code.gs` | Backend (simpan ke Sheet, OCR SPK, upload foto, serah terima, indikator) |
| `Index.html` | Tampilan aplikasi untuk smartphone (Lapor, Serah Terima, Riwayat) |
| `appsscript.json` | Pengaturan proyek (zona waktu Jakarta, layanan Drive API untuk OCR, akses web app) |

## 2. Langkah pemasangan (±10 menit, cukup sekali)

1. Buka **Google Sheets** → buat spreadsheet baru, beri nama mis. `Laporan Part Bekas`.
2. Klik menu **Extensions (Ekstensi) → Apps Script**.
3. Di editor Apps Script:
   - Hapus isi `Code.gs` bawaan, lalu **tempel isi `Code.gs`** dari paket ini.
   - Klik **＋ → HTML**, beri nama **`Index`** (tanpa .html), lalu tempel isi `Index.html`.
   - Klik ⚙️ **Project Settings** → centang **"Show 'appsscript.json' manifest file in editor"**,
     kembali ke Editor, buka `appsscript.json`, ganti isinya dengan file dari paket ini.
     *(Ini otomatis mengaktifkan layanan **Drive API** yang dipakai untuk membaca foto SPK.)*
4. Klik **Save** 💾.
5. Pilih fungsi **`setup`** di dropdown atas → klik **Run** ▶.
   Saat diminta izin: **Review permissions → pilih akun → Advanced → Go to project (unsafe) → Allow**.
   Setup membuat sheet: `Laporan`, `Detail Part`, `Master Teknisi`, `Master PIC`, `Master Type Mobil` dan folder Drive
   `Laporan Part Bekas - Foto`.
6. Klik **Deploy → New deployment** → ikon ⚙️ pilih **Web app**:
   - *Execute as*: **Me** (pemilik) — teknisi tidak perlu akses ke spreadsheet.
   - *Who has access*: **Anyone** (siapa saja yang punya link) atau **Anyone within [domain kantor]** jika memakai Google Workspace.
   - Klik **Deploy** → salin **Web app URL**.
7. Bagikan URL ke teknisi & PIC. Di HP: buka URL di Chrome → menu ⋮ → **Add to Home screen** agar tampil seperti aplikasi.

> Setiap kali kode diubah: **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy** (URL tetap sama).

## 3. Isi data master (di spreadsheet)

| Sheet | Isi |
|---|---|
| **Master Teknisi** | Nama teknisi, kolom Aktif `Y`/`N` |
| **Master PIC** | Nama PIC, Email (untuk notifikasi laporan baru, boleh kosong), **PIN** login PIC, Aktif `Y`/`N` |
| **Master Type Mobil** | Daftar type mobil — dipakai untuk mendeteksi type dari foto SPK & saran isian |
| **Master Asuransi** | Daftar nama asuransi rekanan — muncul sebagai pilihan saat PIC mengisi pengantaran |

⚠️ Ganti PIN contoh (`1234`, `5678`) sebelum dipakai, dan **jangan bagikan spreadsheet** ke teknisi (PIN ada di sheet ini).

## 4. Alur penggunaan

**Teknisi — menu 📝 Lapor**
1. Pilih nama teknisi & PIC tujuan.
2. Ketuk **Foto SPK** → kamera HP terbuka → foto SPK.
   No. SPK, No. Polisi, dan Type Mobil **terisi otomatis** (tanda ⚡ otomatis). Kolom yang tidak terbaca ditandai *isi manual*.
   Jika No. SPK sudah pernah dilaporkan, muncul peringatan agar tidak dobel.
3. Tambah part bekas: nama, no. part (opsional), qty, kondisi, dan **foto wajib** untuk tiap part.
4. Tanda tangan teknisi (diserahkan oleh).
5. **Indikator Kelengkapan** harus 100% (semua centang hijau) baru tombol **Kirim** aktif.

**PIC — menu 🤝 Serah Terima**
1. Login dengan nama PIC + PIN → muncul daftar laporan yang menunggu.
2. Buka laporan → cocokkan foto dengan part fisik → tandai **✔ Sesuai** / **✖ Tidak Sesuai** (wajib catatan jika tidak sesuai).
3. Tanda tangan PIC → **Terima Semua Part** / **Terima Sebagian**, atau **Tolak Laporan** (wajib alasan).

**PIC — Antar ke Asuransi (menu 🤝 Serah Terima), 2 langkah**

*Langkah 1 — Planning* (bagian **🗓️ Perlu Planning Antar ke Asuransi**)
1. Part yang sudah diterima otomatis masuk ke daftar ini (milik PIC penerima).
2. Buka → isi **Nama Asuransi** (wajib), No. Klaim/Polis, **Tanggal Planning Diantar ke Asuransi** (wajib, default besok,
   tidak boleh sebelum tanggal diterima PIC), catatan → **Simpan Planning**. Laporan pindah ke *Jadwal Antar*.

*Langkah 2 — Aktual* (bagian **🚚 Jadwal Antar ke Asuransi**, urut tanggal planning terdekat)
1. Tiap jadwal diberi tanda: **H-n**, **Besok**, **📍 Hari ini**, atau **⛔ Lewat jadwal n hari** (merah).
2. Buka → isi **Tanggal Aktual Diantar ke Asuransi** (wajib, default hari ini, tidak boleh melebihi hari ini),
   penerima di asuransi, foto bukti tanda terima (opsional), catatan → **Simpan Tanggal Aktual**.
3. Aplikasi langsung menampilkan perbandingan Planning vs Aktual: **✅ Tepat Waktu**, **✅ Lebih Cepat n hari**, atau **⚠️ Terlambat n hari**.
4. Perlu mundur jadwal? Tekan **✏️ Ubah Planning / Reschedule** — tanggal lama → baru beserta alasannya tercatat otomatis di *Catatan Pengantaran*.

Hanya part yang ditandai *Sesuai* saat serah terima yang tercantum sebagai part yang diantar.

**Semua — menu 📊 Riwayat**: ringkasan jumlah, pencarian SPK/No. Polisi/teknisi, dan detail lengkap beserta foto & tanda tangan.

## 5. Arti indikator

| Indikator | Arti |
|---|---|
| ✅ Lengkap (n/n) | Foto SPK ada & semua part sudah difoto |
| ❌ Belum Lengkap | Ada part tanpa foto |
| 🟡 Menunggu Serah Terima PIC | Laporan terkirim, PIC belum mengecek |
| 🟢 Selesai - Semua Part Sesuai & Diterima | Semua part dicek sesuai & diterima PIC |
| 🟠 Diterima Sebagian (x/n) | Sebagian part tidak sesuai |
| 🔴 Ditolak PIC | Laporan ditolak, teknisi perlu membuat laporan ulang |
| 🟣 Terjadwal ke Asuransi - Planning tgl | Sudah ada planning, part belum diantar (kolom tanggal planning **merah** jika sudah lewat) |
| 🔵 Sudah Diantar ke Asuransi tgl (x/n part) - Tepat Waktu / Terlambat n hari | Part sudah diantar pada tanggal aktual tersebut |

Kolom **Ketepatan Waktu**: ✅ hijau = tepat waktu / lebih cepat, ⚠️ merah = terlambat.
Menu Riwayat menampilkan jumlah **Belum Planning**, **Terjadwal** (+ berapa yang lewat jadwal), dan **Sudah Diantar** (+ berapa tepat vs telat).

Di spreadsheet, kolom **Indikator Foto**, **Status Serah Terima**, **Indikator** (sheet `Laporan`) dan **Cek PIC** (sheet `Detail Part`)
otomatis berwarna hijau/kuning/oranye/merah.

## 6. Kolom yang tersimpan

**Sheet `Laporan`** (1 baris per laporan): ID Laporan, Tanggal Lapor, Nama Teknisi, PIC Tujuan, No. SPK, No. Polisi, Type Mobil,
Jumlah Part, Part Berfoto, Indikator Foto, Part Diterima PIC, Status Serah Terima, Indikator, **Diserahkan Oleh, Diterima Oleh,
Tanggal Diterima** (kolom serah terima), Catatan Teknisi, Catatan PIC, link Foto SPK, link Folder Foto, link TTD Teknisi, link TTD PIC, Teks OCR SPK,
lalu kolom pengantaran asuransi: **Status Asuransi** (Belum Dijadwalkan / Terjadwal / Sudah Diantar), **Nama Asuransi, No. Klaim,
Tanggal Planning Diantar ke Asuransi, Tanggal Aktual Diantar ke Asuransi, Selisih Hari (Aktual - Planning), Ketepatan Waktu,
Diantar Oleh (PIC), Penerima di Asuransi**, link Bukti Tanda Terima Asuransi, Catatan Pengantaran (termasuk riwayat reschedule),
Waktu Input Planning, Waktu Input Pengantaran.

**Sheet `Detail Part`** (1 baris per part): ID Laporan, No, Nama Part, No. Part, Qty, Kondisi, link Foto Part, Cek PIC, Catatan PIC, Waktu Cek.

Foto tersimpan rapi di Drive: `Laporan Part Bekas - Foto / 2026-11 / LP-261121-001_SPK.../`.

## 7. Memperbarui aplikasi yang sudah terpasang

Jika versi sebelumnya sudah dipakai: tempel ulang `Code.gs` dan `Index.html`, jalankan **`setup`** sekali lagi
(kolom baru otomatis disisipkan di posisinya & kolom lama "Tanggal Diantar ke Asuransi" diganti nama menjadi
"Tanggal Aktual Diantar ke Asuransi" — data lama tidak bergeser; sheet `Master Asuransi` dibuat jika belum ada),
lalu **Deploy → Manage deployments → Edit → New version**. Laporan lama yang sudah diterima PIC sebelum update
otomatis muncul di daftar **Perlu Planning Antar ke Asuransi**.

## 8. Tips akurasi baca SPK

- Foto dari atas, kertas rata, cahaya cukup, tidak buram, seluruh SPK masuk frame.
- Deteksi mengenali label umum: `No. SPK / WO / PKB / RO`, `No. Polisi / Nopol / Plat`, `Type / Tipe / Model`.
- Jika format nomor SPK bengkel Anda khas (mis. `SPK-26110001`), isi `SPK_CUSTOM_REGEX` di bagian **KONFIGURASI** `Code.gs`,
  contoh: `'SPK-?\\d{8}'`. Uji dengan fungsi `testParse` (ganti teks contoh dengan teks SPK Anda).
- Hasil otomatis selalu bisa dikoreksi manual oleh teknisi sebelum dikirim.

## 9. Troubleshooting

| Masalah | Solusi |
|---|---|
| "OCR gagal … Drive API" | Pastikan `appsscript.json` sudah diganti, atau di editor: **Services ＋ → Drive API → Add** |
| Perubahan kode tidak muncul | Deploy ulang dengan **New version** |
| "Spreadsheet belum di-setup" | Jalankan fungsi `setup` sekali lagi |
| Email ke PIC tidak terkirim | Isi kolom Email di `Master PIC`; kuota MailApp akun gratis ±100 email/hari |
| Kamera tidak terbuka | Buka link di Chrome/Safari (bukan browser dalam aplikasi chat) |
