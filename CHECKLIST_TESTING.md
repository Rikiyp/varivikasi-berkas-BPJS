# ✅ Checklist Testing Aplikasi Validator Berkas & Coding BPJS

Dokumen ini berisi daftar uji (*checkbox testing*) untuk memastikan seluruh fitur pada aplikasi **Validator Kelengkapan Berkas & Coding BPJS (Rawat Inap & Rawat Jalan)** berjalan lancar.

---

## 1. Persiapan & Menjalankan Aplikasi
- [ ] Dependensi terpasang tanpa error (`pip install -r requirements.txt`)
- [ ] Jalankan `python verify_audit.py` dan pastikan hasil akhir `8 passed, 0 failed` (syntax `validator.py` & `app.py` OK)
- [ ] Jalankan aplikasi dengan `streamlit run app.py`
- [ ] Halaman utama terbuka di browser tanpa pesan error / traceback

---

## 2. Upload File Data Klaim (Sidebar)
- [ ] Upload file data klaim MIX (`.txt`, `.tsv`, `.xls`, `.xlsx`, atau `.csv`), misal `3511011_20261125_MIX.TXT`
- [ ] Aplikasi membaca file dengan cepat dan tidak error pada encoding/pemisah kolom
- [ ] Data Rawat Inap (`PTD = 1`) dan Rawat Jalan (`PTD = 2`) sama-sama diproses (tidak ada notifikasi data dilewati)
- [ ] Kolom `JENIS_RAWAT` terisi `Rawat Inap` untuk `PTD=1` dan `Rawat Jalan` untuk `PTD=2`
- [ ] Uji upload file tanpa kolom kunci (`SEP` / `DIAGLIST`) → muncul pesan error yang jelas

---

## 3. Tab `📊 Validasi Klaim` — KPI, Filter & Paginasi
- [ ] 7 KPI Metrics tampil lengkap (`Total SEP`, `Rawat Inap`, `Rawat Jalan`, `ACC`, `Cek Berkas`, `Perlu Revisi`, `Diluar Modul`)
- [ ] 4 Progress bar status tampil sejajar dalam 4 kolom sama lebar (teks dan persentase tidak terpotong/terlipat)
- [ ] Tombol filter **Jenis Rawat** (`Semua`, `Rawat Inap`, `Rawat Jalan`) memfilter tabel sesuai pilihan dan angkanya akurat
- [ ] Tombol filter **Status** (`Semua`, `ACC`, `Cek Berkas`, `Perlu Revisi`, `Diluar Modul`) memfilter tabel sesuai pilihan
- [ ] Filter Sidebar (`Filter Status Validasi`, `Cari SEP / MRN / Nama Pasien`, `Filter DPJP`, `Filter Kode Diagnosis`, `Filter Jenis Rawat`) bekerja dengan benar
- [ ] Pencarian yang menghasilkan tepat 1 SEP menampilkan banner ringkasan SEP tersebut
- [ ] Paginasi (`Baris per halaman`: `100`, `250`, `500`, `Semua` & navigasi `Halaman`) bekerja responsif dan tidak lemot

---

## 4. Akurasi Mesin Validasi (`validator.py` & `rules.json`)
- [ ] SEP dengan diagnosis yang tidak ada di modul (misal `Z00.0`) berstatus `Diluar Modul` (⚪)
- [ ] SEP dengan diagnosis di modul yang memenuhi seluruh aturan tanpa syarat berkas fisik berstatus `ACC` (🟢)
- [ ] SEP dengan diagnosis di modul yang memiliki syarat dokumen fisik berstatus `ACC (Cek Berkas Fisik)` (🟡)
- [ ] SEP dengan diagnosis di modul yang melanggar aturan / kurang kode prosedur berstatus `PERLU REVISI` (🔴)

---

## 5. Edit Status Inline, Detail SEP & Export Excel Utama
- [ ] Ubah status langsung di kolom `Status ✏️` pada tabel → muncul form konfirmasi perubahan di bawah tabel
- [ ] Validasi form: menyimpan tanpa mengisi `Catatan` atau `Nama Verifikator` ditolak dengan pesan error
- [ ] Mengisi `Catatan` & `Nama Verifikator` lalu klik simpan → status di tabel & KPI langsung berubah
- [ ] Pilih nomor SEP lalu klik `📄 Lihat Detail & Edit Status` → modal detail tampil lengkap (identitas, diagnosis, prosedur, hasil validasi, riwayat, dan form edit)
- [ ] Klik `📊 Siapkan File Excel (.xlsx)` → tombol `⬇️ Download Hasil Validasi (.xlsx)` muncul
- [ ] Buka `Hasil_Validasi_BPJS.xlsx` di Excel:
  - [ ] Header rapi, berwarna, *freeze panes*, dan *auto-filter* aktif
  - [ ] Warna baris sesuai status (`ACC` hijau, `Cek Berkas` kuning, `Perlu Revisi` merah, `Diluar Modul` abu-abu)
  - [ ] Kolom `Jenis Rawat` tampil dan sheet terpisah (`Semua SEP`, `Perlu Revisi`, `Cek Berkas Fisik`, `ACC`, `Diluar Modul`, serta `Riwayat Perubahan` jika ada edit)

---

## 6. Tab `🔍 Cek Data Pending`
- [ ] Upload file Excel Data Pending (contoh: `Data Pending Agustus RSDK Reguler 2026.xlsx`)
- [ ] Sheet yang memiliki `NOSEP`/`SEP` dan `JNSPENDING`/`ALASANPENDING` (`Pending RJTL`, `Pending RITL`) terpilih otomatis
- [ ] KPI ringkasan pending tampil akurat (`Data Tervalidasi`, `SEP di File Pending`, `Cocok Masuk Pending`, `Pending RI / RJ`, `Keterangan Diisi`)
- [ ] Tabel hanya menampilkan SEP yang cocok dengan Data Pending dan memiliki 4 kolom tambahan: `FILTRASI`, `JNSPENDING`, `ALASANPENDING`, `KETERANGAN`
- [ ] Filter `Filter Jenis Rawat`, `Filter Status Validasi`, dan `Cari SEP / Nama / Alasan Pending` berfungsi
- [ ] Ketik catatan pada kolom `KETERANGAN` di tabel → klik `💾 Simpan Perubahan Keterangan` → tersimpan dan jumlah `Keterangan Diisi` bertambah
- [ ] **Uji Persistensi Antar Tab:** Pindah ke tab `📊 Validasi Klaim` → klik filter `🛏️ Rawat Inap` → klik filter `🟢 ACC` → kembali ke tab `🔍 Cek Data Pending` → **file pending dan tabel tetap tampil (tidak hilang / tidak minta upload ulang)**
- [ ] Klik `🗑️ Hapus File` → file pending terhapus dan siap untuk upload file pending baru
- [ ] Klik `📊 Siapkan File Excel Data Pending (.xlsx)` → `⬇️ Download Excel Data Pending (.xlsx)` → buka `Hasil_Cek_Data_Pending_BPJS.xlsx`:
  - [ ] Hanya berisi data SEP yang masuk Data Pending
  - [ ] Memiliki 4 kolom pending (`FILTRASI`, `JNSPENDING`, `ALASANPENDING`, `KETERANGAN` beserta isian keterangannya)
  - [ ] Sheet `Data Pending`, `Pending Rawat Inap`, dan `Pending Rawat Jalan` terbentuk rapi

---

## 7. Tab `✏️ Edit Status (Bulk)`
- [ ] Filter `Filter Jenis Rawat` (`Rawat Inap`, `Rawat Jalan`, `Semua`) memisahkan daftar SEP RI dan RJ (tidak tergabung saat memilih RI/RJ)
- [ ] Mode `Pilih berdasarkan status saat ini` menampilkan SEP sesuai jenis rawat & status yang dipilih
- [ ] Mode `Pilih manual dari daftar` hanya menampilkan SEP sesuai jenis rawat yang dipilih
- [ ] Simpan edit massal memperbarui status seluruh SEP terpilih dan mencatatnya di riwayat

---

## 8. Tab `📜 Riwayat Perubahan`
- [ ] Seluruh perubahan status (dari inline edit, detail dialog, maupun bulk edit) tercatat dengan waktu, status lama, status baru, catatan, dan verifikator
- [ ] Filter SEP dan Filter Verifikator pada tabel riwayat berfungsi
- [ ] `🕐 Timeline Per SEP` menampilkan kronologi perubahan untuk SEP yang dipilih
- [ ] `⬇️ Download Riwayat (.json)` mengunduh file JSON riwayat yang valid
- [ ] `🗑️ Hapus Semua Riwayat` mengosongkan riwayat dan mereset override status

---

## 9. Tab `⚙️ Kelola Aturan Modul`
- [ ] Tabel daftar aturan menampilkan seluruh kategori & aturan dari `rules.json`
- [ ] Tambah aturan baru berhasil tersimpan ke `rules.json`
- [ ] Edit aturan existing berhasil memperbarui aturan di `rules.json`
- [ ] Hapus aturan berhasil menghapus aturan terpilih
- [ ] `⬇️ Download rules.json Terbaru` dan `📂 Import rules.json` bekerja tanpa error
