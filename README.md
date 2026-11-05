# Servify — Mobile-ready MVP v3

## Menjalankan di komputer
1. Install Python 3.10 atau lebih baru.
2. Buka terminal pada folder proyek.
3. Jalankan:
   ```bash
   pip install -r requirements.txt
   streamlit run app.py
   ```

## Membuka dari HP
Tampilan sudah responsif untuk layar HP:
- Navigasi samping dimulai tertutup.
- Kolom dan kartu ditumpuk vertikal pada layar sempit.
- Tombol dibuat lebih mudah disentuh.
- Tab fitur dapat digeser ke samping.
- Tabel dapat digeser horizontal jika kolomnya banyak.

Agar bisa dibuka dari HP tanpa menjalankan Python di HP, aplikasi perlu di-hosting online (misalnya Streamlit Community Cloud). Setelah deploy, buka URL aplikasi dari Chrome di HP. ZIP ini belum dipublikasikan dan belum memiliki URL online.

## Fitur
- Beranda multi-service.
- Excel/Spreadsheet: instruksi otomatis berbasis aturan, data cleaning, join/lookup, formatting workbook, katalog kemampuan.
- Upload CSV/XLSX dan unduh hasil XLSX.

## Batasan
- Pemrosesan instruksi masih berbasis aturan, bukan AI generatif.
- Belum ada login, database, pembayaran, atau penyimpanan cloud.
- Belum mendukung kalkulasi penuh semua formula Excel.
- Jangan gunakan file berisi data pribadi/rahasia pada demo.
