# 📖 Kamus Arab-Indonesia Parser

Web app Streamlit untuk mengekstrak entri kamus Indonesia-Arab dari PDF ke Excel secara otomatis menggunakan Google Gemini AI.

---

## 🚀 Cara Install & Jalankan

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Jalankan app
```bash
streamlit run kamus_app.py
```

App akan terbuka di browser: `http://localhost:8501`

---

## ⚙️ Cara Pakai

1. **Masukkan Gemini API Key**  
   Buat di: https://aistudio.google.com/app/apikey  
   *(gratis, cukup daftar Google account)*

2. **Pilih Model Gemini**  
   - `gemini-2.0-flash` → paling cepat (recommended)  
   - `gemini-1.5-flash` → cepat & stabil  
   - `gemini-1.5-pro` → paling akurat tapi lebih lambat

3. **Upload PDF** kamus Indonesia-Arab

4. **Setting batch:**  
   - **Halaman per batch**: 5 = standar, 10 = lebih cepat  
   - **DPI**: 150 = standar yang bagus  
   - **Jeda antar batch**: 3 detik (mencegah rate limit)

5. **Klik "Mulai Proses"**  
   Pantau entri yang muncul satu per satu di Live Feed!

6. **Download Excel** setelah selesai

---

## 📊 Kolom Output Excel

| Kolom | Isi | Contoh |
|-------|-----|--------|
| LEMA | Aksara Arab utama | قَرْنٌ |
| SUBLEMA | Aksara Arab turunan | لِعِدَّةِ قُرُونٍ |
| TRANSLITERASI | Romanisasi Latin | qarnun |
| EKUIVALEN | Arti Indonesia ≤2 kata | abad |
| ALTERNATIF | Arti alternatif lain | - |
| ARTI | Keterangan/penjelasan | pada abad ini |
| SINONIM | Sinonim | - |

---

## 💡 Tips

- Untuk **221 halaman** dengan batch 5: ~45 panggilan API, estimasi 5-8 menit
- Gunakan **range halaman** di sidebar untuk proses bertahap (misal 1-50, lalu 51-100, dst.)
- **Reset** tidak menghapus file Excel yang sudah didownload
- Font **Noto Sans Arabic** di Excel: download di https://fonts.google.com/noto/specimen/Noto+Sans+Arabic

---

## 🔧 Troubleshooting

| Masalah | Solusi |
|---------|--------|
| Error 429 (rate limit) | Naikkan jeda antar batch ke 5-10 detik |
| JSON parse error | Kurangi halaman per batch (coba 3) |
| Entri kosong/sedikit | Naikkan DPI ke 200 |
| Banyak entri salah kolom | Gunakan model 1.5-pro yang lebih akurat |
