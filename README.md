# 📡 WCL GEMOY - Panduan Deploy Online

WCL GEMOY adalah aplikasi analitik telekomunikasi 5G NR26 cepat, presisi, dan 100% kompatibel dengan Master Template Excel & PPTX.

Berikut adalah 3 metode deploy online yang dapat Anda pilih sesuai kebutuhan:

---

## ⚡ METODE 1: Instant Share (Online Langsung dalam 5 Detik)
*Cocok jika Anda ingin langsung membagikan link ke rekan kerja/atasan tanpa perlu setup server atau upload code ke GitHub.*

1. Pastikan laptop terhubung ke internet.
2. Double-click file:
   `run_wcl_public_link.bat`
3. Tunggu 5 detik, link HTTPS publik (contoh: `https://famous-tiger-42.loca.lt`) akan muncul di layar.
4. Bagikan link tersebut ke rekan kerja Anda di mana saja.
> **Tips**: Jika browser rekan kerja meminta "Tunnel Password", buka https://loca.lt/mytunnelpassword di laptop Anda dan berikan angka IP tersebut.

---

## ☁️ METODE 2: Deploy Permanen ke Streamlit Community Cloud (GRATIS)
*Cocok untuk deploy hosting cloud permanen 24/7 dengan domain sendiri seperti `https://wcl-gemoy.streamlit.app`.*

### Langkah-langkah:
1. **Buat Repository di GitHub**:
   - Buka [GitHub.com](https://github.com) dan login.
   - Buat repository baru, misalnya: `wcl-gemoy` (bisa diset **Private** atau **Public**).
2. **Upload File ke Repository**:
   Upload file-file berikut dari folder proyek ini (atau gunakan isi file `deploy_wcl_online.zip`):
   - `app.py`
   - `wcl_web_app.py`
   - `wcl_processor.py`
   - `requirements.txt`
   - `.streamlit/config.toml`
   - `Dockerfile`
   - `.gitignore`
3. **Deploy di Streamlit Cloud**:
   - Buka [share.streamlit.io](https://share.streamlit.io).
   - Klik tombol **"Create app"** (atau "New app").
   - Pilih repository GitHub Anda: `username/wcl-gemoy`.
   - Branch: `main`.
   - Main file path: `app.py` (atau `wcl_web_app.py`).
   - App URL (opsional): ketik nama domain yang diinginkan, misal: `wcl-gemoy`.
   - Klik **"Deploy!"**.
4. Dalam 1-2 menit, aplikasi Anda sudah live online di seluruh dunia!

---

## 🤗 METODE 3: Deploy ke Hugging Face Spaces (GRATIS & PRIVATE READY)
*Cocok untuk hosting gratis dengan RAM 16GB dan opsi Private Space.*

1. Buka [huggingface.co/spaces](https://huggingface.co/spaces) dan buat Space baru.
2. Pilih Space SDK: **Streamlit**.
3. Pilih visibilitas: **Public** atau **Private** (hanya tim Anda yang bisa akses).
4. Upload file proyek (`app.py`, `wcl_web_app.py`, `wcl_processor.py`, `requirements.txt`, `.streamlit/config.toml`).
5. Space akan otomatis build dan aplikasi siap digunakan.

---

## 🐳 METODE 4: Docker / VPS Server (NOC / On-Premise)
*Cocok untuk server internal perusahaan / NOC Linux / Cloud VM (AWS, GCP, DigitalOcean).*

Cukup jalankan satu perintah:
```bash
docker compose up -d --build
```
Aplikasi langsung aktif di port 8501 dan dapat di-reverse-proxy dengan Nginx/Caddy dengan SSL/HTTPS.
