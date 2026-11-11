# Praktikum 05 Web Service Engineering
## Desain Endpoint dan Implementasi Web Service dengan Framework Express

**Identitas Mahasiswa:**
- **Nama:** Muhammad Daffa Khairi Addin
- **NIM:** 240104040147
- **Program Studi:** S1 Teknologi Informasi
- **Mata Kuliah:** Web Service Engineering
- **Dosen Pengampu:** Muhayat, M.IT
- **Studi Kasus:** Layanan Katalog Buku Perpustakaan Kampus

---

### Lingkungan Sistem & Versi Perangkat
- **Sistem Operasi:** Windows 11 Home / Pro (x64)
- **Runtime:** Node.js v24.11.0 LTS
- **Package Manager:** npm v11.6.1
- **Framework & Dependencies:**
  - `express@5.3.0` (Fixed exact version)
  - `swagger-ui-express@5.0.1` (Dokumentasi OpenAPI interaktif lokal)

---

### Struktur Direktori Proyek
```text
WSE_Praktikum_05_240104040147_Muhammad Daffa Khairi Addin/
├── app.cjs                     # Konfigurasi aplikasi Express, middleware, routing, Swagger UI, error handler
├── server.cjs                  # Entry point listener HTTP (127.0.0.1:3200)
├── client.cjs                  # HTTP client CLI untuk pengujian operasi API
├── package.json                # Metadata proyek dan dependensi
├── package-lock.json           # Lockfile dependensi npm
├── README.md                   # Petunjuk instalasi dan eksekusi
├── docs/
│   └── openapi.json            # Dokumen spesifikasi OpenAPI 3.1.0
├── scripts/
│   └── generate-openapi.cjs    # Generator dokumen OpenAPI otomatis
├── samples/                    # Berkas uji muatan JSON & format salah
│   ├── buku-baru.json          # Contoh body POST valid
│   ├── buku-ubah.json          # Contoh body PUT valid
│   ├── buku-salah-tipe.json    # Contoh uji tipe boolean salah (string "true")
│   ├── buku-parsial.json       # Contoh uji kelengkapan field (hanya judul)
│   └── json-rusak.txt          # Contoh uji sintaks JSON rusak/tidak lengkap
└── src/                        # Lapisan arsitektur layanan
    ├── errors.cjs              # Definisi HttpError dan fungsi validasi
    ├── buku.repository.cjs     # Penyimpanan katalog in-memory array
    ├── buku.service.cjs        # Logika bisnis dan aturan penggunaan katalog
    ├── buku.controller.cjs     # Penanganan konteks HTTP (request/response)
    └── buku.routes.cjs         # Pemetaan rute HTTP dan middleware media type
```

---

### Petunjuk Instalasi dan Menjalankan Proyek

#### 1. Instalasi Dependensi (Reproducible Build)
Dari direktori akar proyek, jalankan:
```bash
npm ci
```

#### 2. Menghasilkan Dokumen OpenAPI
Jalankan skrip generator dokumentasi:
```bash
npm run docs
```
Hasil berkas akan disimpan di `docs/openapi.json`.

#### 3. Menjalankan Server API
Jalankan listener server pada terminal pertama:
```bash
npm start
```
Server akan aktif mendengarkan koneksi lokal pada:
- **Alamat Layanan:** `http://127.0.0.1:3200/api/buku`
- **Dokumentasi OpenAPI Mentah:** `http://127.0.0.1:3200/openapi.json`
- **Swagger UI Interaktif:** `http://127.0.0.1:3200/docs/`

#### 4. Menjalankan Pengujian Client
Buka terminal kedua pada direktori akar proyek untuk mengirim permintaan HTTP menggunakan `client.cjs`:

- **Membaca seluruh katalog:**
  ```bash
  node client.cjs GET /api/buku
  ```
- **Pencarian judul:**
  ```bash
  node client.cjs GET "/api/buku?q=kotlin"
  ```
- **Membaca detail buku (ID 1):**
  ```bash
  node client.cjs GET /api/buku/1
  ```
- **Menambah buku baru (POST):**
  ```bash
  node client.cjs POST /api/buku samples/buku-baru.json
  ```
- **Mengganti data buku (PUT):**
  ```bash
  node client.cjs PUT /api/buku/3 samples/buku-ubah.json
  ```
- **Menghapus buku (DELETE):**
  ```bash
  node client.cjs DELETE /api/buku/3
  ```
- **Uji coba validasi dan error handling:**
  ```bash
  node client.cjs GET /api/buku/1abc
  node client.cjs POST /api/buku samples/buku-salah-tipe.json
  node client.cjs POST /api/buku samples/json-rusak.txt
  node client.cjs POST /api/buku samples/buku-baru.json text/plain
  ```

---

### Batasan Praktikum (Scope Boundaries)
1. **Lokal Loopback:** Berjalan di `127.0.0.1` tanpa TLS/HTTPS (hanya untuk latihan pengembangan lokal).
2. **In-Memory Array:** Penyimpanan data bersifat volatil di memori RAM proses Node.js. Ketika server dihentikan atau di-restart, seluruh perubahan data (tambah/ubah/hapus) akan kembali ke 2 data awal (ID 1 dan ID 2).
3. **Stateless CRUD:** Tanpa sistem autentikasi/login, tanpa database persisten (SQL/NoSQL), dan tanpa transaksi peminjaman perpustakaan.
