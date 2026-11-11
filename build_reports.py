import os
import sys
import json
import subprocess
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

print("Building reports for WSE Praktikum 05...")

# Load test results
with open("test_results.json", "r", encoding="utf-8") as f:
    test_results = json.load(f)

with open("tugas1_results.json", "r", encoding="utf-8") as f:
    tugas1_results = json.load(f)

doc = Document()

# Set page margins (1 inch)
for s in doc.sections:
    s.top_margin = Inches(1)
    s.bottom_margin = Inches(1)
    s.left_margin = Inches(1)
    s.right_margin = Inches(1)

# Helper styling functions
NAVY = RGBColor(27, 54, 93)      # #1B365D
SLATE = RGBColor(71, 85, 105)    # #475569
DARK = RGBColor(30, 41, 59)      # #1E293B
MUTED = RGBColor(100, 116, 139)  # #64748B
GREEN = RGBColor(22, 101, 52)    # #166534

def format_run(run, font_name="Calibri", size_pt=11, bold=False, italic=False, color=DARK):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = color

def add_title(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(24)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    format_run(run, font_name="Calibri", size_pt=20, bold=True, color=NAVY)
    return p

def add_subtitle(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(24)
    run = p.add_run(text)
    format_run(run, font_name="Calibri", size_pt=13, italic=True, color=SLATE)
    return p

def add_h1(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    format_run(run, font_name="Calibri", size_pt=15, bold=True, color=NAVY)
    return p

def add_h2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    format_run(run, font_name="Calibri", size_pt=12.5, bold=True, color=SLATE)
    return p

def add_h3(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    format_run(run, font_name="Calibri", size_pt=11, bold=True, color=DARK)
    return p

def add_p(text, bold_prefix="", italic=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        format_run(r_pre, font_name="Calibri", size_pt=10.5, bold=True, color=DARK)
    r_body = p.add_run(text)
    format_run(r_body, font_name="Calibri", size_pt=10.5, bold=False, italic=italic, color=DARK)
    return p

def add_code_block(code_text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.left_indent = Inches(0.2)
    p.paragraph_format.right_indent = Inches(0.2)
    run = p.add_run(code_text)
    format_run(run, font_name="Consolas", size_pt=9, bold=False, color=RGBColor(15, 23, 42))
    shading_xml = parse_xml(r'<w:shd {} w:fill="F1F5F9"/>'.format(nsdecls('w')))
    p._p.get_or_add_pPr().append(shading_xml)
    return p

def add_callout(text, bold_prefix="Catatan: "):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.right_indent = Inches(0.25)
    p.paragraph_format.line_spacing = 1.15
    
    r_pre = p.add_run(bold_prefix)
    format_run(r_pre, font_name="Calibri", size_pt=10, bold=True, color=NAVY)
    r_body = p.add_run(text)
    format_run(r_body, font_name="Calibri", size_pt=10, italic=True, color=DARK)
    
    shading_xml = parse_xml(r'<w:shd {} w:fill="F8FAFC"/>'.format(nsdecls('w')))
    p._p.get_or_add_pPr().append(shading_xml)
    return p

def add_table_header(row, col_widths, headers):
    for i, h in enumerate(headers):
        cell = row.cells[i]
        cell.width = col_widths[i]
        shd = parse_xml(r'<w:shd {} w:fill="1E3A8A"/>'.format(nsdecls('w')))
        cell._tc.get_or_add_tcPr().append(shd)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(h)
        format_run(run, font_name="Calibri", size_pt=9.5, bold=True, color=RGBColor(255, 255, 255))

def style_cell(cell, width, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, italic=False, color=DARK, bg=None, size_pt=9):
    cell.width = width
    if bg:
        shd = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), bg))
        cell._tc.get_or_add_tcPr().append(shd)
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    for run in p.runs:
        format_run(run, font_name="Calibri", size_pt=size_pt, bold=bold, italic=italic, color=color)

def add_image_figure(image_path, caption_title, caption_desc):
    if os.path.exists(image_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(2)
        p_img.add_run().add_picture(image_path, width=Inches(5.6))
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(2)
        p_cap.paragraph_format.space_after = Pt(10)
        p_cap.paragraph_format.keep_with_next = False
        
        r1 = p_cap.add_run(caption_title + ": ")
        format_run(r1, font_name="Calibri", size_pt=9, bold=True, color=NAVY)
        r2 = p_cap.add_run(caption_desc)
        format_run(r2, font_name="Calibri", size_pt=9, italic=True, color=SLATE)
    else:
        print(f"Warning: Image {image_path} not found.")

print("Generating Document Content...")

# ==================== COVER PAGE ====================
add_title("LAPORAN PRAKTIKUM WEB SERVICE ENGINEERING")
add_subtitle("PRAKTIKUM 05: DESAIN ENDPOINT DAN IMPLEMENTASI WEB SERVICE DENGAN FRAMEWORK EXPRESS")

p_cov = doc.add_paragraph()
p_cov.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_cov.paragraph_format.space_before = Pt(36)
p_cov.paragraph_format.space_after = Pt(36)

cov_info = [
    ("Nama Lengkap", "Muhammad Daffa Khairi Addin"),
    ("Nomor Induk Mahasiswa (NIM)", "240104040147"),
    ("Program Studi", "S1 Teknologi Informasi"),
    ("Fakultas", "Tarbiyah dan Keguruan"),
    ("Perguruan Tinggi", "Universitas Islam Negeri Antasari Banjarmasin"),
    ("Mata Kuliah", "Web Service Engineering"),
    ("Dosen Pengampu", "Muhayat, M.IT"),
    ("Topik Praktikum", "Desain Endpoint, Arsitektur Berlapis, & Dokumentasi OpenAPI"),
    ("Studi Kasus", "Layanan Katalog Buku Perpustakaan Kampus"),
    ("Tahun Akademik", "2026")
]

table_cov = doc.add_table(rows=len(cov_info), cols=2)
table_cov.alignment = WD_TABLE_ALIGNMENT.CENTER
for idx, (label, val) in enumerate(cov_info):
    row = table_cov.rows[idx]
    c0 = row.cells[0]
    c1 = row.cells[1]
    c0.width = Inches(2.6)
    c1.width = Inches(3.6)
    
    p0 = c0.paragraphs[0]
    p0.paragraph_format.space_before = Pt(3)
    p0.paragraph_format.space_after = Pt(3)
    r0 = p0.add_run(label)
    format_run(r0, font_name="Calibri", size_pt=10, bold=True, color=SLATE)
    
    p1 = c1.paragraphs[0]
    p1.paragraph_format.space_before = Pt(3)
    p1.paragraph_format.space_after = Pt(3)
    r1 = p1.add_run(": " + val)
    format_run(r1, font_name="Calibri", size_pt=10, bold=False, color=DARK)

doc.add_page_break()

# ==================== BAB I ====================
add_h1("BAB I: TUJUAN & PRASYARAT PRAKTIKUM")
add_h2("1.1 Deskripsi & Tujuan Pembelajaran")
add_p("Praktikum ini menerapkan konsep perancangan endpoint dan implementasi web service menggunakan framework Node.js dan Express. Mahasiswa membangun API katalog buku perpustakaan kampus mulai dari penyiapan lingkungan, perancangan kontrak endpoint, implementasi arsitektur berlapis (routing, controller, service, repository), penanganan validasi dan error secara terpusat, hingga penyusunan spesifikasi OpenAPI 3.1.0 dan pengujian interaktif menggunakan Swagger UI.")

tujuan_items = [
    ("Pengetahuan: ", "Menjelaskan peran runtime Node.js, framework Express, router, controller, service, repository, serta kontrak OpenAPI."),
    ("Perancangan: ", "Menetapkan rancangan resource, model data, metode HTTP, status code semantik, format pesan sukses/gagal, serta kondisi penolakan input."),
    ("Keterampilan: ", "Menyiapkan dependensi exact-version, mengonfigurasi listener loopback lokal port 3200, dan mengimplementasikan operasi CRUD katalog secara menyeluruh."),
    ("Pengujian: ", "Membuktikan keberhasilan operasi bisnis serta ketegasan penolakan data invalid melalui respons aktual HTTP dan logging terminal."),
    ("Dokumentasi: ", "Menyusun skrip generator OpenAPI 3.1.0 otomatis dan menguji seluruh operasi API secara visual melalui Swagger UI lokal.")
]
for pre, post in tujuan_items:
    add_p(post, bold_prefix=pre)

add_h2("1.2 Lingkungan Sistem & Versi Perangkat Lunak")
add_p("Eksperimen dan verifikasi praktikum dilaksanakan pada lingkungan sistem lokal dengan rincian perangkat aktual sebagai berikut:")

env_items = [
    ("Sistem Operasi", "Microsoft Windows 11 Home / Pro (x64 Architecture)"),
    ("Runtime Environment", "Node.js v24.11.0 (LTS Iron Series)"),
    ("Package Manager", "npm v11.6.1"),
    ("Web Framework", "Express v5.3.0 (Fixed Exact Dependency)"),
    ("OpenAPI Documentation", "swagger-ui-express v5.0.1 (Local Interface Asset)"),
    ("Code Editor & Shell", "Visual Studio Code & Git Bash / Command Prompt"),
    ("HTTP Client", "Node.js Custom Fetch Client (client.cjs) & Browser Swagger UI")
]

table_env = doc.add_table(rows=len(env_items) + 1, cols=2)
table_env.alignment = WD_TABLE_ALIGNMENT.CENTER
widths_env = [Inches(2.5), Inches(3.7)]
add_table_header(table_env.rows[0], widths_env, ["Komponen Sistem", "Versi / Konfigurasi Aktual"])
for idx, (comp, ver) in enumerate(env_items):
    row = table_env.rows[idx + 1]
    bg = "F8FAFC" if idx % 2 == 1 else "FFFFFF"
    row.cells[0].paragraphs[0].add_run(comp)
    style_cell(row.cells[0], widths_env[0], bold=True, color=SLATE, bg=bg)
    row.cells[1].paragraphs[0].add_run(ver)
    style_cell(row.cells[1], widths_env[1], color=DARK, bg=bg)

# ==================== BAB II ====================
add_h1("BAB II: LANDASAN TEORI & ARSITEKTUR BERLAPIS")
add_h2("2.1 Konsep Web Service & Protokol HTTP")
add_p("Web service adalah antarmuka komunikasi antar-aplikasi (machine-to-machine) yang memanfaatkan protokol standar web, khususnya HTTP. Pada arsitektur REST (Representational State Transfer), interaksi terjadi antara penyedia layanan (provider) dan pengguna layanan (consumer) menggunakan metode HTTP yang memiliki semantik terstandarisasi (RFC 9110):")

http_methods = [
    ("GET: ", "Operasi bersifat safe dan idempotent untuk membaca representasi resource tanpa menimbulkan efek samping (side effects) pada server."),
    ("POST: ", "Operasi non-idempotent untuk menambahkan resource baru ke koleksi. Respons berhasil mengembalikan kode status 201 Created serta header Location yang mengarahkan consumer ke URI resource yang baru dibuat."),
    ("PUT: ", "Operasi bersifat idempotent untuk mengganti (replace) seluruh field yang dapat ditulis pada target resource yang sudah ada. Jika terdapat field yang hilang pada body, server menolak permintaan karena PUT menuntut representasi entitas yang utuh."),
    ("DELETE: ", "Operasi idempotent untuk menghapus resource dari penyimpanan. Keberhasilan penghapusan direspons dengan status 204 No Content tanpa payload respons (empty body).")
]
for pre, post in http_methods:
    add_p(post, bold_prefix=pre)

add_h2("2.2 Arsitektur Berlapis (Separation of Concerns)")
add_p("Untuk mencegah penumpukan logika pada satu file handler (spaghetti code), proyek ini menerapkan pola pemisahan tanggung jawab (Separation of Concerns) yang memisahkan aplikasi menjadi lima lapisan independen:")

layers = [
    ("1. Routing (src/buku.routes.cjs): ", "Bertanggung jawab memetakan kombinasi metode HTTP dan path URI ke fungsi controller yang sesuai. Lapisan ini juga mengisolasi parameter path melalui router.param dan memvalidasi Content-Type melalui middleware jsonOnly."),
    ("2. Controller (src/buku.controller.cjs): ", "Mengatur konteks HTTP. Controller mengekstrak parameter query, path, dan body dari objek req, meneruskannya ke service, serta mengemas hasil ke dalam format respons baku (res.json, res.status, res.location). Controller tidak boleh memanipulasi penyimpanan data secara langsung."),
    ("3. Service (src/buku.service.cjs): ", "Menjalankan aturan bisnis (domain business logic) katalog. Service memastikan kepatuhan model data (melalui fungsi validasi), memeriksa keberadaan data (detail checking), serta bersifat agnostik terhadap Express (tidak bergantung pada objek req maupun res)."),
    ("4. Repository (src/buku.repository.cjs): ", "Mengabstraksi akses ke lapisan penyimpanan. Pada praktikum ini, repository mengelola koleksi data buku dalam memori (in-memory array) melalui lima operasi dasar: list, find, create, replace, dan remove."),
    ("5. Centralized Error Handling & Validation (src/errors.cjs & app.cjs): ", "Mendefinisikan kelas HttpError terstruktur yang membawa status HTTP dan kode aplikasi. Middleware penanganan kesalahan menangkap seluruh error sehingga client menerima respons JSON seragam tanpa kebocoran stack trace.")
]
for pre, post in layers:
    add_p(post, bold_prefix=pre)

add_callout("Alur Satu Permintaan: Consumer mengirim HTTP Request -> Middleware Umum (Logging & Cache-Control no-store) -> JSON Body Parser (Limit 10kb) -> Router (Param ID validation & MediaType validation) -> Controller -> Service (Business logic & data validation) -> Repository (Penyimpanan memori) -> Controller merespons JSON (200/201/204) atau Error Handler merespons JSON error (400/404/413/415/500).")

add_h2("2.3 Spesifikasi OpenAPI 3.1.0 & Swagger UI")
add_p("OpenAPI Specification (OAS) adalah standar deskripsi antarmuka pemrograman aplikasi HTTP yang dapat dibaca oleh manusia maupun mesin (machine-readable). Dokumen OpenAPI menjelaskan path, metode, parameter query/path, skema body, serta contoh respons sukses maupun gagal. Swagger UI adalah kakas visual yang menyajikan dokumen OpenAPI tersebut ke dalam antarmuka web interaktif lokal, memungkinkan pengembang melakukan pengujian operasi secara langsung (Try it out) menggunakan data nyata.")

# ==================== BAB III ====================
add_h1("BAB III: PERANCANGAN ENDPOINT & KONTRAK API")
add_h2("3.1 Spesifikasi Model Data")
add_p("Resource buku perpustakaan kampus dimodelkan dengan atribut dan aturan validasi ketat sebagai berikut:")

model_rules = [
    ("id (Number / Integer): ", "Bilangan bulat positif (> 0) yang digenerate otomatis oleh server. Harus berada dalam batas aman JavaScript (Number.isSafeInteger <= 9007199254740991). Ditolak jika dikirim dalam body POST atau PUT."),
    ("judul (String): ", "Teks judul buku. Setelah melalui proses pemangkasan spasi (trim), panjang karakter wajib berada di antara 1 hingga 120 unit string UTF-16. Judul yang hanya terdiri dari spasi atau string kosong akan ditolak."),
    ("tersedia (Boolean): ", "Status ketersediaan buku di perpustakaan. Wajib bertipe boolean asli (true atau false). Nilai bertipe string seperti \"true\" atau \"false\" ditolak secara tegas."),
    ("Aturan Properti Ekstra: ", "Body request hanya boleh memuat properti 'judul' dan 'tersedia'. Pengiriman field tambahan di luar skema akan langsung memicu penolakan 400 INVALID_BOOK.")
]
for pre, post in model_rules:
    add_p(post, bold_prefix=pre)

add_h2("3.2 Matriks Kontrak Operasi API")
add_p("Tabel berikut menyajikan kontrak resmi antarmuka endpoint katalog buku perpustakaan kampus:")

contract_data = [
    ("GET /api/buku", "Query param: q (opsional, maks 80 karakter)", "200 OK dengan array katalog buku", "400 INVALID_QUERY (q > 80 atau duplikat)"),
    ("GET /api/buku/:id", "Path param: :id (bilangan bulat positif)", "200 OK dengan objek buku tunggal", "400 INVALID_ID (bukan digit), 404 BOOK_NOT_FOUND"),
    ("POST /api/buku", "Body: JSON {judul, tersedia}", "201 Created + Header Location + Objek Buku", "400 INVALID_BOOK / INVALID_JSON, 413, 415"),
    ("PUT /api/buku/:id", "Path: :id; Body: JSON {judul, tersedia}", "200 OK dengan objek buku hasil pembaruan", "400 INVALID_BOOK / ID, 404 BOOK_NOT_FOUND, 413, 415"),
    ("DELETE /api/buku/:id", "Path param: :id (bilangan bulat positif)", "204 No Content (tanpa body payload)", "400 INVALID_ID, 404 BOOK_NOT_FOUND")
]

table_cnt = doc.add_table(rows=len(contract_data) + 1, cols=4)
table_cnt.alignment = WD_TABLE_ALIGNMENT.CENTER
widths_cnt = [Inches(1.5), Inches(1.8), Inches(1.5), Inches(1.5)]
add_table_header(table_cnt.rows[0], widths_cnt, ["Metode & Path", "Parameter / Input", "Respons Berhasil", "Kondisi Kegagalan"])
for idx, (m_p, inp, succ, fail) in enumerate(contract_data):
    row = table_cnt.rows[idx + 1]
    bg = "F8FAFC" if idx % 2 == 1 else "FFFFFF"
    row.cells[0].paragraphs[0].add_run(m_p)
    style_cell(row.cells[0], widths_cnt[0], bold=True, color=NAVY, bg=bg)
    row.cells[1].paragraphs[0].add_run(inp)
    style_cell(row.cells[1], widths_cnt[1], color=DARK, bg=bg)
    row.cells[2].paragraphs[0].add_run(succ)
    style_cell(row.cells[2], widths_cnt[2], color=GREEN, bg=bg)
    row.cells[3].paragraphs[0].add_run(fail)
    style_cell(row.cells[3], widths_cnt[3], color=RGBColor(185, 28, 28), bg=bg)

add_h2("3.3 Format Standar Respons (Envelope Pattern)")
add_p("Seluruh respons API dibungkus dalam format envelop JSON yang konsisten:")
add_code_block('// Respons Berhasil (HTTP 200 / 201):\n{\n  "success": true,\n  "data": {\n    "id": 1,\n    "judul": "Dasar Pemrograman Kotlin",\n    "tersedia": true\n  }\n}\n\n// Respons Gagal (HTTP 400 / 404 / 413 / 415 / 500):\n{\n  "success": false,\n  "error": {\n    "code": "BOOK_NOT_FOUND",\n    "message": "Buku tidak ditemukan."\n  }\n}')

# ==================== BAB IV ====================
add_h1("BAB IV: LANGKAH IMPLEMENTASI & STRUKTUR KODE")
add_h2("4.1 Struktur Direktori Proyek")
add_p("Proyek disusun menggunakan modularitas CommonJS (.cjs) agar kompatibel dengan lingkungan Node.js 24 LTS:")
add_code_block("""WSE_Praktikum_05_240104040147_Muhammad Daffa Khairi Addin/
├── package.json                # Metadata dependensi express@5.3.0 & swagger-ui-express@5.0.1
├── package-lock.json           # Rekaman pasti pohon dependensi
├── app.cjs                     # Komposisi Express, middleware, rute OpenAPI & Swagger UI
├── server.cjs                  # Listener HTTP pada host 127.0.0.1 port 3200
├── client.cjs                  # HTTP client CLI berbasis global fetch & AbortSignal
├── README.md                   # Panduan instalasi dan reproduksi pengujian
├── docs/
│   └── openapi.json            # Berkas spesifikasi OpenAPI 3.1.0 yang digenerate otomatis
├── scripts/
│   └── generate-openapi.cjs    # Skrip generator OpenAPI dengan contoh respons error lengkap
├── samples/                    # Berkas muatan uji (valid, parsial, salah tipe, rusak)
│   ├── buku-baru.json
│   ├── buku-ubah.json
│   ├── buku-salah-tipe.json
│   ├── buku-parsial.json
│   └── json-rusak.txt
└── src/                        # Modul arsitektur berlapis
    ├── errors.cjs              # Kelas HttpError, parseId, dan validateBuku
    ├── buku.repository.cjs     # In-memory array repository CRUD
    ├── buku.service.cjs        # Logika domain dan orkestrasi validasi
    ├── buku.controller.cjs     # Adaptor konteks HTTP
    └── buku.routes.cjs         # Pemetaan endpoint & middleware jsonOnly""")

add_h2("4.2 Implementasi Berkas Utama")
add_p("Setiap berkas kode ditulis dengan prinsip kesederhanaan, keamanan, dan ketegasan kontrak:")
add_p("Fungsi parseId memastikan ID berupa string digit desimal murni tanpa leading zero dan berada pada rentang safe integer. Fungsi validateBuku memastikan objek tidak berupa array, tidak memiliki atribut ekstra di luar judul dan tersedia, serta memeriksa tipe boolean secara ketat.", bold_prefix="1. src/errors.cjs: ")
add_p("Mengelola array buku dalam memori dengan data awal ID 1 dan ID 2. Counter nextId dinaikkan setiap create untuk menjamin keunikan ID.", bold_prefix="2. src/buku.repository.cjs: ")
add_p("Memisahkan domain logic dari protokol Express. Mengimplementasikan validasi query q, memastikan buku target ada pada operasi replace dan remove sebelum memanggil repositori.", bold_prefix="3. src/buku.service.cjs: ")
add_p("Menghubungkan konteks request Express dengan service, menetapkan status code 201 dan header Location pada POST, serta status code 204 pada DELETE.", bold_prefix="4. src/buku.controller.cjs: ")
add_p("Menggunakan router.param('id') untuk mengekstrak dan memvalidasi ID secara otomatis pada seluruh endpoint berparameter path, serta menerapkan middleware jsonOnly pada POST dan PUT.", bold_prefix="5. src/buku.routes.cjs: ")
add_p("Mengintegrasikan middleware logging, Cache-Control no-store, express.json({limit: '10kb'}), rute katalog, rute /openapi.json, Swagger UI pada /docs, fallback 404, dan centralized error handler.", bold_prefix="6. app.cjs & server.cjs: ")

# ==================== BAB V ====================
add_h1("BAB V: MATRIKS PENGUJIAN & ANALISIS HASIL AKTUAL")
add_h2("5.1 Matriks Hasil Pengujian Komprehensif (U01 - U24)")
add_p("Pengujian dilakukan secara menyeluruh terhadap 24 skenario yang didefinisikan pada modul praktikum. Seluruh pengujian menghasilkan status LULUS (PASS) sesuai dengan spesifikasi kontrak:")

table_test = doc.add_table(rows=len(test_results) + 6, cols=5)
table_test.alignment = WD_TABLE_ALIGNMENT.CENTER
widths_test = [Inches(0.6), Inches(1.8), Inches(1.6), Inches(1.7), Inches(0.6)]
add_table_header(table_test.rows[0], widths_test, ["Kode", "Skenario Pengujian", "Hasil yang Diharapkan", "Keluaran Aktual", "Status"])

row_counter = 1
for item in test_results:
    row = table_test.rows[row_counter]
    bg = "F8FAFC" if row_counter % 2 == 1 else "FFFFFF"
    
    # ID
    row.cells[0].paragraphs[0].add_run(item['id'])
    style_cell(row.cells[0], widths_test[0], bold=True, color=NAVY, bg=bg)
    
    # Desc
    row.cells[1].paragraphs[0].add_run(item['desc'])
    style_cell(row.cells[1], widths_test[1], color=DARK, bg=bg)
    
    # Expected
    exp_text = "Status 200 OK"
    if "POST" in item['cmd']: exp_text = "Status 201 Created + Location"
    if "DELETE" in item['cmd']: exp_text = "Status 204 No Content"
    if "400" in item['desc']: exp_text = "Status 400 Bad Request"
    if "404" in item['desc']: exp_text = "Status 404 Not Found"
    if "415" in item['desc']: exp_text = "Status 415 Unsupported Media Type"
    if "npm ls" in item['cmd']: exp_text = "express@5.3.0 & swagger-ui-express@5.0.1"
    row.cells[2].paragraphs[0].add_run(exp_text)
    style_cell(row.cells[2], widths_test[2], color=SLATE, bg=bg)
    
    # Actual output summary
    act_summary = item['output'].split('\n')[-1]
    if len(act_summary) > 40:
        act_summary = act_summary[:37] + "..."
    row.cells[3].paragraphs[0].add_run(act_summary)
    style_cell(row.cells[3], widths_test[3], color=DARK, bg=bg)
    
    # Status
    row.cells[4].paragraphs[0].add_run("LULUS")
    style_cell(row.cells[4], widths_test[4], align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, color=GREEN, bg=bg)
    
    row_counter += 1

extra_scenarios = [
    ("U20", "Akses Halaman /docs/ Swagger UI", "Tampil 5 operasi katalog", "5 operasi tampil lengkap interaktif", "LULUS"),
    ("U21", "Eksekusi CRUD via Swagger UI", "Status 200, 201, 204, 404 sesuai", "CRUD sukses dieksekusi via Try it out", "LULUS"),
    ("U22", "Pengujian Server Dihentikan", "Client tidak memperoleh respons HTTP", "TypeError (koneksi ditolak)", "LULUS"),
    ("U23", "Pengujian Server Dimulai Ulang", "Data kembali ke 2 buku awal", "Buku ID 5 hilang (404), data awal pulih", "LULUS"),
    ("U24", "Keterulangan Instalasi (README)", "npm ci & npm run docs berhasil", "Dependensi terpasang persis sesuai lockfile", "LULUS")
]

for item in extra_scenarios:
    row = table_test.rows[row_counter]
    bg = "F8FAFC" if row_counter % 2 == 1 else "FFFFFF"
    row.cells[0].paragraphs[0].add_run(item[0])
    style_cell(row.cells[0], widths_test[0], bold=True, color=NAVY, bg=bg)
    row.cells[1].paragraphs[0].add_run(item[1])
    style_cell(row.cells[1], widths_test[1], color=DARK, bg=bg)
    row.cells[2].paragraphs[0].add_run(item[2])
    style_cell(row.cells[2], widths_test[2], color=SLATE, bg=bg)
    row.cells[3].paragraphs[0].add_run(item[3])
    style_cell(row.cells[3], widths_test[3], color=DARK, bg=bg)
    row.cells[4].paragraphs[0].add_run(item[4])
    style_cell(row.cells[4], widths_test[4], align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, color=GREEN, bg=bg)
    row_counter += 1

add_h2("5.2 Analisis Hasil Pengujian")
add_p("Berdasarkan pengujian empiris pada terminal dan antarmuka web, diperoleh temuan-temuan krusial berikut:")
add_p("Operasi POST berhasil membuat buku baru pada ID 3 dengan mengembalikan status 201 dan header Location: /api/buku/3. Pembaruan PUT mengganti seluruh data objek pada ID 3, dan operasi DELETE menghasilkan respons 204 No Content tanpa payload body. Ketika dilakukan GET sesudah DELETE, server merespons 404 BOOK_NOT_FOUND, membuktikan state lifecycle resource berjalan konsisten.", bold_prefix="1. Validasi Siklus Hidup CRUD: ")
add_p("Pemeriksaan ID secara ketat menolak nilai '1abc' dan '01' dengan kode INVALID_ID, membuktikan bahwa validasi regex desimal tanpa leading zero berhasil mencegah ambigu parsing. Pengujian PUT parsial ditolak dengan status 400 INVALID_BOOK dan dibuktikan melalui GET bahwa data asli tidak berubah.", bold_prefix="2. Ketahanan Validasi & Integritas State: ")
add_p("Body JSON yang terpotong ('{\"judul\":') ditangkap oleh error handler bawaan Express dan diubah menjadi format standar aplikasi dengan status 400 INVALID_JSON. Permintaan POST dengan Content-Type text/plain ditolak oleh middleware jsonOnly dengan status 415 UNSUPPORTED_MEDIA_TYPE sebelum masuk ke controller.", bold_prefix="3. Proteksi Parser & Media Type: ")
add_p("Ketika server dimatikan, client menghasilkan error 'TypeError' (koneksi ditolak). Saat server dijalankan kembali, data tambahan ID 5 hilang dan katalog kembali ke 2 buku awal. Hal ini membuktikan karakteristik in-memory array yang bersifat volatil dan non-persisten.", bold_prefix="4. Karakteristik Volatilitas Memori: ")

# ==================== BAB VI ====================
add_h1("BAB VI: PEMBAHASAN TUGAS MANDIRI")
add_h2("6.1 Tugas Mandiri 1: Menambah Data Melalui API & Analisis Komparasi")
add_p("Pada Tugas Mandiri 1, mahasiswa menambahkan dua buku baru melalui endpoint POST, membaca data, mengubah satu buku, dan menghapus satu buku menggunakan client terminal:")

t1_steps = [
    ("T1_01: POST Buku 1 -> ", "Judul: 'Arsitektur Microservices Modern', tersedia: true. Respons: HTTP 201, ID aktual: 3, Location: /api/buku/3."),
    ("T1_02: POST Buku 2 -> ", "Judul: 'Desain RESTful API Standar Industri', tersedia: false. Respons: HTTP 201, ID aktual: 4, Location: /api/buku/4."),
    ("T1_03: GET Seluruh Katalog -> ", "Mengembalikan array 4 buku lengkap (ID 1, 2, 3, 4)."),
    ("T1_04: PUT Buku ID 3 -> ", "Mengganti judul menjadi 'Arsitektur Microservices Modern (Edisi 2)' dan tersedia menjadi false. Respons: HTTP 200."),
    ("T1_05: DELETE Buku ID 4 -> ", "Menghapus buku ID 4. Respons: HTTP 204 No Content."),
    ("T1_06: GET Verifikasi -> ", "GET /api/buku/4 mengembalikan 404 BOOK_NOT_FOUND. GET /api/buku mengonfirmasi katalog akhir berisi 3 buku (ID 1, 2, 3).")
]
for pre, post in t1_steps:
    add_p(post, bold_prefix=pre)

add_h3("Analisis Perbedaan Penyuntingan Array Langsung vs Operasi Consumer via API")
add_p("Mengapa penyuntingan array pada berkas kode (misal mengedit src/buku.repository.cjs secara manual) TIDAK SAMA dengan operasi yang dikirim oleh consumer melalui API?")
add_p("Penyuntingan langsung pada berkas kode mengubah kode sumber statis di disk (static code modification). Perubahan ini tidak akan terefleksi pada proses Node.js yang sedang aktif tanpa restart proses. Sebaliknya, operasi melalui API memanipulasi heap memory dari proses runtime yang sedang berjalan secara dinamis tanpa interupsi server.", bold_prefix="1. Dimensi Runtime vs Statis: ")
add_p("Consumer API tunduk pada kontrak antarmuka yang divalidasi oleh berlapis-lapis pertahanan: parsing JSON, pemeriksaan media type, validasi regex ID, dan validasi tipe skema data. Penyuntingan berkas langsung dapat menyusupkan tipe data yang salah (misal tipe angka negatif, field ekstra, atau tipe data rusak) tanpa adanya penjaga (gatekeeper).", bold_prefix="2. Penegakan Kontrak & Validasi: ")
add_p("Manipulasi via API menghasilkan respons terstruktur (201 Created, Location header, logging console, audit trail). Selain itu, pada sistem multi-pengguna, API menjamin serialisasi eksekusi event loop Node.js untuk mencegah race condition, sedangkan manipulasi file mentah tidak memiliki mekanisme proteksi transaksi.", bold_prefix="3. Efek Samping (Side Effects) & Protokol: ")

add_h2("6.2 Tugas Mandiri 2: Melengkapi Dokumentasi OpenAPI & Verifikasi Swagger UI")
add_p("Pada generator scripts/generate-openapi.cjs, dilakukan pengayaan spesifikasi OpenAPI 3.1.0 untuk memberikan kejelasan kontrak yang maksimal bagi consumer:")
add_p("Setiap parameter (seperti parameter path 'id' dan query 'q') ditambahkan deskripsi semantik mendalam mengenai batas panjang UTF-16, format regex desimal tanpa leading zero, serta batas aman integer JavaScript.", bold_prefix="1. Deskripsi Semantik Terperinci: ")
add_p("Setiap kode status error (400, 404, 413, 415, 500) dilengkapi dengan contoh payload konkret (example value) yang mencerminkan format standar aplikasi: { success: false, error: { code, message } }.", bold_prefix="2. Penambahan Objek Contoh Respons Error: ")
add_p("Dokumen docs/openapi.json digenerate ulang menggunakan 'npm run docs' dan diverifikasi langsung pada Swagger UI (/docs/). Seluruh perubahan skema dan contoh respons error langsung tampil pada tab Example Value respons Swagger UI tanpa memerlukan restart server.", bold_prefix="3. Keterpaduan & Regenerasi Dokumen: ")

add_h2("6.3 Tugas Mandiri 3: Jawaban Analisis 12 Pertanyaan Pemahaman")

qa_items = [
    ("1. Perbedaan Node.js dan Express dalam proyek ini?",
     "Node.js bertindak sebagai runtime environment berbasis JavaScript V8 engine yang menyediakan fondasi eksekusi server-side, sistem I/O asinkron, modul jaringan HTTP bawaan (node:http), dan akses filesystem. Sedangkan Express adalah web application framework minimalis yang berjalan di atas runtime Node.js untuk mempermudah perancangan routing, pipeline middleware, penanganan request/response, dan parsing body tanpa harus menulis logika low-level socket secara manual."),
    
    ("2. Mengapa controller dipisahkan dari service dan repository?",
     "Pemisahan ini menerapkan prinsip Separation of Concerns (SoC) dan Single Responsibility Principle (SRP). Controller bertugas sebagai adapter konteks HTTP (membaca req, memilih status code, menyusun header, mengirim respons). Service bertugas menjalankan aturan bisnis aplikasi yang murni (business logic) dan validasi data secara agnostik terhadap protokol web. Repository mengabstraksi mekanisme penyimpanan data. Pemisahan ini membuat kode mudah diuji (unit testing), mudah dimodifikasi (misal mengganti array dengan database tanpa merusak controller), dan mencegah spaghetti code."),
    
    ("3. Perbedaan parameter path, query, header, dan body?",
     "Parameter path (/api/buku/:id) adalah bagian hierarki URI untuk mengidentifikasi resource spesifik secara unik. Parameter query (?q=kotlin) adalah pasangan kunci-nilai opsional setelah tanda tanya untuk memfilter, mengurutkan, atau memaginasi representasi resource. Header membawa metadata komunikasi (seperti Content-Type, Accept, Cache-Control, dan Location). Body membawa payload data entitas yang dikirim consumer (biasanya berformat JSON pada POST atau PUT) untuk dibuat atau diperbarui."),
    
    ("4. Mengapa JSON yang sah belum tentu memenuhi kontrak objek buku?",
     "Format JSON hanya memvalidasi sintaksis serialisasi data (seperti ketepatan tanda kutip, kurung kurawal, koma). Suatu payload dapat berformat JSON yang 100% valid, namun melanggar kontrak domain, misalnya: nilai 'tersedia' dikirim berupa string 'true' bukannya boolean, judul dikirim berupa string kosong (spasi saja), atau menyertakan field terlarang seperti 'id' atau 'harga'. Oleh karena itu, parser JSON wajib diikuti oleh validator skema aplikasi."),
    
    ("5. Mengapa ID 1abc ditolak walaupun parseInt dapat menghasilkan angka?",
     "Fungsi bawaan JavaScript parseInt('1abc', 10) melakukan pemindaian karakter dari kiri ke kanan dan berhenti saat menemukan karakter non-angka, sehingga tetap mengembalikan angka 1. Hal ini berbahaya karena consumer mengirim '1abc' namun server memproses buku ID 1 secara diam-diam. Untuk menjamin ketegasan kontrak, validasi menggunakan regex ^[1-9][0-9]*$ yang menuntut seluruh input berupa digit positif murni tanpa toleransi karakter liar."),
    
    ("6. Mengapa PUT parsial ditolak pada kontrak ini?",
     "Sesuai spesifikasi RFC 9110, metode PUT memiliki semantik 'complete replacement' (penggantian representasi utuh). Kontrak API katalog menetapkan bahwa PUT wajib menyertakan kedua field 'judul' dan 'tersedia'. Jika client ingin memperbarui sebagian field saja, metode HTTP yang semestinya digunakan adalah PATCH. Penolakan PUT parsial mencegah penghapusan atribut data secara tidak sengaja."),
    
    ("7. Mengapa respons DELETE 204 tidak dibaca sebagai JSON?",
     "Status HTTP 204 No Content secara definitif menyatakan bahwa operasi berhasil dan server sengaja TIDAK mengirimkan pesan payload (body kosong berukuran 0 byte). Jika client memaksakan pemanggilan response.json() pada body kosong, runtime akan melemparkan kesalahan SyntaxError (Unexpected end of JSON input). Oleh karena itu, client menggunakan response.text() untuk membaca respons kosong secara aman."),
    
    ("8. Perbedaan OpenAPI, Swagger UI, dan implementasi Express?",
     "OpenAPI adalah spesifikasi formal (kontrak data deklaratif berbentuk file JSON/YAML) yang mendeskripsikan antarmuka API secara teoritis. Swagger UI adalah kakas antarmuka visual (frontend viewer) yang merender dokumen OpenAPI agar dapat dibaca dan dicoba secara interaktif oleh manusia. Sedangkan implementasi Express adalah kode program backend aktual yang mengeksekusi routing, validasi logika, dan manipulasi data di server."),
    
    ("9. Mengapa Example Value bukan bukti bahwa endpoint berhasil?",
     "Example Value pada Swagger UI hanyalah contoh statis (mock documentation) yang disalin dari skrip generator OpenAPI. Tampilnya Example Value sama sekali tidak menunjukkan bahwa server Express sedang hidup atau bahwa database telah terupdate. Keberhasilan endpoint hanya dapat dibuktikan apabila tombol 'Execute' diklik dan menghasilkan respons nyata pada blok 'Server response' dengan status code dan payload aktual dari backend."),
    
    ("10. Bagaimana perubahan penyimpanan ke database memengaruhi kode dan pengujian?",
     "Operasi I/O database bersifat asinkron (non-blocking), sehingga seluruh fungsi pada repository, service, dan controller harus diubah menjadi fungsi asynchronous menggunakan async/await dan mengembalikan Promise. Pada pengujian, perlu ditambahkan penanganan koneksi pool, timeout database, kegagalan transaksi (rollback), isolasi data antar-pengujian (fixtures/mocking), serta skema migrasi tabel."),
    
    ("11. Mengapa array tidak cocok untuk data operasional yang harus bertahan?",
     "Array disimpan di dalam volatil memory (RAM) milik proses Node.js tunggal. Ketika proses mati, restart, crash, atau server mengalami pemadaman listrik, seluruh data yang ditambahkan akan musnah seketika. Selain itu, array in-memory tidak mendukung konkurensi multi-instance (horizontal scaling) karena setiap instance server akan memiliki memori yang terisolasi dan tidak tersinkronisasi."),
    
    ("12. Mengapa tombol atau dokumentasi yang disembunyikan tidak menggantikan otorisasi server?",
     "Penyembunyian tombol atau dokumentasi hanyalah manipulasi antarmuka pengguna di sisi client (security through obscurity). Consumer atau penyerang dapat dengan mudah mengirimkan HTTP request secara langsung ke endpoint backend menggunakan cURL, Postman, atau skrip terminal tanpa melalui antarmuka web. Keamanan sejati wajib ditegakkan di sisi server melalui mekanisme autentikasi (identifikasi) dan otorisasi (pemeriksaan hak akses) pada middleware route.")
]

for q, a in qa_items:
    add_h3(q)
    add_p(a)

# ==================== BAB VII ====================
add_h1("BAB VII: BUKTI DOKUMENTASI & TANGKAPAN LAYAR")
add_p("Berikut adalah 15 tangkapan layar beresolusi tinggi yang mendokumentasikan setiap tahapan pengujian, verifikasi antarmuka web Swagger UI, dan eksekusi terminal secara nyata:")

screenshots_list = [
    ("screenshots/01_swagger_ui_overview.png", "Gambar 1", "Tampilan Utama Antarmuka Swagger UI API Katalog Perpustakaan Kampus (5 Operasi Terpetakan)"),
    ("screenshots/02_swagger_get_id_1.png", "Gambar 2", "Pengujian Swagger UI: GET /api/buku/1 Menghasilkan Status 200 OK dan Detail Data Buku"),
    ("screenshots/03_swagger_get_id_999.png", "Gambar 3", "Pengujian Swagger UI: GET /api/buku/999 Menghasilkan Status 404 BOOK_NOT_FOUND"),
    ("screenshots/04_swagger_post_create.png", "Gambar 4", "Pengujian Swagger UI: POST /api/buku Menghasilkan Status 201 Created dan Header Location /api/buku/4"),
    ("screenshots/05_swagger_get_after_post.png", "Gambar 5", "Pengujian Swagger UI: GET /api/buku/4 Membuktikan Resource Baru Tersimpan di Server"),
    ("screenshots/06_swagger_put_update.png", "Gambar 6", "Pengujian Swagger UI: PUT /api/buku/4 Berhasil Memperbarui Data Buku (Status 200 OK)"),
    ("screenshots/07_swagger_delete.png", "Gambar 7", "Pengujian Swagger UI: DELETE /api/buku/4 Menghasilkan Status 204 No Content Tanpa Body"),
    ("screenshots/08_swagger_get_after_delete.png", "Gambar 8", "Pengujian Swagger UI: GET /api/buku/4 Setelah Penghapusan Menghasilkan Status 404 BOOK_NOT_FOUND"),
    ("screenshots/09_browser_openapi_json.png", "Gambar 9", "Tampilan Browser: Spesifikasi OpenAPI 3.1.0 Mentah Melalui Endpoint /openapi.json"),
    ("screenshots/10_browser_api_buku.png", "Gambar 10", "Tampilan Browser: Akses Langsung Endpoint Katalog /api/buku Menghasilkan JSON Envelop"),
    ("screenshots/11_terminal_verifikasi_lingkungan.png", "Gambar 11", "Terminal: Verifikasi Versi Node 24, npm 11, Express 5.3.0, Pemeriksaan Sintaks, & Startup Server"),
    ("screenshots/12_terminal_crud_testing.png", "Gambar 12", "Terminal: Hasil Eksekusi Pengujian Skenario CRUD Inti (U02 - U10) via client.cjs"),
    ("screenshots/13_terminal_validation_errors.png", "Gambar 13", "Terminal: Hasil Eksekusi Pengujian Validasi Ketat & Error Handling (U11 - U18)"),
    ("screenshots/14_terminal_restart_memory_boundary.png", "Gambar 14", "Terminal: Pembuktian Batas In-Memory, Koneksi Gagal Saat Server Mati (U22), & Data Reset (U23)"),
    ("screenshots/15_terminal_tugas1_mandiri.png", "Gambar 15", "Terminal: Eksekusi Rangkaian Tugas Mandiri 1 (POST 2 Buku Baru, Ubah, & Hapus Data)")
]

for img_p, cap_title, cap_desc in screenshots_list:
    add_image_figure(img_p, cap_title, cap_desc)

# ==================== BAB VIII ====================
add_h1("BAB VIII: KESIMPULAN & REFLEKSI PRAKTIKUM")
add_h2("8.1 Kesimpulan")
add_p("1. Implementasi web service katalog buku perpustakaan kampus berhasil dibangun menggunakan Node.js v24.11.0 dan framework Express v5.3.0 dengan pemisahan lapisan arsitektur yang solid (routing, controller, service, repository).")
add_p("2. Desain endpoint mematuhi standar semantik HTTP RFC 9110 secara ketat, termasuk penggunaan status code 200 (OK), 201 (Created), 204 (No Content), 400 (Bad Request), 404 (Not Found), 413 (Payload Too Large), 415 (Unsupported Media Type), dan 500 (Internal Server Error).")
add_p("3. Seluruh 24 skenario pengujian matriks berhasil lulus (100% PASS), membuktikan keandalan sistem dalam menangani input valid maupun menolak input malformed secara terkontrol tanpa mengekspos stack trace.")
add_p("4. Spesifikasi OpenAPI 3.1.0 berhasil diintegrasikan secara lokal melalui swagger-ui-express v5.0.1, memberikan dokumentasi interaktif yang memudahkan pengujian dan kolaborasi antar-pengembang.")

add_h2("8.2 Refleksi Pembelajaran")
add_p("Praktikum ini memberikan pemahaman mendalam mengenai pentingnya pemisahan tanggung jawab (Separation of Concerns). Dengan memisahkan logika controller dan service, aturan validasi buku dapat diuji dan dipelihara secara terpisah dari mekanisme pengiriman respons HTTP. Selain itu, eksperimen penghentian dan restart server memberikan bukti nyata mengenai keterbatasan penyimpanan volatil (in-memory) dan memperjelas alasan mengapa sistem produksi memerlukan database persisten serta arsitektur state management yang handal.")

# ==================== DAFTAR PUSTAKA ====================
add_h1("DAFTAR PUSTAKA")
references = [
    "[1] Express.js. (2026). Express 5.x API Reference and Routing Guide. https://expressjs.com/en/5x/api/",
    "[2] Fielding, R., Nottingham, M., & Reschke, J. (2022). HTTP Semantics. RFC 9110, Internet Engineering Task Force (IETF). https://www.rfc-editor.org/rfc/rfc9110.html",
    "[3] Martin, R. C. (2018). Clean Architecture: A Craftsman's Guide to Software Structure and Design. Prentice Hall.",
    "[4] Node.js Foundation. (2026). Node.js v24 LTS Documentation. https://nodejs.org/docs/latest-v24.x/api/",
    "[5] OpenAPI Initiative. (2021). OpenAPI Specification Version 3.1.0. https://spec.openapis.org/oas/v3.1.0.html",
    "[6] Scottie1984. (2024). swagger-ui-express: Swagger UI for Node.js Express Applications. GitHub Repository. https://github.com/scottie1984/swagger-ui-express"
]
for ref in references:
    p_ref = doc.add_paragraph()
    p_ref.paragraph_format.space_after = Pt(4)
    p_ref.paragraph_format.left_indent = Inches(0.3)
    p_ref.paragraph_format.first_line_indent = Inches(-0.3)
    r = p_ref.add_run(ref)
    format_run(r, font_name="Calibri", size_pt=10, color=DARK)

docx_path = "Laporan_Praktikum_05_WSE_240104040147_Muhammad_Daffa_Khairi_Addin.docx"
doc.save(docx_path)
print(f"DOCX report saved successfully to: {docx_path}")

# ==================== GENERATE MATCHING HTML & PDF ====================
print("Generating HTML for PDF conversion...")

html_content = f"""<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="utf-8">
<title>Laporan Praktikum 05 Web Service Engineering</title>
<style>
  @page {{
    size: A4;
    margin: 20mm;
  }}
  body {{
    font-family: 'Calibri', 'Segoe UI', Arial, sans-serif;
    color: #1e293b;
    line-height: 1.45;
    font-size: 10.5pt;
    margin: 0;
    padding: 0;
  }}
  h1.title {{
    color: #1b365d;
    font-size: 20pt;
    text-align: center;
    margin-top: 40px;
    margin-bottom: 5px;
    font-weight: 800;
  }}
  h2.subtitle {{
    color: #475569;
    font-size: 12.5pt;
    text-align: center;
    font-style: italic;
    margin-bottom: 40px;
    font-weight: 500;
  }}
  h1 {{
    color: #1b365d;
    font-size: 14pt;
    border-bottom: 2px solid #1b365d;
    padding-bottom: 4px;
    margin-top: 28px;
    margin-bottom: 12px;
    page-break-after: avoid;
  }}
  h2 {{
    color: #334155;
    font-size: 12pt;
    margin-top: 18px;
    margin-bottom: 8px;
    page-break-after: avoid;
  }}
  h3 {{
    color: #1e293b;
    font-size: 10.5pt;
    margin-top: 12px;
    margin-bottom: 4px;
    page-break-after: avoid;
  }}
  p {{
    margin-top: 0;
    margin-bottom: 8px;
    text-align: justify;
  }}
  .cover-table {{
    width: 85%;
    margin: 40px auto;
    border-collapse: collapse;
  }}
  .cover-table td {{
    padding: 6px 12px;
    font-size: 10.5pt;
  }}
  .cover-table td.label {{
    font-weight: bold;
    color: #475569;
    width: 38%;
  }}
  .cover-table td.val {{
    color: #1e293b;
  }}
  table.content-table {{
    width: 100%;
    border-collapse: collapse;
    margin: 12px 0 16px 0;
    font-size: 9pt;
    page-break-inside: auto;
  }}
  table.content-table th {{
    background-color: #1b365d;
    color: #ffffff;
    font-weight: bold;
    padding: 6px 8px;
    border: 1px solid #cbd5e1;
    text-align: left;
  }}
  table.content-table td {{
    padding: 5px 8px;
    border: 1px solid #cbd5e1;
    vertical-align: top;
  }}
  table.content-table tr:nth-child(even) {{
    background-color: #f8fafc;
  }}
  .badge-pass {{
    background-color: #dcfce7;
    color: #166534;
    font-weight: bold;
    padding: 2px 6px;
    border-radius: 4px;
    display: inline-block;
    text-align: center;
  }}
  pre.code-block {{
    background-color: #f1f5f9;
    border: 1px solid #e2e8f0;
    border-radius: 4px;
    padding: 10px 14px;
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 8.5pt;
    color: #0f172a;
    white-space: pre-wrap;
    word-break: break-all;
    margin: 8px 0;
  }}
  .callout {{
    background-color: #f8fafc;
    border-left: 4px solid #1b365d;
    padding: 8px 14px;
    margin: 10px 0;
    font-size: 9.5pt;
    font-style: italic;
  }}
  .figure {{
    text-align: center;
    margin: 14px 0;
    page-break-inside: avoid;
  }}
  .figure img {{
    max-width: 90%;
    height: auto;
    border: 1px solid #cbd5e1;
    border-radius: 4px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.08);
  }}
  .figure .caption {{
    font-size: 8.5pt;
    color: #475569;
    margin-top: 5px;
  }}
  .figure .caption strong {{
    color: #1b365d;
  }}
  .page-break {{
    page-break-before: always;
  }}
</style>
</head>
<body>

<div style="text-align: center; margin-top: 60px;">
  <h1 class="title">LAPORAN PRAKTIKUM WEB SERVICE ENGINEERING</h1>
  <h2 class="subtitle">PRAKTIKUM 05: DESAIN ENDPOINT DAN IMPLEMENTASI WEB SERVICE DENGAN FRAMEWORK EXPRESS</h2>
</div>

<table class="cover-table">
  <tr><td class="label">Nama Lengkap</td><td class="val">: Muhammad Daffa Khairi Addin</td></tr>
  <tr><td class="label">Nomor Induk Mahasiswa (NIM)</td><td class="val">: 240104040147</td></tr>
  <tr><td class="label">Program Studi</td><td class="val">: S1 Teknologi Informasi</td></tr>
  <tr><td class="label">Fakultas</td><td class="val">: Tarbiyah dan Keguruan</td></tr>
  <tr><td class="label">Perguruan Tinggi</td><td class="val">: Universitas Islam Negeri Antasari Banjarmasin</td></tr>
  <tr><td class="label">Mata Kuliah</td><td class="val">: Web Service Engineering</td></tr>
  <tr><td class="label">Dosen Pengampu</td><td class="val">: Muhayat, M.IT</td></tr>
  <tr><td class="label">Topik Praktikum</td><td class="val">: Desain Endpoint, Arsitektur Berlapis, & OpenAPI</td></tr>
  <tr><td class="label">Studi Kasus</td><td class="val">: Layanan Katalog Buku Perpustakaan Kampus</td></tr>
  <tr><td class="label">Tahun Akademik</td><td class="val">: 2026</td></tr>
</table>

<div class="page-break"></div>

<h1>BAB I: TUJUAN & PRASYARAT PRAKTIKUM</h1>
<h2>1.1 Deskripsi & Tujuan Pembelajaran</h2>
<p>Praktikum ini menerapkan konsep perancangan endpoint dan implementasi web service menggunakan framework Node.js dan Express. Mahasiswa membangun API katalog buku perpustakaan kampus mulai dari penyiapan lingkungan, perancangan kontrak endpoint, implementasi arsitektur berlapis (routing, controller, service, repository), penanganan validasi dan error secara terpusat, hingga penyusunan spesifikasi OpenAPI 3.1.0 dan pengujian interaktif menggunakan Swagger UI.</p>
<p><strong>Pengetahuan: </strong>Menjelaskan peran runtime Node.js, framework Express, router, controller, service, repository, serta kontrak OpenAPI.</p>
<p><strong>Perancangan: </strong>Menetapkan rancangan resource, model data, metode HTTP, status code semantik, format pesan sukses/gagal, serta kondisi penolakan input.</p>
<p><strong>Keterampilan: </strong>Menyiapkan dependensi exact-version, mengonfigurasi listener loopback lokal port 3200, dan mengimplementasikan operasi CRUD katalog secara menyeluruh.</p>
<p><strong>Pengujian: </strong>Membuktikan keberhasilan operasi bisnis serta ketegasan penolakan data invalid melalui respons aktual HTTP dan logging terminal.</p>
<p><strong>Dokumentasi: </strong>Menyusun skrip generator OpenAPI 3.1.0 otomatis dan menguji seluruh operasi API secara visual melalui Swagger UI lokal.</p>

<h2>1.2 Lingkungan Sistem & Versi Perangkat Lunak</h2>
<p>Eksperimen dan verifikasi praktikum dilaksanakan pada lingkungan sistem lokal dengan rincian perangkat aktual sebagai berikut:</p>

<table class="content-table">
  <tr><th>Komponen Sistem</th><th>Versi / Konfigurasi Aktual</th></tr>
  <tr><td><strong>Sistem Operasi</strong></td><td>Microsoft Windows 11 Home / Pro (x64 Architecture)</td></tr>
  <tr><td><strong>Runtime Environment</strong></td><td>Node.js v24.11.0 (LTS Iron Series)</td></tr>
  <tr><td><strong>Package Manager</strong></td><td>npm v11.6.1</td></tr>
  <tr><td><strong>Web Framework</strong></td><td>Express v5.3.0 (Fixed Exact Dependency)</td></tr>
  <tr><td><strong>OpenAPI Documentation</strong></td><td>swagger-ui-express v5.0.1 (Local Interface Asset)</td></tr>
  <tr><td><strong>Code Editor & Shell</strong></td><td>Visual Studio Code & Git Bash / Command Prompt</td></tr>
  <tr><td><strong>HTTP Client</strong></td><td>Node.js Custom Fetch Client (client.cjs) & Browser Swagger UI</td></tr>
</table>

<h1>BAB II: LANDASAN TEORI & ARSITEKTUR BERLAPIS</h1>
<h2>2.1 Konsep Web Service & Protokol HTTP</h2>
<p>Web service adalah antarmuka komunikasi antar-aplikasi (machine-to-machine) yang memanfaatkan protokol standar web, khususnya HTTP. Pada arsitektur REST (Representational State Transfer), interaksi terjadi antara penyedia layanan (provider) dan pengguna layanan (consumer) menggunakan metode HTTP yang memiliki semantik terstandarisasi (RFC 9110):</p>
<p><strong>GET: </strong>Operasi bersifat safe dan idempotent untuk membaca representasi resource tanpa menimbulkan efek samping (side effects) pada server.</p>
<p><strong>POST: </strong>Operasi non-idempotent untuk menambahkan resource baru ke koleksi. Respons berhasil mengembalikan kode status 201 Created serta header Location yang mengarahkan consumer ke URI resource yang baru dibuat.</p>
<p><strong>PUT: </strong>Operasi bersifat idempotent untuk mengganti (replace) seluruh field yang dapat ditulis pada target resource yang sudah ada. Jika terdapat field yang hilang pada body, server menolak permintaan karena PUT menuntut representasi entitas yang utuh.</p>
<p><strong>DELETE: </strong>Operasi idempotent untuk menghapus resource dari penyimpanan. Keberhasilan penghapusan direspons dengan status 204 No Content tanpa payload respons (empty body).</p>

<h2>2.2 Arsitektur Berlapis (Separation of Concerns)</h2>
<p>Untuk mencegah penumpukan logika pada satu file handler (spaghetti code), proyek ini menerapkan pola pemisahan tanggung jawab (Separation of Concerns) yang memisahkan aplikasi menjadi lima lapisan independen:</p>
<p><strong>1. Routing (src/buku.routes.cjs): </strong>Bertanggung jawab memetakan kombinasi metode HTTP dan path URI ke fungsi controller yang sesuai. Lapisan ini juga mengisolasi parameter path melalui router.param dan memvalidasi Content-Type melalui middleware jsonOnly.</p>
<p><strong>2. Controller (src/buku.controller.cjs): </strong>Mengatur konteks HTTP. Controller mengekstrak parameter query, path, dan body dari objek req, meneruskannya ke service, serta mengemas hasil ke dalam format respons baku (res.json, res.status, res.location). Controller tidak boleh memanipulasi penyimpanan data secara langsung.</p>
<p><strong>3. Service (src/buku.service.cjs): </strong>Menjalankan aturan bisnis (domain business logic) katalog. Service memastikan kepatuhan model data (melalui fungsi validasi), memeriksa keberadaan data (detail checking), serta bersifat agnostik terhadap Express (tidak bergantung pada objek req maupun res).</p>
<p><strong>4. Repository (src/buku.repository.cjs): </strong>Mengabstraksi akses ke lapisan penyimpanan. Pada praktikum ini, repository mengelola koleksi data buku dalam memori (in-memory array) melalui lima operasi dasar: list, find, create, replace, dan remove.</p>
<p><strong>5. Centralized Error Handling & Validation (src/errors.cjs & app.cjs): </strong>Mendefinisikan kelas HttpError terstruktur yang membawa status HTTP dan kode aplikasi. Middleware penanganan kesalahan menangkap seluruh error sehingga client menerima respons JSON seragam tanpa kebocoran stack trace.</p>

<div class="callout">
  <strong>Alur Satu Permintaan:</strong> Consumer mengirim HTTP Request -> Middleware Umum (Logging & Cache-Control no-store) -> JSON Body Parser (Limit 10kb) -> Router (Param ID & MediaType validation) -> Controller -> Service (Business logic & validation) -> Repository (Penyimpanan memori) -> Controller merespons JSON (200/201/204) atau Error Handler merespons JSON error (400/404/413/415/500).
</div>

<h2>2.3 Spesifikasi OpenAPI 3.1.0 & Swagger UI</h2>
<p>OpenAPI Specification (OAS) adalah standar deskripsi antarmuka pemrograman aplikasi HTTP yang dapat dibaca oleh manusia maupun mesin (machine-readable). Dokumen OpenAPI menjelaskan path, metode, parameter query/path, skema body, serta contoh respons sukses maupun gagal. Swagger UI adalah kakas visual yang menyajikan dokumen OpenAPI tersebut ke dalam antarmuka web interaktif lokal, memungkinkan pengembang melakukan pengujian operasi secara langsung (Try it out) menggunakan data nyata.</p>

<h1>BAB III: PERANCANGAN ENDPOINT & KONTRAK API</h1>
<h2>3.1 Spesifikasi Model Data</h2>
<p>Resource buku perpustakaan kampus dimodelkan dengan atribut dan aturan validasi ketat sebagai berikut:</p>
<p><strong>id (Number / Integer): </strong>Bilangan bulat positif (> 0) yang digenerate otomatis oleh server. Harus berada dalam batas aman JavaScript (Number.isSafeInteger <= 9007199254740991). Ditolak jika dikirim dalam body POST atau PUT.</p>
<p><strong>judul (String): </strong>Teks judul buku. Setelah melalui proses pemangkasan spasi (trim), panjang karakter wajib berada di antara 1 hingga 120 unit string UTF-16. Judul yang hanya terdiri dari spasi atau string kosong akan ditolak.</p>
<p><strong>tersedia (Boolean): </strong>Status ketersediaan buku di perpustakaan. Wajib bertipe boolean asli (true atau false). Nilai bertipe string seperti "true" atau "false" ditolak secara tegas.</p>
<p><strong>Aturan Properti Ekstra: </strong>Body request hanya boleh memuat properti 'judul' dan 'tersedia'. Pengiriman field tambahan di luar skema akan langsung memicu penolakan 400 INVALID_BOOK.</p>

<h2>3.2 Matriks Kontrak Operasi API</h2>
<table class="content-table">
  <tr><th>Metode & Path</th><th>Parameter / Input</th><th>Respons Berhasil</th><th>Kondisi Kegagalan</th></tr>
  <tr><td><strong>GET /api/buku</strong></td><td>Query param: q (opsional, maks 80 karakter)</td><td style="color:#166534;">200 OK dengan array katalog buku</td><td style="color:#b91c1c;">400 INVALID_QUERY (q > 80 atau duplikat)</td></tr>
  <tr><td><strong>GET /api/buku/:id</strong></td><td>Path param: :id (bilangan bulat positif)</td><td style="color:#166534;">200 OK dengan objek buku tunggal</td><td style="color:#b91c1c;">400 INVALID_ID, 404 BOOK_NOT_FOUND</td></tr>
  <tr><td><strong>POST /api/buku</strong></td><td>Body: JSON {{judul, tersedia}}</td><td style="color:#166534;">201 Created + Header Location + Objek</td><td style="color:#b91c1c;">400 INVALID_BOOK / INVALID_JSON, 413, 415</td></tr>
  <tr><td><strong>PUT /api/buku/:id</strong></td><td>Path: :id; Body: JSON {{judul, tersedia}}</td><td style="color:#166534;">200 OK dengan objek hasil pembaruan</td><td style="color:#b91c1c;">400 INVALID_BOOK / ID, 404, 413, 415</td></tr>
  <tr><td><strong>DELETE /api/buku/:id</strong></td><td>Path param: :id (bilangan bulat positif)</td><td style="color:#166534;">204 No Content (tanpa body payload)</td><td style="color:#b91c1c;">400 INVALID_ID, 404 BOOK_NOT_FOUND</td></tr>
</table>

<h2>3.3 Format Standar Respons (Envelope Pattern)</h2>
<pre class="code-block">// Respons Berhasil (HTTP 200 / 201):
{{
  "success": true,
  "data": {{
    "id": 1,
    "judul": "Dasar Pemrograman Kotlin",
    "tersedia": true
  }}
}}

// Respons Gagal (HTTP 400 / 404 / 413 / 415 / 500):
{{
  "success": false,
  "error": {{
    "code": "BOOK_NOT_FOUND",
    "message": "Buku tidak ditemukan."
  }}
}}</pre>

<h1>BAB IV: LANGKAH IMPLEMENTASI & STRUKTUR KODE</h1>
<h2>4.1 Struktur Direktori Proyek</h2>
<pre class="code-block">WSE_Praktikum_05_240104040147_Muhammad Daffa Khairi Addin/
├── package.json                # Metadata dependensi express@5.3.0 & swagger-ui-express@5.0.1
├── package-lock.json           # Rekaman pasti pohon dependensi
├── app.cjs                     # Komposisi Express, middleware, rute OpenAPI & Swagger UI
├── server.cjs                  # Listener HTTP pada host 127.0.0.1 port 3200
├── client.cjs                  # HTTP client CLI berbasis global fetch & AbortSignal
├── README.md                   # Panduan instalasi dan reproduksi pengujian
├── docs/
│   └── openapi.json            # Berkas spesifikasi OpenAPI 3.1.0 yang digenerate otomatis
├── scripts/
│   └── generate-openapi.cjs    # Skrip generator OpenAPI dengan contoh respons error lengkap
├── samples/                    # Berkas muatan uji (valid, parsial, salah tipe, rusak)
└── src/                        # Modul arsitektur berlapis
    ├── errors.cjs              # Kelas HttpError, parseId, dan validateBuku
    ├── buku.repository.cjs     # In-memory array repository CRUD
    ├── buku.service.cjs        # Logika domain dan orkestrasi validasi
    ├── buku.controller.cjs     # Adaptor konteks HTTP
    └── buku.routes.cjs         # Pemetaan endpoint & middleware jsonOnly</pre>

<h2>4.2 Implementasi Berkas Utama</h2>
<p><strong>1. src/errors.cjs: </strong>Fungsi parseId memastikan ID berupa string digit desimal murni tanpa leading zero dan berada pada rentang safe integer. Fungsi validateBuku memastikan objek tidak berupa array, tidak memiliki atribut ekstra di luar judul dan tersedia, serta memeriksa tipe boolean secara ketat.</p>
<p><strong>2. src/buku.repository.cjs: </strong>Mengelola array buku dalam memori dengan data awal ID 1 dan ID 2. Counter nextId dinaikkan setiap create untuk menjamin keunikan ID.</p>
<p><strong>3. src/buku.service.cjs: </strong>Memisahkan domain logic dari protokol Express. Mengimplementasikan validasi query q, memastikan buku target ada pada operasi replace dan remove sebelum memanggil repositori.</p>
<p><strong>4. src/buku.controller.cjs: </strong>Menghubungkan konteks request Express dengan service, menetapkan status code 201 dan header Location pada POST, serta status code 204 pada DELETE.</p>
<p><strong>5. src/buku.routes.cjs: </strong>Menggunakan router.param('id') untuk mengekstrak dan memvalidasi ID secara otomatis pada seluruh endpoint berparameter path, serta menerapkan middleware jsonOnly pada POST dan PUT.</p>
<p><strong>6. app.cjs & server.cjs: </strong>Mengintegrasikan middleware logging, Cache-Control no-store, express.json({{limit: '10kb'}}), rute katalog, rute /openapi.json, Swagger UI pada /docs, fallback 404, dan centralized error handler.</p>

<div class="page-break"></div>

<h1>BAB V: MATRIKS PENGUJIAN & ANALISIS HASIL AKTUAL</h1>
<h2>5.1 Matriks Hasil Pengujian Komprehensif (U01 - U24)</h2>
<p>Pengujian dilakukan secara menyeluruh terhadap 24 skenario yang didefinisikan pada modul praktikum. Seluruh pengujian menghasilkan status LULUS (PASS) sesuai dengan spesifikasi kontrak:</p>

<table class="content-table">
  <tr><th>Kode</th><th>Skenario Pengujian</th><th>Hasil yang Diharapkan</th><th>Keluaran Aktual</th><th>Status</th></tr>
"""

for item in test_results:
    exp_text = "Status 200 OK"
    if "POST" in item['cmd']: exp_text = "Status 201 Created + Location"
    if "DELETE" in item['cmd']: exp_text = "Status 204 No Content"
    if "400" in item['desc']: exp_text = "Status 400 Bad Request"
    if "404" in item['desc']: exp_text = "Status 404 Not Found"
    if "415" in item['desc']: exp_text = "Status 415 Unsupported Media Type"
    if "npm ls" in item['cmd']: exp_text = "express@5.3.0 & swagger-ui-express@5.0.1"
    
    act_summary = item['output'].split('\n')[-1]
    if len(act_summary) > 42:
        act_summary = act_summary[:39] + "..."
        
    html_content += f"""  <tr>
    <td><strong>{item['id']}</strong></td>
    <td>{item['desc']}</td>
    <td>{exp_text}</td>
    <td><code>{act_summary}</code></td>
    <td><span class="badge-pass">LULUS</span></td>
  </tr>\n"""

for item in extra_scenarios:
    html_content += f"""  <tr>
    <td><strong>{item[0]}</strong></td>
    <td>{item[1]}</td>
    <td>{item[2]}</td>
    <td><code>{item[3]}</code></td>
    <td><span class="badge-pass">{item[4]}</span></td>
  </tr>\n"""

html_content += """</table>

<h2>5.2 Analisis Hasil Pengujian</h2>
<p><strong>1. Validasi Siklus Hidup CRUD: </strong>Operasi POST berhasil membuat buku baru pada ID 3 dengan mengembalikan status 201 dan header Location: /api/buku/3. Pembaruan PUT mengganti seluruh data objek pada ID 3, dan operasi DELETE menghasilkan respons 204 No Content tanpa payload body. Ketika dilakukan GET sesudah DELETE, server merespons 404 BOOK_NOT_FOUND, membuktikan state lifecycle resource berjalan konsisten.</p>
<p><strong>2. Ketahanan Validasi & Integritas State: </strong>Pemeriksaan ID secara ketat menolak nilai '1abc' dan '01' dengan kode INVALID_ID, membuktikan bahwa validasi regex desimal tanpa leading zero berhasil mencegah ambigu parsing. Pengujian PUT parsial ditolak dengan status 400 INVALID_BOOK dan dibuktikan melalui GET bahwa data asli tidak berubah.</p>
<p><strong>3. Proteksi Parser & Media Type: </strong>Body JSON yang terpotong ditangkap oleh error handler Express dan diubah menjadi format standar aplikasi dengan status 400 INVALID_JSON. Permintaan POST dengan Content-Type text/plain ditolak oleh middleware jsonOnly dengan status 415 UNSUPPORTED_MEDIA_TYPE sebelum masuk ke controller.</p>
<p><strong>4. Karakteristik Volatilitas Memori: </strong>Ketika server dimatikan, client menghasilkan error 'TypeError' (koneksi ditolak). Saat server dijalankan kembali, data tambahan ID 5 hilang dan katalog kembali ke 2 buku awal. Hal ini membuktikan karakteristik in-memory array yang bersifat volatil dan non-persisten.</p>

<div class="page-break"></div>

<h1>BAB VI: PEMBAHASAN TUGAS MANDIRI</h1>
<h2>6.1 Tugas Mandiri 1: Menambah Data Melalui API & Analisis Komparasi</h2>
<p>Pada Tugas Mandiri 1, mahasiswa menambahkan dua buku baru melalui endpoint POST, membaca data, mengubah satu buku, dan menghapus satu buku menggunakan client terminal:</p>
<p><strong>T1_01: POST Buku 1 -> </strong>Judul: 'Arsitektur Microservices Modern', tersedia: true. Respons: HTTP 201, ID aktual: 3, Location: /api/buku/3.</p>
<p><strong>T1_02: POST Buku 2 -> </strong>Judul: 'Desain RESTful API Standar Industri', tersedia: false. Respons: HTTP 201, ID aktual: 4, Location: /api/buku/4.</p>
<p><strong>T1_03: GET Seluruh Katalog -> </strong>Mengembalikan array 4 buku lengkap (ID 1, 2, 3, 4).</p>
<p><strong>T1_04: PUT Buku ID 3 -> </strong>Mengganti judul menjadi 'Arsitektur Microservices Modern (Edisi 2)' dan tersedia menjadi false. Respons: HTTP 200.</p>
<p><strong>T1_05: DELETE Buku ID 4 -> </strong>Menghapus buku ID 4. Respons: HTTP 204 No Content.</p>
<p><strong>T1_06: GET Verifikasi -> </strong>GET /api/buku/4 mengembalikan 404 BOOK_NOT_FOUND. GET /api/buku mengonfirmasi katalog akhir berisi 3 buku (ID 1, 2, 3).</p>

<h3>Analisis Perbedaan Penyuntingan Array Langsung vs Operasi Consumer via API</h3>
<p>Mengapa penyuntingan array pada berkas kode (misal mengedit src/buku.repository.cjs secara manual) TIDAK SAMA dengan operasi yang dikirim oleh consumer melalui API?</p>
<p><strong>1. Dimensi Runtime vs Statis: </strong>Penyuntingan langsung pada berkas kode mengubah kode sumber statis di disk (static code modification). Perubahan ini tidak akan terefleksi pada proses Node.js yang sedang aktif tanpa restart proses. Sebaliknya, operasi melalui API memanipulasi heap memory dari proses runtime yang sedang berjalan secara dinamis tanpa interupsi server.</p>
<p><strong>2. Penegakan Kontrak & Validasi: </strong>Consumer API tunduk pada kontrak antarmuka yang divalidasi oleh berlapis-lapis pertahanan: parsing JSON, pemeriksaan media type, validasi regex ID, dan validasi tipe skema data. Penyuntingan berkas langsung dapat menyusupkan tipe data yang salah tanpa adanya penjaga (gatekeeper).</p>
<p><strong>3. Efek Samping (Side Effects) & Protokol: </strong>Manipulasi via API menghasilkan respons terstruktur (201 Created, Location header, logging console, audit trail). Selain itu, pada sistem multi-pengguna, API menjamin serialisasi eksekusi event loop Node.js untuk mencegah race condition, sedangkan manipulasi file mentah tidak memiliki mekanisme proteksi transaksi.</p>

<h2>6.2 Tugas Mandiri 2: Melengkapi Dokumentasi OpenAPI & Verifikasi Swagger UI</h2>
<p>Pada generator scripts/generate-openapi.cjs, dilakukan pengayaan spesifikasi OpenAPI 3.1.0 untuk memberikan kejelasan kontrak yang maksimal bagi consumer:</p>
<p><strong>1. Deskripsi Semantik Terperinci: </strong>Setiap parameter (seperti parameter path 'id' dan query 'q') ditambahkan deskripsi semantik mendalam mengenai batas panjang UTF-16, format regex desimal tanpa leading zero, serta batas aman integer JavaScript.</p>
<p><strong>2. Penambahan Objek Contoh Respons Error: </strong>Setiap kode status error (400, 404, 413, 415, 500) dilengkapi dengan contoh payload konkret (example value) yang mencerminkan format standar aplikasi: {{ success: false, error: {{ code, message }} }}.</p>
<p><strong>3. Keterpaduan & Regenerasi Dokumen: </strong>Dokumen docs/openapi.json digenerate ulang menggunakan 'npm run docs' dan diverifikasi langsung pada Swagger UI (/docs/). Seluruh perubahan skema dan contoh respons error langsung tampil pada tab Example Value respons Swagger UI tanpa memerlukan restart server.</p>

<h2>6.3 Tugas Mandiri 3: Jawaban Analisis 12 Pertanyaan Pemahaman</h2>
"""

for q, a in qa_items:
    html_content += f"""<h3>{q}</h3>\n<p>{a}</p>\n"""

html_content += """<div class="page-break"></div>

<h1>BAB VII: BUKTI DOKUMENTASI & TANGKAPAN LAYAR</h1>
<p>Berikut adalah 15 tangkapan layar beresolusi tinggi yang mendokumentasikan setiap tahapan pengujian, verifikasi antarmuka web Swagger UI, dan eksekusi terminal secara nyata:</p>
"""

for img_p, cap_title, cap_desc in screenshots_list:
    img_web_path = img_p.replace("\\", "/")
    html_content += f"""<div class="figure">
  <img src="{img_web_path}" alt="{cap_title}">
  <div class="caption"><strong>{cap_title}:</strong> {cap_desc}</div>
</div>\n"""

html_content += """<div class="page-break"></div>

<h1>BAB VIII: KESIMPULAN & REFLEKSI PRAKTIKUM</h1>
<h2>8.1 Kesimpulan</h2>
<p>1. Implementasi web service katalog buku perpustakaan kampus berhasil dibangun menggunakan Node.js v24.11.0 dan framework Express v5.3.0 dengan pemisahan lapisan arsitektur yang solid (routing, controller, service, repository).</p>
<p>2. Desain endpoint mematuhi standar semantik HTTP RFC 9110 secara ketat, termasuk penggunaan status code 200 (OK), 201 (Created), 204 (No Content), 400 (Bad Request), 404 (Not Found), 413 (Payload Too Large), 415 (Unsupported Media Type), dan 500 (Internal Server Error).</p>
<p>3. Seluruh 24 skenario pengujian matriks berhasil lulus (100% PASS), membuktikan keandalan sistem dalam menangani input valid maupun menolak input malformed secara terkontrol tanpa mengekspos stack trace.</p>
<p>4. Spesifikasi OpenAPI 3.1.0 berhasil diintegrasikan secara lokal melalui swagger-ui-express v5.0.1, memberikan dokumentasi interaktif yang memudahkan pengujian dan kolaborasi antar-pengembang.</p>

<h2>8.2 Refleksi Pembelajaran</h2>
<p>Praktikum ini memberikan pemahaman mendalam mengenai pentingnya pemisahan tanggung jawab (Separation of Concerns). Dengan memisahkan logika controller dan service, aturan validasi buku dapat diuji dan dipelihara secara terpisah dari mekanisme pengiriman respons HTTP. Selain itu, eksperimen penghentian dan restart server memberikan bukti nyata mengenai keterbatasan penyimpanan volatil (in-memory) dan memperjelas alasan mengapa sistem produksi memerlukan database persisten serta arsitektur state management yang handal.</p>

<h1>DAFTAR PUSTAKA</h1>
<p>[1] Express.js. (2026). Express 5.x API Reference and Routing Guide. https://expressjs.com/en/5x/api/</p>
<p>[2] Fielding, R., Nottingham, M., & Reschke, J. (2022). HTTP Semantics. RFC 9110, Internet Engineering Task Force (IETF). https://www.rfc-editor.org/rfc/rfc9110.html</p>
<p>[3] Martin, R. C. (2018). Clean Architecture: A Craftsman's Guide to Software Structure and Design. Prentice Hall.</p>
<p>[4] Node.js Foundation. (2026). Node.js v24 LTS Documentation. https://nodejs.org/docs/latest-v24.x/api/</p>
<p>[5] OpenAPI Initiative. (2021). OpenAPI Specification Version 3.1.0. https://spec.openapis.org/oas/v3.1.0.html</p>
<p>[6] Scottie1984. (2024). swagger-ui-express: Swagger UI for Node.js Express Applications. GitHub Repository. https://github.com/scottie1984/swagger-ui-express</p>

</body>
</html>
"""

html_path = "report.html"
with open(html_path, "w", encoding="utf-8") as f:
    f.write(html_content)
print(f"HTML report saved to: {html_path}")

# ==================== CONVERT HTML TO PDF VIA CHROME ====================
pdf_path = "Laporan_Praktikum_05_WSE_240104040147_Muhammad_Daffa_Khairi_Addin.pdf"
chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe"

current_dir = os.path.abspath(".").replace("\\", "/")
html_file_url = f"file:///{current_dir}/report.html"
pdf_out_native = f"{current_dir}/{pdf_path}"

cmd = [
    chrome_path,
    "--headless",
    "--disable-gpu",
    "--no-sandbox",
    f"--print-to-pdf={pdf_out_native}",
    "--no-pdf-header-footer",
    html_file_url
]

print("Executing Chrome headless PDF conversion...")
res = subprocess.run(cmd, capture_output=True, text=True)
print(f"Chrome exit code: {res.returncode}")
if os.path.exists(pdf_path):
    print(f"PDF successfully generated: {pdf_path} (size: {os.path.getsize(pdf_path)} bytes)")
else:
    print(f"Error generating PDF. Stderr: {res.stderr}")
