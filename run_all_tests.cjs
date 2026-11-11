const { execSync } = require('child_process');
const fs = require('fs');

const tests = [
  { id: 'U01', cmd: 'npm ls express swagger-ui-express', desc: 'Versi dan paket express 5.3.0 & swagger-ui-express 5.0.1' },
  { id: 'U02', cmd: 'node client.cjs GET /api/buku', desc: 'GET seluruh daftar buku awal' },
  { id: 'U03', cmd: 'node client.cjs GET /api/buku/2', desc: 'GET detail buku ID 2 (tersedia false)' },
  { id: 'U04', cmd: 'node client.cjs GET "/api/buku?q=kotlin"', desc: 'Pencarian buku dengan q=kotlin' },
  { id: 'U05', cmd: 'node client.cjs GET "/api/buku?q=tidakada"', desc: 'Pencarian buku dengan q=tidakada (array kosong)' },
  { id: 'U06', cmd: 'node client.cjs POST /api/buku samples/buku-baru.json', desc: 'POST buku baru valid' },
  { id: 'U07', cmd: 'node client.cjs GET /api/buku/3', desc: 'GET lokasi hasil POST buku baru (ID 3)' },
  { id: 'U08', cmd: 'node client.cjs PUT /api/buku/3 samples/buku-ubah.json', desc: 'PUT penggantian lengkap buku ID 3' },
  { id: 'U09', cmd: 'node client.cjs DELETE /api/buku/3', desc: 'DELETE buku ID 3 (respons 204 No Content)' },
  { id: 'U10', cmd: 'node client.cjs GET /api/buku/3', desc: 'GET sesudah DELETE buku ID 3 (404 BOOK_NOT_FOUND)' },
  { id: 'U11_a', cmd: 'node client.cjs GET /api/buku/1abc', desc: 'GET ID tidak valid 1abc (400 INVALID_ID)' },
  { id: 'U11_b', cmd: 'node client.cjs GET /api/buku/01', desc: 'GET ID dengan leading zero 01 (400 INVALID_ID)' },
  { id: 'U12', cmd: 'node client.cjs GET /api/buku/999', desc: 'GET ID tidak ditemukan 999 (404 BOOK_NOT_FOUND)' },
  { id: 'U13', cmd: 'node client.cjs POST /api/buku samples/buku-salah-tipe.json', desc: 'POST dengan tipe tersedia string "true" (400 INVALID_BOOK)' },
  { id: 'U14_a', cmd: 'node client.cjs PUT /api/buku/1 samples/buku-parsial.json', desc: 'PUT parsial tanpa tersedia (400 INVALID_BOOK)' },
  { id: 'U14_b', cmd: 'node client.cjs GET /api/buku/1', desc: 'GET ID 1 memastikan data tidak berubah setelah PUT gagal' },
  { id: 'U15', cmd: 'node client.cjs POST /api/buku samples/json-rusak.txt', desc: 'POST body JSON rusak (400 INVALID_JSON)' },
  { id: 'U16', cmd: 'node client.cjs POST /api/buku samples/buku-baru.json text/plain', desc: 'POST media type text/plain (415 UNSUPPORTED_MEDIA_TYPE)' },
  { id: 'U17', cmd: 'node client.cjs GET "/api/buku?q=a&q=b"', desc: 'GET query parameter q berulang (400 INVALID_QUERY)' },
  { id: 'U18', cmd: 'node client.cjs GET /api/tidak-ada', desc: 'GET route tidak dikenal (404 ROUTE_NOT_FOUND)' },
  { id: 'U19', cmd: 'node client.cjs GET /openapi.json', desc: 'GET spesifikasi OpenAPI 3.1.0' }
];

const results = [];
for (const t of tests) {
  let output = '';
  let exitCode = 0;
  try {
    output = execSync(t.cmd, { encoding: 'utf8', stdio: ['pipe', 'pipe', 'pipe'] });
  } catch (err) {
    output = (err.stdout || '') + (err.stderr || '');
    exitCode = err.status || 1;
  }
  console.log(`=== [${t.id}] ${t.desc} ===`);
  console.log(`Command: ${t.cmd}`);
  console.log(output.trim());
  console.log('--------------------------------------------------\n');
  results.push({ id: t.id, desc: t.desc, cmd: t.cmd, exitCode, output: output.trim() });
}

fs.writeFileSync('test_results.json', JSON.stringify(results, null, 2));
console.log('Hasil pengujian disimpan di test_results.json');
