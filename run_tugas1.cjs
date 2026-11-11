const { execSync } = require('child_process');
const fs = require('fs');

const ops = [
  { step: 'T1_01', cmd: 'node client.cjs POST /api/buku samples/tugas1-buku1.json', desc: 'POST Buku 1 (Arsitektur Microservices Modern)' },
  { step: 'T1_02', cmd: 'node client.cjs POST /api/buku samples/tugas1-buku2.json', desc: 'POST Buku 2 (Desain RESTful API Standar Industri)' },
  { step: 'T1_03', cmd: 'node client.cjs GET /api/buku', desc: 'GET seluruh katalog sesudah penambahan 2 buku' },
  { step: 'T1_04', cmd: 'node client.cjs GET /api/buku/3', desc: 'GET detail Buku ID 3' },
  { step: 'T1_05', cmd: 'node client.cjs GET /api/buku/4', desc: 'GET detail Buku ID 4' },
  { step: 'T1_06', cmd: 'node client.cjs PUT /api/buku/3 samples/tugas1-buku-ubah.json', desc: 'PUT Buku ID 3 mengubah judul dan status ketersediaan' },
  { step: 'T1_07', cmd: 'node client.cjs GET /api/buku/3', desc: 'GET detail Buku ID 3 setelah pembaruan' },
  { step: 'T1_08', cmd: 'node client.cjs DELETE /api/buku/4', desc: 'DELETE Buku ID 4' },
  { step: 'T1_09', cmd: 'node client.cjs GET /api/buku/4', desc: 'GET detail Buku ID 4 setelah penghapusan (harus 404)' },
  { step: 'T1_10', cmd: 'node client.cjs GET /api/buku', desc: 'GET seluruh katalog akhir setelah penghapusan' }
];

const results = [];
for (const op of ops) {
  let output = '';
  try {
    output = execSync(op.cmd, { encoding: 'utf8' });
  } catch (err) {
    output = (err.stdout || '') + (err.stderr || '');
  }
  console.log(`=== [${op.step}] ${op.desc} ===`);
  console.log(`Command: ${op.cmd}`);
  console.log(output.trim());
  console.log('--------------------------------------------------\n');
  results.push({ step: op.step, desc: op.desc, cmd: op.cmd, output: output.trim() });
}

fs.writeFileSync('tugas1_results.json', JSON.stringify(results, null, 2));
console.log('Hasil Tugas 1 disimpan di tugas1_results.json');
