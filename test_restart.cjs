const { execSync } = require('child_process');

console.log("=== Menambahkan buku tambahan (POST ID 5) ===");
let res = execSync('node client.cjs POST /api/buku samples/buku-baru.json', { encoding: 'utf8' });
console.log(res);

console.log("=== Membaca buku ID 5 (harus 200) ===");
res = execSync('node client.cjs GET /api/buku/5', { encoding: 'utf8' });
console.log(res);
