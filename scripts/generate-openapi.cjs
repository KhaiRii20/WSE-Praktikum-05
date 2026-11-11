const fs = require('node:fs');
const path = require('node:path');

const ref = (name) => ({ $ref: `#/components/schemas/${name}` });

const fields = {
  judul: { type: 'string', description: 'Setelah trim, panjang 1-120 unit UTF-16.' },
  tersedia: { type: 'boolean', description: 'Status ketersediaan buku di katalog (true atau false).' }
};

const book = { id: 1, judul: 'Dasar Pemrograman Kotlin', tersedia: true };
const input = { judul: 'Belajar API dengan Express', tersedia: true };

function jsonResponse(description, schema, example) {
  return {
    description,
    content: {
      'application/json': {
        schema: ref(schema),
        ...(example ? { example } : {})
      }
    }
  };
}

const errorDescriptions = {
  400: 'Permintaan tidak valid, query salah, ID tidak valid, atau body JSON melanggar aturan kontrak.',
  404: 'Buku atau rute yang diminta tidak ditemukan di server.',
  413: 'Body JSON yang dikirim melebihi batas parser 10kb.',
  415: 'Format media tidak didukung. Wajib gunakan Content-Type application/json.',
  500: 'Terjadi kegagalan internal pada server.'
};

const errorExamples = {
  400: { success: false, error: { code: 'INVALID_BOOK', message: 'Isi hanya judul 1-120 karakter dan tersedia boolean.' } },
  404: { success: false, error: { code: 'BOOK_NOT_FOUND', message: 'Buku tidak ditemukan.' } },
  413: { success: false, error: { code: 'PAYLOAD_TOO_LARGE', message: 'Body melebihi batas 10kb.' } },
  415: { success: false, error: { code: 'UNSUPPORTED_MEDIA_TYPE', message: 'Gunakan Content-Type application/json.' } },
  500: { success: false, error: { code: 'INTERNAL_ERROR', message: 'Terjadi kesalahan pada server.' } }
};

function responses(success, codes) {
  const result = { ...success };
  for (const code of codes) {
    result[code] = jsonResponse(errorDescriptions[code], 'ErrorResponse', errorExamples[code]);
  }
  return result;
}

const body = {
  required: true,
  description: 'Data buku baru atau pengganti (judul 1-120 karakter dan tersedia boolean)',
  content: {
    'application/json': {
      schema: ref('BukuInput'),
      example: input
    }
  }
};

const id = {
  name: 'id',
  in: 'path',
  required: true,
  description: 'Digit desimal positif tanpa nol di depan (contoh: 1, 2, 3); nilai <= 9007199254740991.',
  schema: { type: 'string', pattern: '^[1-9][0-9]*$' },
  example: '1'
};

const created = jsonResponse('Buku berhasil dibuat.', 'BukuResponse', {
  success: true,
  data: { id: 3, ...input }
});
created.headers = {
  Location: {
    description: 'URI resource buku baru yang berhasil dibuat.',
    schema: { type: 'string' },
    example: '/api/buku/3'
  }
};

const spec = {
  openapi: '3.1.0',
  info: {
    title: 'API Katalog Perpustakaan',
    version: '1.0.0',
    description:
      'Praktikum lokal tanpa login. Data memori hilang saat restart. ' +
      'POST dan PUT wajib judul serta tersedia, tanpa field tambahan. ' +
      'Kode error: INVALID_ID, INVALID_QUERY, INVALID_BOOK, INVALID_JSON, ' +
      'BOOK_NOT_FOUND, PAYLOAD_TOO_LARGE, UNSUPPORTED_MEDIA_TYPE, ' +
      'ROUTE_NOT_FOUND, BAD_REQUEST, INTERNAL_ERROR.'
  },
  servers: [{ url: 'http://127.0.0.1:3200', description: 'Server Lokal Pengembangan' }],
  tags: [{ name: 'Buku', description: 'Operasi katalog buku perpustakaan kampus.' }],
  paths: {
    '/api/buku': {
      get: {
        tags: ['Buku'],
        summary: 'Membaca atau mencari daftar buku',
        description: 'Mengambil seluruh daftar buku dalam katalog atau mencari buku berdasarkan substring judul.',
        parameters: [
          {
            name: 'q',
            in: 'query',
            description: 'Satu teks pencarian judul, maksimal 80 unit UTF-16; kosong menampilkan semua buku.',
            schema: { type: 'string' },
            example: 'kotlin'
          }
        ],
        responses: responses(
          {
            200: jsonResponse('Daftar buku berhasil diambil; array boleh kosong jika tidak ada hasil.', 'DaftarResponse', {
              success: true,
              data: [book]
            })
          },
          [400, 500]
        )
      },
      post: {
        tags: ['Buku'],
        summary: 'Menambahkan buku baru',
        description: 'Menambahkan entri buku baru ke katalog dengan ID yang digenerate oleh server.',
        requestBody: body,
        responses: responses({ 201: created }, [400, 413, 415, 500])
      }
    },
    '/api/buku/{id}': {
      parameters: [id],
      get: {
        tags: ['Buku'],
        summary: 'Membaca satu buku berdasarkan ID',
        description: 'Mengambil detail data buku berdasarkan parameter path ID positif.',
        responses: responses(
          {
            200: jsonResponse('Buku ditemukan dan data dikembalikan.', 'BukuResponse', {
              success: true,
              data: book
            })
          },
          [400, 404, 500]
        )
      },
      put: {
        tags: ['Buku'],
        summary: 'Mengganti seluruh field buku',
        description: 'Mengganti seluruh data buku yang ada pada ID target. Target harus sudah ada. Field judul dan tersedia wajib.',
        requestBody: body,
        responses: responses(
          {
            200: jsonResponse('Buku berhasil diperbarui.', 'BukuResponse', {
              success: true,
              data: { id: 1, judul: 'Dasar Pemrograman Kotlin (Edisi Revisi)', tersedia: false }
            })
          },
          [400, 404, 413, 415, 500]
        )
      },
      delete: {
        tags: ['Buku'],
        summary: 'Menghapus buku berdasarkan ID',
        description: 'Menghapus buku dari katalog. Menghasilkan respons 204 No Content tanpa body jika berhasil.',
        responses: responses(
          {
            204: { description: 'Buku berhasil dihapus; tidak memiliki body respons.' }
          },
          [400, 404, 500]
        )
      }
    }
  },
  components: {
    schemas: {
      BukuInput: {
        type: 'object',
        additionalProperties: false,
        required: ['judul', 'tersedia'],
        properties: fields
      },
      Buku: {
        type: 'object',
        additionalProperties: false,
        required: ['id', 'judul', 'tersedia'],
        properties: {
          id: {
            type: 'integer',
            minimum: 1,
            maximum: 9007199254740991,
            description: 'ID buku integer positif hasil generate server.'
          },
          ...fields
        }
      },
      BukuResponse: {
        type: 'object',
        required: ['success', 'data'],
        properties: {
          success: { const: true, description: 'Menandakan operasi berhasil.' },
          data: ref('Buku')
        }
      },
      DaftarResponse: {
        type: 'object',
        required: ['success', 'data'],
        properties: {
          success: { const: true, description: 'Menandakan operasi berhasil.' },
          data: { type: 'array', items: ref('Buku') }
        }
      },
      ErrorResponse: {
        type: 'object',
        required: ['success', 'error'],
        properties: {
          success: { const: false, description: 'Menandakan operasi gagal.' },
          error: {
            type: 'object',
            required: ['code', 'message'],
            properties: {
              code: { type: 'string', description: 'Kode kesalahan standar aplikasi.' },
              message: { type: 'string', description: 'Pesan kesalahan ramah pengguna.' }
            }
          }
        }
      }
    }
  }
};

const target = path.join(__dirname, '..', 'docs', 'openapi.json');
fs.mkdirSync(path.dirname(target), { recursive: true });
fs.writeFileSync(target, JSON.stringify(spec, null, 2) + '\n');
console.log('Dokumen dibuat: docs/openapi.json');
