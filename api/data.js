/**
 * SI PINTAR - Backend di Vercel (tanpa Google Apps Script / Spreadsheet)
 * Endpoint: /api/data
 *   GET  ?action=list&type=arsip|anak|config
 *   GET  ?action=track&code=SP-YYMMDD-XXXXXX
 *   GET  ?action=ping
 *   POST (JSON) {type, action, ...}
 * Database: Upstash Redis (dipasang dari Vercel > Storage). Variabel lingkungan dibuat otomatis:
 *   KV_REST_API_URL + KV_REST_API_TOKEN  (atau UPSTASH_REDIS_REST_URL + UPSTASH_REDIS_REST_TOKEN)
 * Login admin (opsional ubah lewat Vercel > Settings > Environment Variables): ADMIN_USER, ADMIN_PASS
 */
const crypto = require('crypto');

const DB_URL = () => process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL || '';
const DB_TOKEN = () => process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN || '';
const ADMIN_USER = () => process.env.ADMIN_USER || 'dinsos';
const ADMIN_PASS = () => process.env.ADMIN_PASS || 'admin123';
const TL_OK = ['Menunggu', 'Diproses', 'Selesai'];
const ANAK_STATUS = ['Dalam Pembinaan', 'Pendampingan', 'Selesai'];
const DEFAULT_WA = '081312403028';
const K = { arsip: 'sipintar:arsip', anak: 'sipintar:anak', config: 'sipintar:config', tiket: 'sipintar:tiket' };
const MAX_ARSIP = 20000;

async function cmd(...args) {
  const r = await fetch(DB_URL(), {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + DB_TOKEN(), 'Content-Type': 'application/json' },
    body: JSON.stringify(args.map(String))
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok || j.error) throw new Error('Database: ' + (j.error || 'HTTP ' + r.status));
  return j.result;
}
async function hgetall(key) {
  const flat = (await cmd('HGETALL', key)) || [];
  const o = {};
  if (Array.isArray(flat)) for (let i = 0; i < flat.length; i += 2) o[flat[i]] = flat[i + 1];
  else Object.assign(o, flat);
  return o;
}
const parse = s => { try { return JSON.parse(s); } catch (_) { return null; } };
const safe = (v, max) => {
  const s = String(v == null ? '' : v).trim().slice(0, max || 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;       // cegah formula injection saat diunduh ke Excel
};
const norm = s => String(s == null ? '' : s).toUpperCase().replace(/[^A-Z0-9]/g, '');
const uuid = () => crypto.randomUUID();
function newTiket() {
  const d = new Date(Date.now() + 7 * 3600e3).toISOString().slice(2, 10).replace(/-/g, '');
  const abc = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = ''; const b = crypto.randomBytes(6); for (let i = 0; i < 6; i++) s += abc[b[i] % abc.length];
  return `SP-${d}-${s}`;
}
function isAdmin(req) {
  const u = String(req.headers['x-admin-user'] || ''), p = String(req.headers['x-admin-pass'] || '');
  const a = Buffer.from(u + '\n' + p), b = Buffer.from(ADMIN_USER() + '\n' + ADMIN_PASS());
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}
const publicArsip = a => ({ id: a.id, waktu: a.waktu, alamat: a.alamat, tindakLanjut: a.tindakLanjut });
async function readBody(req) {
  if (req.body !== undefined && req.body !== null && req.body !== '') {
    return typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
  }
  const chunks = []; for await (const c of req) chunks.push(c);
  const t = Buffer.concat(chunks).toString('utf8');
  return t ? JSON.parse(t) : {};
}
const cleanAnak = x => ({
  ini: safe(x.ini, 20), umur: Math.max(0, Math.min(18, Number(x.umur) || 0)), kec: safe(x.kec, 40), kel: safe(x.kel, 40),
  titik: safe(x.titik, 60), masalah: safe(x.masalah, 300), status: ANAK_STATUS.includes(x.status) ? x.status : 'Dalam Pembinaan'
});

module.exports = async function handler(req, res) {
  const send = (code, obj) => { res.statusCode = code; res.setHeader('Content-Type', 'application/json; charset=utf-8'); res.setHeader('Cache-Control', 'no-store'); res.end(JSON.stringify(obj)); };
  try {
    const q = req.query || Object.fromEntries(new URL(req.url, 'http://x').searchParams);
    if (q.action === 'ping') return send(200, { ok: true, version: 'SI-PINTAR vercel-v1', databaseTerhubung: !!(DB_URL() && DB_TOKEN()) });
    if (!DB_URL() || !DB_TOKEN()) {
      return send(200, { ok: false, error: 'Database belum terhubung. Di Vercel buka Storage > Create > Upstash Redis, hubungkan ke project ini, lalu Redeploy.' });
    }
    const admin = isAdmin(req);

    /* ───────── GET ───────── */
    if (req.method === 'GET') {
      if (q.action === 'track') {
        const want = norm(q.code);
        if (!want) return send(200, { ok: true, data: null });
        const id = await cmd('HGET', K.tiket, want);
        const a = id ? parse(await cmd('HGET', K.arsip, id)) : null;
        return send(200, { ok: true, data: a ? { tiket: a.tiket, waktu: a.waktu, nama: a.nama, alamat: a.alamat, tindakLanjut: a.tindakLanjut } : null });
      }
      if (q.action !== 'list') return send(200, { ok: false, error: 'action tidak dikenal' });
      if (q.type === 'config') {
        const c = await hgetall(K.config);
        return send(200, { ok: true, data: { wa: c.wa || DEFAULT_WA, url: c.url || '' } });
      }
      if (q.type === 'arsip') {
        const all = Object.values(await hgetall(K.arsip)).map(parse).filter(Boolean).sort((a, b) => String(b.waktu).localeCompare(String(a.waktu)));
        return send(200, { ok: true, data: admin ? all : all.map(publicArsip) });
      }
      if (q.type === 'anak') {
        const all = Object.values(await hgetall(K.anak)).map(parse).filter(Boolean);
        return send(200, { ok: true, data: all });
      }
      return send(200, { ok: false, error: 'type tidak dikenal' });
    }

    /* ───────── POST ───────── */
    if (req.method !== 'POST') return send(405, { ok: false, error: 'Metode tidak didukung' });
    const b = await readBody(req), type = b.type, action = b.action;

    if (type === 'arsip' && action === 'create') {                 // publik: form penemuan anak
      if (!String(b.alamat || '').trim() && !String(b.nama || '').trim()) return send(200, { ok: false, error: 'Data laporan kosong' });
      if ((await cmd('HLEN', K.arsip)) >= MAX_ARSIP) return send(200, { ok: false, error: 'Penyimpanan penuh' });
      let tiket = String(b.tiket || '').toUpperCase();
      if (!/^SP-\d{6}-[A-Z0-9]{6}$/.test(tiket)) tiket = newTiket();
      const id = uuid();
      for (let i = 0; i < 5; i++) {                                 // pastikan kode unik
        if (Number(await cmd('HSETNX', K.tiket, norm(tiket), id)) === 1) break;
        tiket = newTiket();
        if (i === 4) return send(200, { ok: false, error: 'Gagal membuat kode tiket, coba lagi' });
      }
      const d = { id, waktu: new Date().toISOString(), nama: safe(b.nama, 120), tempatLahir: safe(b.tempatLahir, 80), ttl: safe(b.ttl, 60),
        alamat: safe(b.alamat, 300), pelapor: safe(b.pelapor, 120), kontak: safe(b.kontak, 60), masalah: safe(b.masalah, 1000),
        tindakLanjut: 'Menunggu', createdBy: 'publik', tiket };
      await cmd('HSET', K.arsip, id, JSON.stringify(d));
      return send(200, { ok: true, id, tiket });
    }

    if (!admin) return send(401, { ok: false, error: 'Akses admin diperlukan. Silakan login ulang.' });

    if (type === 'arsip' && action === 'update') {
      if (!TL_OK.includes(b.tindakLanjut)) return send(200, { ok: false, error: 'Status tidak valid' });
      const a = parse(await cmd('HGET', K.arsip, String(b.id)));
      if (!a) return send(200, { ok: false, error: 'Data tidak ditemukan' });
      a.tindakLanjut = b.tindakLanjut;
      await cmd('HSET', K.arsip, a.id, JSON.stringify(a));
      return send(200, { ok: true });
    }
    if (type === 'anak') {
      if (action === 'create') { const d = { ...cleanAnak(b), id: uuid() }; await cmd('HSET', K.anak, d.id, JSON.stringify(d)); return send(200, { ok: true, id: d.id }); }
      const old = parse(await cmd('HGET', K.anak, String(b.id)));
      if (!old) return send(200, { ok: false, error: 'Data tidak ditemukan' });
      if (action === 'update') { const d = { ...cleanAnak(b), id: old.id }; await cmd('HSET', K.anak, d.id, JSON.stringify(d)); return send(200, { ok: true }); }
      if (action === 'delete') { await cmd('HDEL', K.anak, old.id); return send(200, { ok: true }); }
    }
    if (type === 'config' && action === 'set') {
      const wa = String(b.wa || '').replace(/\D/g, '').replace(/^0/, '62').replace(/^8/, '628');
      if (!/^62\d{8,13}$/.test(wa) && b.wa) return send(200, { ok: false, error: 'Nomor WhatsApp tidak valid' });
      const url = String(b.url || '');
      if (url && !/^https:\/\/\S+$/.test(url)) return send(200, { ok: false, error: 'URL harus https' });
      await cmd('HSET', K.config, 'wa', String(b.wa || ''), 'url', url);
      return send(200, { ok: true });
    }
    return send(200, { ok: false, error: 'action tidak dikenal' });
  } catch (err) {
    return send(200, { ok: false, error: String(err && err.message || err) });
  }
};
