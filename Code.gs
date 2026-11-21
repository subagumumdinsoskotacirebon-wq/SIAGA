/**
 * SI PINTAR - Backend Google Apps Script (Code.gs)
 * Cara pakai:
 *  1. Buka Google Sheet yang dipakai SI PINTAR > Ekstensi > Apps Script, tempel seluruh kode ini.
 *  2. Terapkan > Deployment baru > Aplikasi web > Jalankan sebagai: Saya, Akses: Siapa saja.
 *  3. Salin URL /exec ke konstanta APPS_SCRIPT_URL di file HTML (bila URL berubah).
 *  Tab "Arsip", "Anak", dan "Config" dibuat otomatis. Bila tab Anda bernama lain, ubah 'name' di SHEETS.
 *  Kolom dicocokkan lewat nama header; kolom yang belum ada (mis. tiket) ditambahkan otomatis.
 */
const SHEETS = {
  arsip:  { name: 'Arsip',  cols: ['id','waktu','nama','tempatLahir','ttl','alamat','pelapor','kontak','masalah','tindakLanjut','createdBy','tiket'] },
  anak:   { name: 'Anak',   cols: ['id','ini','umur','kec','kel','titik','masalah','status'] },
  config: { name: 'Config', cols: ['key','value'] }
};
const TL_OK = ['Menunggu','Diproses','Selesai'];
const ANAK_STATUS = ['Dalam Pembinaan','Pendampingan','Selesai'];

function out_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function safe_(v, max) {            // cegah formula injection + batasi panjang
  let s = String(v == null ? '' : v).trim().slice(0, max || 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function sheet_(type) {
  const def = SHEETS[type]; if (!def) throw new Error('type tidak dikenal');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(def.name) || ss.insertSheet(def.name);
  const lastCol = Math.max(sh.getLastColumn(), 1);
  let head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  if (head.every(h => h === '')) head = [];
  def.cols.forEach(c => { if (head.indexOf(c) < 0) { head.push(c); } });
  sh.getRange(1, 1, 1, head.length).setValues([head]);
  return { sh, head };
}
function rows_(type) {
  const { sh, head } = sheet_(type), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, head.length).getValues().map(r => {
    const o = {}; head.forEach((h, i) => { o[h] = r[i] instanceof Date ? r[i].toISOString() : r[i]; }); return o;
  }).filter(o => o.id !== '' || type === 'config');
}
function findRow_(sh, head, id) {
  const col = head.indexOf('id') + 1, n = sh.getLastRow();
  if (n < 2) return -1;
  const ids = sh.getRange(2, col, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function put_(sh, head, row, data) {
  head.forEach((h, i) => { if (h in data) sh.getRange(row, i + 1).setValue(data[h]); });
}

function doGet(e) {
  try {
    const p = e.parameter || {};
    if (p.action === 'track') {
      const want = String(p.code || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
      if (!want) return out_({ ok: true, data: null });
      const hit = rows_('arsip').filter(r => String(r.tiket || '').toUpperCase().replace(/[^A-Z0-9]/g, '') === want)[0];
      return out_({ ok: true, data: hit ? { tiket: hit.tiket, waktu: hit.waktu, nama: hit.nama, alamat: hit.alamat, tindakLanjut: hit.tindakLanjut } : null });
    }
    if (p.action !== 'list') return out_({ ok: false, error: 'action tidak dikenal' });
    if (p.type === 'config') {
      const o = {}; rows_('config').forEach(r => { o[r.key] = r.value; });
      return out_({ ok: true, data: o });
    }
    return out_({ ok: true, data: rows_(p.type) });
  } catch (err) { return out_({ ok: false, error: String(err.message || err) }); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const b = JSON.parse(e.postData.contents), type = b.type, action = b.action;

    if (type === 'config' && action === 'set') {
      const wa = String(b.wa || '').replace(/\D/g, '').replace(/^0/, '62').replace(/^8/, '628');
      if (!/^62\d{8,13}$/.test(wa) && b.wa) return out_({ ok: false, error: 'Nomor WhatsApp tidak valid' });
      const url = String(b.url || '');
      if (url && !/^https:\/\/\S+$/.test(url)) return out_({ ok: false, error: 'URL harus https' });
      const { sh, head } = sheet_('config');
      [['wa', String(b.wa || '')], ['url', url]].forEach(([k, v]) => {
        const n = sh.getLastRow(); let row = -1;
        if (n >= 2) { const ks = sh.getRange(2, 1, n - 1, 1).getValues(); for (let i = 0; i < ks.length; i++) if (ks[i][0] === k) row = i + 2; }
        if (row < 0) row = n + 1 > 1 ? n + 1 : 2;
        sh.getRange(row, 1, 1, 2).setValues([[k, v]]);
      });
      return out_({ ok: true });
    }

    if (type === 'arsip') {
      const { sh, head } = sheet_('arsip');
      if (action === 'create') {
        let tiket = String(b.tiket || '').toUpperCase();
        if (!/^SP-\d{6}-[A-Z0-9]{6}$/.test(tiket)) tiket = 'SP-' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyMMdd') + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase();
        const d = { id: Utilities.getUuid(), waktu: new Date().toISOString(), nama: safe_(b.nama, 120), tempatLahir: safe_(b.tempatLahir, 80), ttl: safe_(b.ttl, 60),
          alamat: safe_(b.alamat, 300), pelapor: safe_(b.pelapor, 120), kontak: safe_(b.kontak, 60), masalah: safe_(b.masalah, 1000),
          tindakLanjut: 'Menunggu', createdBy: safe_(b.createdBy || 'publik', 30), tiket: tiket };
        sh.appendRow(head.map(h => (h in d ? d[h] : '')));
        return out_({ ok: true, id: d.id, tiket: tiket });
      }
      if (action === 'update') {
        if (TL_OK.indexOf(b.tindakLanjut) < 0) return out_({ ok: false, error: 'Status tidak valid' });
        const row = findRow_(sh, head, b.id); if (row < 0) return out_({ ok: false, error: 'Data tidak ditemukan' });
        put_(sh, head, row, { tindakLanjut: b.tindakLanjut });
        return out_({ ok: true });
      }
    }

    if (type === 'anak') {
      const { sh, head } = sheet_('anak');
      const clean = x => ({ ini: safe_(x.ini, 20), umur: Math.max(0, Math.min(18, Number(x.umur) || 0)), kec: safe_(x.kec, 40), kel: safe_(x.kel, 40),
        titik: safe_(x.titik, 60), masalah: safe_(x.masalah, 300), status: ANAK_STATUS.indexOf(x.status) >= 0 ? x.status : 'Dalam Pembinaan' });
      if (action === 'create') {
        const d = clean(b); d.id = Utilities.getUuid();
        sh.appendRow(head.map(h => (h in d ? d[h] : '')));
        return out_({ ok: true, id: d.id });
      }
      const row = findRow_(sh, head, b.id); if (row < 0) return out_({ ok: false, error: 'Data tidak ditemukan' });
      if (action === 'update') { put_(sh, head, row, clean(b)); return out_({ ok: true }); }
      if (action === 'delete') { sh.deleteRow(row); return out_({ ok: true }); }
    }
    return out_({ ok: false, error: 'action tidak dikenal' });
  } catch (err) { return out_({ ok: false, error: String(err.message || err) }); }
  finally { try { lock.releaseLock(); } catch (_) {} }
}
