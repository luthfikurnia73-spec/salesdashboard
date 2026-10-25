/**
 * Sales Dashboard and Tracking Progress
 * Developed By Luthfi Kurnia, SKM
 * Backend Google Apps Script. Data permanen di Google Sheets (Instansi, Log, Todo, Agenda, Pengguna); foto di Google Drive.
 * Semua akses data lewat fungsi rpc() dan wajib login (kata sandi di-hash, sesi bertanda tangan).
 */
const APP_NAME = 'Sales Dashboard and Tracking Progress';
const SS_ID = ''; // Kosongkan jika script dibuat dari Extensions > Apps Script di dalam spreadsheet. Isi ID spreadsheet jika script berdiri sendiri.
const TZ = 'Asia/Makassar'; // WITA (Balikpapan)
const ONESIGNAL_APP_ID = ''; // dari dashboard OneSignal (Settings > Keys & IDs)
const ONESIGNAL_REST_KEY = ''; // REST API Key OneSignal. RAHASIA, jangan dibagikan
const URL_APLIKASI = ''; // alamat hosting aplikasi (dibuka saat notifikasi diketuk), berakhiran /
const JAM_PENGINGAT = 8; // Jam ringkasan pagi (WITA). Pengingat berulang berjalan sampai jam 17.
const ULANG_JAM = 2; // Pengingat berulang tiap N jam selama masih ada pekerjaan tersisa
const SAPA_JENIS = ['Perusahaan', 'Asuransi']; // Jenis instansi yang wajib disapa tiap hari (tambahkan 'Agensi' bila perlu)
const LIBUR_AKHIR_PEKAN = true; // true = tidak kirim pengingat Sabtu dan Minggu
const SANDI_AWAL = 'hermina123'; // Kata sandi awal pengguna bawaan. Wajib diganti saat masuk pertama kali.
const MASA_SESI_HARI = 30; // Lama sesi login
const NAMA_FOLDER_FOTO = 'Foto Kunjungan - Sales Dashboard';
const AGENDA_AWAL = ['Sapa Harian', 'Perkenalan', 'Penawaran', 'Health Talk', 'Hospital Tour', 'MoU / Kerja Sama', 'Lainnya'];
const PENGGUNA_AWAL = [['luthfi', 'Luthfi Kurnia, SKM', 'Admin'], ['cahya', 'Cahyanindyah Puspitasari, SKM', 'Sales']];

const H = {
  Instansi: ['ID', 'Nama', 'Jenis', 'PIC1', 'Kontak1', 'PIC2', 'Kontak2', 'Alamat', 'Lat', 'Lng', 'Dibuat', 'Diubah', 'DibuatOleh', 'DiubahOleh'],
  Log: ['ID', 'InstansiID', 'Tanggal', 'Jenis', 'Agenda', 'Status', 'Catatan', 'DariID', 'Dibuat', 'Diubah', 'Foto', 'Jam', 'Petugas', 'DibuatOleh', 'DiubahOleh'],
  Agenda: ['Nama'],
  Todo: ['ID', 'Judul', 'Tanggal', 'Jam', 'Prioritas', 'InstansiID', 'Status', 'Catatan', 'Dibuat', 'Diubah', 'Notif', 'Petugas', 'DibuatOleh', 'DiubahOleh'],
  Pengguna: ['Username', 'Nama', 'Peran', 'Aktif', 'Salt', 'Hash', 'WajibGanti', 'Dibuat', 'Diubah']
};

let ME_ = null; // pengguna yang sedang memanggil (diisi oleh rpc)

function doGet() {
  try {
    const out = HtmlService.createHtmlOutputFromFile('index')
      .setTitle(APP_NAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    if (URL_APLIKASI) out.setFaviconUrl(URL_APLIKASI.replace(/\/?$/, '/') + 'icon-192.png');
    return out;
  } catch (e) { // tidak ada file index di Apps Script: backend dipakai sebagai API saja
    return ContentService.createTextOutput(APP_NAME + ' API aktif');
  }
}

/* ===== Spreadsheet ===== */
function ss_() { return SS_ID ? SpreadsheetApp.openById(SS_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }

function tab_(name) {
  const ss = ss_();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, 1, H[name].length).setValues([H[name]]).setFontWeight('bold').setBackground('#2f8f4e').setFontColor('#ffffff');
    s.getRange(2, 1, 1000, H[name].length).setNumberFormat('@'); // format teks agar tanggal, kontak, koordinat tidak berubah
    s.setFrozenRows(1);
    if (name === 'Agenda') s.getRange(2, 1, AGENDA_AWAL.length, 1).setValues(AGENDA_AWAL.map(a => [a]));
  }
  return s;
}

// Memastikan semua kolom yang dibutuhkan ada (otomatis menambah kolom baru di sheet lama). Mengembalikan header sebenarnya.
function hdr_(s, name) {
  const h = s.getLastColumn() ? s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].map(String) : [];
  H[name].forEach(k => {
    if (h.indexOf(k) < 0) {
      if (s.getMaxColumns() < h.length + 1) s.insertColumnAfter(s.getMaxColumns());
      s.getRange(1, h.length + 1).setValue(k).setFontWeight('bold').setBackground('#2f8f4e').setFontColor('#ffffff');
      s.getRange(2, h.length + 1, Math.max(s.getMaxRows() - 1, 1), 1).setNumberFormat('@');
      h.push(k);
    }
  });
  return h;
}

function conv_(h, r) {
  const o = {};
  h.forEach((k, i) => {
    const x = r[i];
    o[k] = x instanceof Date ? Utilities.formatDate(x, TZ, k === 'Jam' ? 'HH:mm' : 'yyyy-MM-dd') : String(x === null || x === undefined ? '' : x);
  });
  return o;
}

function read_(name) {
  const s = tab_(name), h = hdr_(s, name), n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, h.length).getValues().filter(r => r[0] !== '').map(r => conv_(h, r));
}

function row_(s, h, id) {
  const last = s.getLastRow();
  if (last < 2) return 0;
  const ids = s.getRange(2, h.indexOf('ID') + 1, last - 1, 1).getValues().flat().map(String);
  const i = ids.indexOf(String(id));
  return i < 0 ? 0 : i + 2;
}

function upsert_(name, o) {
  const s = tab_(name), h = hdr_(s, name), now = now_(), who = ME_ ? ME_.Username : 'sistem';
  if (!o.ID) o.ID = Utilities.getUuid().slice(0, 8);
  let r = row_(s, h, o.ID);
  o.Diubah = now; o.DiubahOleh = who;
  if (!r) {
    o.Dibuat = now; o.DibuatOleh = who; o.Foto = '';
    r = s.getLastRow() + 1;
    if (r > s.getMaxRows()) s.insertRowsAfter(s.getMaxRows(), 100);
  } else {
    const old = s.getRange(r, 1, 1, h.length).getValues()[0];
    o.Dibuat = old[h.indexOf('Dibuat')] || now;
    o.DibuatOleh = old[h.indexOf('DibuatOleh')] || '';
    if (name === 'Log') o.Foto = old[h.indexOf('Foto')]; // foto hanya diubah lewat uploadFoto / hapusFoto
  }
  if ((name === 'Log' || name === 'Todo') && !o.Petugas && ME_) o.Petugas = ME_.Username;
  s.getRange(r, 1, 1, h.length).setNumberFormat('@').setValues([h.map(k => o[k] === undefined || o[k] === null ? '' : String(o[k]))]);
  return o;
}

function lock_(fn) {
  const L = LockService.getScriptLock();
  L.waitLock(30000);
  try { return fn(); } finally { L.releaseLock(); }
}

function del_(name, col, val) {
  const s = tab_(name), h = hdr_(s, name), v = s.getDataRange().getValues(), c = h.indexOf(col);
  for (let i = v.length - 1; i > 0; i--) if (String(v[i][c]) === String(val)) s.deleteRow(i + 1);
}

/* ===== Login, sesi, dan pengguna ===== */
function hash_(salt, sandi) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + sandi, Utilities.Charset.UTF_8));
}

function secret_() {
  const P = PropertiesService.getScriptProperties();
  let s = P.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); P.setProperty('SECRET', s); }
  return s;
}

function sign_(m) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(m, secret_())); }

function token_(username) {
  const m = username + '|' + (Date.now() + MASA_SESI_HARI * 864e5);
  return Utilities.base64EncodeWebSafe(m) + '.' + sign_(m);
}

function pub_(r) { return { Username: r.Username, Nama: r.Nama, Peran: r.Peran, Aktif: r.Aktif, WajibGanti: r.WajibGanti }; }
function user_(u) { return read_('Pengguna').filter(r => r.Username.toLowerCase() === String(u).toLowerCase())[0]; }

function seedPengguna_() {
  const s = tab_('Pengguna'), h = hdr_(s, 'Pengguna'), now = now_();
  PENGGUNA_AWAL.forEach((u, i) => {
    const salt = Utilities.getUuid(), r = { Username: u[0], Nama: u[1], Peran: u[2], Aktif: 'Ya', Salt: salt, Hash: hash_(salt, SANDI_AWAL), WajibGanti: '1', Dibuat: now, Diubah: now };
    s.getRange(2 + i, 1, 1, h.length).setNumberFormat('@').setValues([h.map(k => String(r[k]))]);
  });
}

function verify_(t) {
  try {
    const p = String(t).split('.'), m = Utilities.newBlob(Utilities.base64DecodeWebSafe(p[0])).getDataAsString();
    if (sign_(m) !== p[1]) throw new Error('x');
    const q = m.split('|');
    if (Date.now() > Number(q[1])) throw new Error('x');
    const r = user_(q[0]);
    if (!r || r.Aktif === 'Tidak') throw new Error('x');
    return r;
  } catch (e) { throw new Error('Sesi berakhir, silakan masuk lagi'); }
}

function login_(username, sandi) {
  const u = String(username || '').trim().toLowerCase(), C = CacheService.getScriptCache(), k = 'gagal:' + u, n = Number(C.get(k) || 0);
  if (n >= 5) throw new Error('Terlalu banyak percobaan. Coba lagi 15 menit lagi.');
  if (!read_('Pengguna').length) seedPengguna_();
  const r = user_(u);
  if (!r || r.Aktif === 'Tidak' || hash_(r.Salt, String(sandi || '')) !== r.Hash) { C.put(k, String(n + 1), 900); throw new Error('Username atau kata sandi salah'); }
  C.remove(k);
  return { token: token_(r.Username), user: pub_(r) };
}

function admin_() { if (!ME_ || ME_.Peran !== 'Admin') throw new Error('Hanya Admin yang boleh melakukan ini'); }

function gantiSandi_(lama, baru) {
  return lock_(() => {
    if (String(baru || '').length < 6) throw new Error('Kata sandi minimal 6 karakter');
    const s = tab_('Pengguna'), h = hdr_(s, 'Pengguna'), v = s.getDataRange().getValues(), i = v.findIndex((r, x) => x > 0 && String(r[0]).toLowerCase() === ME_.Username.toLowerCase());
    if (i < 1) throw new Error('Pengguna tidak ditemukan');
    const cur = conv_(h, v[i]);
    if (hash_(cur.Salt, String(lama || '')) !== cur.Hash) throw new Error('Kata sandi lama salah');
    cur.Salt = Utilities.getUuid(); cur.Hash = hash_(cur.Salt, baru); cur.WajibGanti = ''; cur.Diubah = now_();
    s.getRange(i + 1, 1, 1, h.length).setNumberFormat('@').setValues([h.map(k => String(cur[k]))]);
    return 1;
  });
}

function daftarPengguna_() { admin_(); return read_('Pengguna').map(pub_); }

function simpanPengguna_(o) {
  admin_();
  return lock_(() => {
    const s = tab_('Pengguna'), h = hdr_(s, 'Pengguna'), u = String(o.Username || '').trim().toLowerCase();
    if (!/^[a-z0-9._-]{3,20}$/.test(u)) throw new Error('Username 3-20 karakter: huruf kecil, angka, titik, strip');
    if (u === ME_.Username.toLowerCase() && o.Aktif === 'Tidak') throw new Error('Anda tidak bisa menonaktifkan akun sendiri');
    const n = s.getLastRow(), all = n > 1 ? s.getRange(2, 1, n - 1, h.length).getValues() : [], i = all.findIndex(r => String(r[0]).toLowerCase() === u);
    const cur = i < 0 ? {} : conv_(h, all[i]), now = now_();
    const r = { Username: u, Nama: String(o.Nama || cur.Nama || u).trim(), Peran: o.Peran === 'Admin' ? 'Admin' : 'Sales', Aktif: o.Aktif === 'Tidak' ? 'Tidak' : 'Ya', Salt: cur.Salt || '', Hash: cur.Hash || '', WajibGanti: cur.WajibGanti || '', Dibuat: cur.Dibuat || now, Diubah: now };
    if (o.Password) {
      if (String(o.Password).length < 6) throw new Error('Kata sandi minimal 6 karakter');
      r.Salt = Utilities.getUuid(); r.Hash = hash_(r.Salt, String(o.Password)); r.WajibGanti = '1';
    } else if (i < 0) throw new Error('Kata sandi awal wajib diisi');
    const row = i < 0 ? n + 1 : i + 2;
    if (row > s.getMaxRows()) s.insertRowsAfter(s.getMaxRows(), 20);
    s.getRange(row, 1, 1, h.length).setNumberFormat('@').setValues([h.map(k => r[k] === undefined ? '' : String(r[k]))]);
    return pub_(r);
  });
}

/* ===== Foto kunjungan (Google Drive) ===== */
function folder_() {
  const P = PropertiesService.getScriptProperties(), id = P.getProperty('FOTO_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const it = DriveApp.getFoldersByName(NAMA_FOLDER_FOTO), f = it.hasNext() ? it.next() : DriveApp.createFolder(NAMA_FOLDER_FOTO);
  P.setProperty('FOTO_FOLDER', f.getId());
  return f;
}

function trash_(logs) {
  logs.forEach(l => String(l.Foto || '').split(',').filter(Boolean).forEach(id => { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }));
}

function uploadFoto_(logId, nama, dataUrl) {
  return lock_(() => {
    const s = tab_('Log'), h = hdr_(s, 'Log'), r = row_(s, h, logId);
    if (!r) throw new Error('Catatan tidak ditemukan');
    const c = h.indexOf('Foto') + 1, cur = String(s.getRange(r, c).getValue() || '').split(',').filter(Boolean);
    if (cur.length >= 5) throw new Error('Maksimal 5 foto per catatan');
    const blob = Utilities.newBlob(Utilities.base64Decode(String(dataUrl).split(',')[1]), 'image/jpeg', nama);
    cur.push(folder_().createFile(blob).getId());
    s.getRange(r, c).setNumberFormat('@').setValue(cur.join(','));
    return conv_(h, s.getRange(r, 1, 1, h.length).getValues()[0]);
  });
}

function hapusFoto_(logId, fileId) {
  return lock_(() => {
    const s = tab_('Log'), h = hdr_(s, 'Log'), r = row_(s, h, logId);
    if (!r) throw new Error('Catatan tidak ditemukan');
    const c = h.indexOf('Foto') + 1, cur = String(s.getRange(r, c).getValue() || '').split(',').filter(x => x && x !== fileId);
    s.getRange(r, c).setNumberFormat('@').setValue(cur.join(','));
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
    return conv_(h, s.getRange(r, 1, 1, h.length).getValues()[0]);
  });
}

function getFoto_(fileId) {
  const f = DriveApp.getFileById(fileId), fid = folder_().getId(), ps = f.getParents();
  let ok = false;
  while (ps.hasNext()) if (ps.next().getId() === fid) ok = true;
  if (!ok) throw new Error('Foto tidak valid');
  const b = f.getBlob();
  return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
}

/* ===== Fungsi data (hanya bisa dipanggil lewat rpc setelah login) ===== */
function getAll_() {
  return { orgs: read_('Instansi'), logs: read_('Log'), agenda: read_('Agenda'), todos: read_('Todo'), users: read_('Pengguna').filter(u => u.Aktif !== 'Tidak').map(pub_) };
}
function saveOrg_(o) { return lock_(() => upsert_('Instansi', o)); }
function saveLogs_(arr) { return lock_(() => arr.map(l => upsert_('Log', l))); }
function saveTodo_(o) { return lock_(() => upsert_('Todo', o)); }
function deleteTodo_(id) { return lock_(() => { del_('Todo', 'ID', id); return 1; }); }
function deleteLog_(id) { return lock_(() => { trash_(read_('Log').filter(l => l.ID === id)); del_('Log', 'ID', id); return 1; }); }
function deleteOrg_(id) { return lock_(() => { trash_(read_('Log').filter(l => l.InstansiID === id)); del_('Log', 'InstansiID', id); del_('Instansi', 'ID', id); return 1; }); }

function saveAgenda_(n) {
  return lock_(() => {
    n = String(n).trim();
    if (n && !read_('Agenda').some(a => a.Nama.toLowerCase() === n.toLowerCase())) tab_('Agenda').appendRow([n]);
    return n;
  });
}

function deleteAgenda_(n) {
  return lock_(() => {
    if (n === 'Sapa Harian') throw new Error('Agenda Sapa Harian tidak bisa dihapus');
    del_('Agenda', 'Nama', n);
    return 1;
  });
}

const API = {
  getAll: getAll_, saveOrg: saveOrg_, saveLogs: saveLogs_, deleteLog: deleteLog_, deleteOrg: deleteOrg_, saveAgenda: saveAgenda_, deleteAgenda: deleteAgenda_,
  saveTodo: saveTodo_, deleteTodo: deleteTodo_, uploadFoto: uploadFoto_, hapusFoto: hapusFoto_, getFoto: getFoto_, tesPush: tesPush_,
  gantiSandi: gantiSandi_, daftarPengguna: daftarPengguna_, simpanPengguna: simpanPengguna_
};

// Satu-satunya pintu masuk data. Dipanggil oleh google.script.run (versi Apps Script) dan doPost (versi hosting).
function rpc(fn, args, token) {
  if (fn === 'login') return login_.apply(null, args || []);
  ME_ = verify_(token);
  if (!API[fn]) throw new Error('Fungsi tidak dikenal');
  return API[fn].apply(null, args || []);
}

function doPost(e) {
  let r;
  try {
    const q = JSON.parse(e.postData.contents);
    r = { ok: 1, data: rpc(q.fn, q.args, q.token) };
  } catch (x) {
    r = { ok: 0, err: String(x && x.message ? x.message : x) };
  }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

/* ===== Notifikasi push (OneSignal) dan pengingat otomatis ===== */
function push_(judul, isi, usernames) {
  if (!ONESIGNAL_APP_ID || !ONESIGNAL_REST_KEY) return false;
  const body = { app_id: ONESIGNAL_APP_ID, target_channel: 'push', headings: { en: judul }, contents: { en: isi }, priority: 10, ttl: 6 * 3600 };
  if (usernames && usernames.length) body.include_aliases = { external_id: usernames }; else body.included_segments = ['Subscribed Users'];
  if (URL_APLIKASI) body.url = URL_APLIKASI;
  const r = UrlFetchApp.fetch('https://api.onesignal.com/notifications?c=push', {
    method: 'post', contentType: 'application/json', headers: { Authorization: 'Key ' + ONESIGNAL_REST_KEY },
    payload: JSON.stringify(body), muteHttpExceptions: true
  });
  Logger.log(r.getResponseCode() + ' ' + r.getContentText());
  return r.getResponseCode() < 300;
}

function tesPush_() {
  return push_('🔔 Tes notifikasi', 'Notifikasi Sales Dashboard berfungsi.', ME_ ? [ME_.Username] : null) ? 'Terkirim' : 'Gagal: cek ONESIGNAL_APP_ID dan ONESIGNAL_REST_KEY di Code.gs, atau pastikan perangkat ini sudah mengizinkan notifikasi';
}

// Dijalankan otomatis tiap 15 menit oleh trigger: (1) to-do yang sudah masuk jamnya, (2) ringkasan pagi, (3) pengingat berulang.
function cekPengingat() {
  ME_ = null;
  const now = new Date(), hari = Utilities.formatDate(now, TZ, 'yyyy-MM-dd'), hm = Utilities.formatDate(now, TZ, 'HH:mm'), jam = Number(Utilities.formatDate(now, TZ, 'H'));
  if (LIBUR_AKHIR_PEKAN && Number(Utilities.formatDate(now, TZ, 'u')) >= 6) return;
  const P = PropertiesService.getScriptProperties(), todos = read_('Todo'), logs = read_('Log'), orgs = read_('Instansi');
  todos.filter(t => t.Status !== 'Selesai' && t.Tanggal === hari && t.Jam && t.Jam <= hm && !t.Notif).forEach(t => {
    if (push_('⏰ ' + t.Judul, 'Jadwal ' + t.Jam + (t.Catatan ? ' · ' + t.Catatan : ''), t.Petugas ? [t.Petugas] : null)) lock_(() => { t.Notif = '1'; upsert_('Todo', t); });
  });
  const sapa = orgs.filter(o => SAPA_JENIS.indexOf(o.Jenis) > -1 && !logs.some(l => l.InstansiID === o.ID && l.Tanggal === hari && l.Status === 'Selesai')).length,
    due = logs.filter(l => l.Status === 'Terjadwal' && l.Tanggal <= hari).length,
    td = todos.filter(t => t.Status !== 'Selesai' && t.Tanggal <= hari).length,
    ring = sapa + ' belum disapa · ' + due + ' follow up · ' + td + ' to-do';
  if (jam >= JAM_PENGINGAT && P.getProperty('pagi') !== hari) {
    P.setProperty('pagi', hari);
    if (sapa + due + td > 0) push_('Selamat pagi, rencana hari ini', ring);
  } else if (jam > JAM_PENGINGAT && jam < 17 && sapa + due + td > 0 && Date.now() - Number(P.getProperty('nag') || 0) >= ULANG_JAM * 3600e3) {
    P.setProperty('nag', String(Date.now()));
    push_('🔔 Masih ada yang belum selesai', ring);
  }
}

function aktifkanPengingat() {
  matikanPengingat();
  ScriptApp.newTrigger('cekPengingat').timeBased().everyMinutes(15).create();
}

function matikanPengingat() {
  ScriptApp.getProjectTriggers().filter(t => ['cekPengingat', 'kirimPengingatHarian'].indexOf(t.getHandlerFunction()) > -1).forEach(t => ScriptApp.deleteTrigger(t));
}

// Jalankan SEKALI dari editor Apps Script (pilih fungsi setup lalu Run) dan setujui izin: membuat sheet, folder foto, akun awal, dan pengingat.
function setup() {
  ['Instansi', 'Log', 'Agenda', 'Todo', 'Pengguna'].forEach(n => hdr_(tab_(n), n));
  if (!read_('Pengguna').length) seedPengguna_();
  folder_();
  aktifkanPengingat();
  return 'Setup selesai. Masuk dengan username luthfi atau cahya, kata sandi awal: ' + SANDI_AWAL;
}
