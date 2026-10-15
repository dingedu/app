/**
 * ============================================================
 *  DING! KIDS CLUB — Portal Owner, Admin, Guru & Orang Tua
 *  Backend: Google Apps Script + Google Spreadsheet
 * ============================================================
 *  Cara pasang (detail di PANDUAN.md):
 *  1. Buat Google Spreadsheet baru → Extensions → Apps Script.
 *  2. Tempel file ini ke Code.gs, buat file HTML bernama "index"
 *     lalu tempel isi index.html.
 *  3. Jalankan fungsi setup() sekali (izinkan akses).
 *     → membuat semua sheet + data guru, tarif private, tarif group,
 *       dan paket dari spreadsheet "Student Attendance ERP".
 *  4. Deploy → New deployment → Web app
 *       Execute as: Me | Who has access: Anyone
 *  5. Buka URL /exec → login Owner: owner / ding12345 · Admin: admin / admin12345
 * ============================================================
 */

// Kosongkan jika script dibuat dari dalam spreadsheet (Extensions → Apps Script).
const SS_ID = '';
const TZ = 'Asia/Jakarta';

/* ---------------- ADAPTER Apps Script ---------------- */
const ENV = {
  uuid: () => Utilities.getUuid().replace(/-/g, ''),
  hash: s => Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8).map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join(''),
  today: () => Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'),
  cacheGet: k => CacheService.getScriptCache().get(k),
  cachePut: (k, v, sec) => CacheService.getScriptCache().put(k, v, sec),
  cacheDel: k => CacheService.getScriptCache().remove(k),
  lock: fn => { const l = LockService.getScriptLock(); l.waitLock(20000); try { return fn(); } finally { l.releaseLock(); } }
};

/* ---------------- ENTRY POINTS ---------------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Ding! Kids Club — Portal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function doPost(e) {
  let res;
  try { res = route_(JSON.parse(e.postData.contents)); }
  catch (err) { res = { ok: false, error: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}
function api(reqJson) {
  try { return JSON.stringify(route_(JSON.parse(reqJson))); }
  catch (err) { return JSON.stringify({ ok: false, error: String(err && err.message || err) }); }
}

/* ---------------- SHEET HELPERS ---------------- */
function ss_() { return SS_ID ? SpreadsheetApp.openById(SS_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
const _checked = {};
const _cache = {};
function sheet_(name) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  const h = SCHEMA[name];
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#2C7FDB').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), h.length).setNumberFormat('@');
  } else if (!_checked[name]) {
    const cur = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getDisplayValues()[0].filter(String);
    if (cur.length < h.length) {
      sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#2C7FDB').setFontColor('#FFFFFF');
      sh.getRange(1, cur.length + 1, sh.getMaxRows(), h.length - cur.length).setNumberFormat('@');
    }
  }
  _checked[name] = 1;
  return sh;
}
function readTable_(name) {
  if (_cache[name]) return _cache[name].map(r => Object.assign({}, r));
  const sh = sheet_(name);
  const last = sh.getLastRow();
  const h = SCHEMA[name];
  const rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, h.length).getDisplayValues()
    .filter(r => r[0] !== '')
    .map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); return o; });
  _cache[name] = rows;
  return rows.map(r => Object.assign({}, r));
}
function upsertMany_(name, rows) {
  const sh = sheet_(name);
  const h = SCHEMA[name];
  const existing = readTable_(name);
  const byId = {}; existing.forEach(r => byId[r.id] = r);
  const last = sh.getLastRow();
  const idx = {}; // id → nomor baris asli di sheet (aman walau ada baris kosong)
  if (last >= 2) sh.getRange(2, 1, last - 1, 1).getDisplayValues().forEach((v, i) => { if (v[0] !== '' && !idx[v[0]]) idx[v[0]] = i + 2; });
  const out = [], appends = [];
  rows.forEach(row => {
    const r = Object.assign({}, row);
    if (!r.id || !idx[r.id]) {
      if (!r.id) r.id = newId_(name, existing.concat(out));
      autoFields_(name, r, existing.concat(out));
      appends.push(h.map(k => r[k] == null ? '' : String(r[k])));
    } else {
      const merged = Object.assign({}, byId[r.id], r);
      Object.assign(r, merged);
      sh.getRange(idx[r.id], 1, 1, h.length).setNumberFormat('@').setValues([h.map(k => r[k] == null ? '' : String(r[k]))]);
    }
    const clean = {}; h.forEach(k => clean[k] = r[k] == null ? '' : String(r[k]));
    out.push(clean);
  });
  if (appends.length) {
    const start = Math.max(sh.getLastRow(), 1) + 1;
    sh.getRange(start, 1, appends.length, h.length).setNumberFormat('@').setValues(appends);
  }
  delete _cache[name];
  return out;
}
function removeRow_(name, id) { removeWhere_(name, r => r.id === String(id)); }
function removeWhere_(name, fn) {
  const sh = sheet_(name);
  const rows = readTable_(name);
  const last = sh.getLastRow();
  if (last < 2) return;
  const ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues().map(r => r[0]);
  const kill = {}; rows.forEach(r => { if (fn(r)) kill[r.id] = 1; });
  for (let i = ids.length - 1; i >= 0; i--) if (kill[ids[i]]) sh.deleteRow(i + 2);
  delete _cache[name];
}
function saveSettings_(obj) {
  const sh = sheet_('Pengaturan');
  const cur = getSettings_();
  Object.keys(obj).forEach(k => cur[k] = obj[k]);
  const rows = Object.keys(cur).map(k => [k, String(cur[k])]);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  sh.getRange(2, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
  delete _cache.Pengaturan;
}
function fillTable_(name, prefix, digits, fillers) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const h = SCHEMA[name];
  const rng = sh.getRange(2, 1, last - 1, h.length);
  const vals = rng.getDisplayValues();
  const c = k => h.indexOf(k);
  let max = 0;
  vals.forEach(r => { const m = String(r[0]).match(new RegExp('^' + prefix + '(\\d+)$')); if (m && +m[1] > max) max = +m[1]; });
  let changed = 0;
  vals.forEach(r => {
    if (!String(r[c('nama')]).trim()) return;
    const before = r.join('|');
    if (!r[0]) r[0] = prefix + String(++max).padStart(digits, '0');
    Object.keys(fillers).forEach(k => { if (!r[c(k)]) r[c(k)] = fillers[k](); });
    if (r.join('|') !== before) changed++;
  });
  if (changed) { rng.setNumberFormat('@').setValues(vals); delete _cache[name]; }
  return changed;
}

/* ---------------- SETUP (jalankan sekali) ---------------- */
function setup() {
  Object.keys(SCHEMA).forEach(sheet_);
  const def = ss_().getSheetByName('Sheet1') || ss_().getSheetByName('Lembar1');
  if (def && def.getLastRow() === 0 && ss_().getSheets().length > 1) ss_().deleteSheet(def);
  if (!readTable_('Users').length) {
    upsertMany_('Users', [
      { username: 'owner', passwordHash: ENV.hash('ding12345'), nama: 'Owner Ding', role: 'owner' },
      { username: 'admin', passwordHash: ENV.hash('admin12345'), nama: 'Admin Ding', role: 'admin' }
    ]);
  }
  const seed = seedRows_();
  ['Guru', 'TarifPrivat', 'TarifGroup', 'Paket'].forEach(t => { if (!readTable_(t).length) upsertMany_(t, seed[t]); });
  saveSettings_({});
  lengkapiDataSiswa(); lengkapiDataGuru();
  Logger.log('Setup selesai. Owner: owner / ding12345 · Admin: admin / admin12345 (segera ganti).');
}
function resetPasswordOwner() {
  const u = readTable_('Users').find(r => r.username === 'owner');
  if (u) { u.passwordHash = ENV.hash('ding12345'); upsertMany_('Users', [u]); }
  Logger.log('Password owner direset ke: ding12345');
}

/* ================================================================
 *  DING! KIDS CLUB — CORE (dipakai bersama oleh Code.gs & mode demo)
 *  Berisi: schema sheet, data awal (tarif dari spreadsheet klien),
 *  login & hak akses, router API, dan ENGINE perhitungan:
 *  sisa sesi paket, pendapatan (revenue) per bulan, gaji guru.
 * ================================================================ */

const SESSION_SECONDS = 21600; // 6 jam

// Kolom baru selalu DITAMBAH DI BELAKANG agar sheet lama tetap cocok.
const SCHEMA = {
  Pengaturan:  ['key', 'value'],
  Users:       ['id', 'username', 'passwordHash', 'nama', 'role'],
  Guru:        ['id', 'nama', 'wa', 'status', 'pin', 'catatan'],
  Siswa:       ['id', 'nama', 'panggilan', 'tglLahir', 'sekolah', 'ortu', 'waOrtu', 'alamat', 'tglDaftar', 'status', 'kodeOrtu', 'referralDari', 'referralGuruId', 'catatan'],
  KelasGroup:  ['id', 'nama', 'program', 'hari', 'jam', 'durasi', 'lokasi', 'guruId', 'kapasitas', 'status', 'catatan'],
  Enrol:       ['id', 'siswaId', 'tipe', 'kelasId', 'tarifId', 'guruId', 'kategori', 'lokasi', 'durasi', 'travelFee', 'jadwal', 'tglMulai', 'status', 'catatan'],
  Paket:       ['id', 'nama', 'jenis', 'untuk', 'kategori', 'sesi', 'bonus', 'jatahIzin', 'harga', 'aktif'],
  Pembelian:   ['id', 'noInvoice', 'siswaId', 'enrolId', 'paketId', 'jenis', 'nama', 'sesi', 'bonus', 'jatahIzin', 'harga', 'biayaLain', 'ketBiayaLain', 'diskon', 'total', 'tglBeli', 'status', 'tglBayar', 'metode', 'asalId', 'catatan'],
  Sesi:        ['id', 'tanggal', 'jam', 'tipe', 'kelasId', 'guruId', 'durasi', 'lokasi', 'materi', 'catatan', 'inputOleh'],
  Hadir:       ['id', 'sesiId', 'siswaId', 'enrolId', 'status', 'catatan'],
  TarifPrivat: ['id', 'guruId', 'kategori', 'lokasi', 'versi', 'tglBerlaku', 'nilaiKelas', 'persenGuru', 'adminFee', 'feeGuruManual', 'classroomFee', 'persenReferral', 'travelFee', 'aktif'],
  TarifGroup:  ['id', 'guruId', 'versi', 'tglBerlaku', 'jumlahMurid', 'feeGuru'],
  Gaji:        ['id', 'guruId', 'periode', 'bonus', 'potongan', 'catatan', 'tglBayar', 'total']
};

// tabel yang boleh diubah per peran
const TABLES_OWNER = ['Guru', 'Siswa', 'KelasGroup', 'Enrol', 'Paket', 'Pembelian', 'Sesi', 'Hadir', 'TarifPrivat', 'TarifGroup', 'Gaji'];
const TABLES_ADMIN = ['Guru', 'Siswa', 'KelasGroup', 'Enrol', 'Paket', 'Pembelian', 'Sesi', 'Hadir'];
// kolom tarif yang boleh dilihat admin (admin TIDAK melihat fee guru)
const TARIF_ADMIN_COLS = ['id', 'guruId', 'kategori', 'lokasi', 'versi', 'tglBerlaku', 'nilaiKelas', 'travelFee', 'aktif'];

const DEFAULT_LOKASI = JSON.stringify([
  { kode: 'HS', nama: 'Home Service', rumah: 'Ya' },
  { kode: 'HSW', nama: 'Home Service (West)', rumah: 'Ya' },
  { kode: 'DS', nama: 'Ding Studio', rumah: '' },
  { kode: 'OS', nama: 'Owner Studio', rumah: '' },
  { kode: 'OC', nama: 'Online Class', rumah: '' }
]);
const DEFAULT_KATEGORI = JSON.stringify(['Private', 'Semi-Private', 'Art Private', 'Art Semi-Private', 'School Tutoring']);

const DEFAULT_SETTINGS = {
  namaLembaga: 'Ding! Kids Club',
  tagline: 'Fun learning club for kids',
  alamat: 'Alamat Ding Studio',
  noWa: '081234567890',
  rekening: 'BCA 1234567890 a.n. Ding Kids Club',
  durasiStandar: '45',          // menit; durasi custom → harga & fee guru dihitung proporsional
  adminFeeGroup: '5',           // % admin fee untuk group class
  referralPersen: '7',          // % referral group class (private ikut kolom di Tarif Private)
  referralBulan: '3',           // referral berlaku N bulan sejak murid daftar
  alpaPotong: 'Ya',             // murid Alpa (tanpa kabar) → sesi terpotong
  izinLewatJatah: 'Potong',     // izin melebihi jatah izin paket → Potong / Tidak
  guruDibayarAlpa: 'Ya',        // private: guru tetap dibayar bila murid Alpa (guru sudah datang)
  warnSisa: '1',                // peringatan bila sisa sesi ≤ angka ini
  lokasi: DEFAULT_LOKASI,
  kategori: DEFAULT_KATEGORI,
  portalUrl: ''
};
const PUBLIC_KEYS = ['namaLembaga', 'tagline', 'alamat', 'noWa', 'rekening', 'durasiStandar', 'warnSisa', 'lokasi', 'kategori', 'portalUrl', 'alpaPotong', 'izinLewatJatah'];

/* ---------------- DATA AWAL dari spreadsheet "Student Attendance ERP" ---------------- */
const SEED_GURU = ['Fenny', 'Inez', 'Jessi', 'Nikita', 'Ribka', 'Selo', 'Lisa', 'Fefe', 'Fio', 'Nia', 'Zia', 'Elis', 'Tessa'];
// [guru, kategori, lokasi, versi, tglBerlaku, nilaiKelas, persenGuru, feeGuruManual, classroomFee]
const SEED_TARIF = [
  ['Fenny', 'Private', 'HS', '01', '2026-02-01', 450000, 70, '', 0],
  ['Fenny', 'Private', 'OS', '01', '2026-02-01', 400000, 70, '', 30000],
  ['Fenny', 'Private', 'DS', '01', '2026-02-01', 400000, 70, '', 100000],
  ['Fenny', 'Private', 'OC', '01', '2026-02-01', 350000, 70, '', 15000],
  ['Inez', 'Private', 'HS', '01', '2026-04-01', 250000, 70, '', 0],
  ['Inez', 'Private', 'HS', '02', '2026-08-01', 300000, 70, '', 0],
  ['Jessi', 'Private', 'HS', '01', '2026-02-01', 300000, 70, '', 0],
  ['Jessi', 'Private', 'DS', '01', '2026-02-01', 275000, 55, '', 100000],
  ['Jessi', 'Private', 'OC', '01', '2026-02-01', 200000, 55, '', 15000],
  ['Nikita', 'Private', 'HS', '01', '2026-02-01', 450000, 70, '', 0],
  ['Nikita', 'Private', 'OS', '01', '2026-02-01', 400000, 70, '', 30000],
  ['Nikita', 'Private', 'DS', '01', '2026-02-01', 400000, 70, '', 100000],
  ['Nikita', 'Private', 'OC', '01', '2026-02-01', 350000, 70, '', 15000],
  ['Ribka', 'Private', 'HS', '01', '2026-02-01', 400000, 70, '', 0],
  ['Ribka', 'Private', 'DS', '01', '2026-02-01', 350000, 70.38, 240000, 100000],
  ['Ribka', 'Private', 'DS', '02', '2026-02-01', 375000, 65, '', 100000],
  ['Ribka', 'Private', 'OC', '01', '2026-02-01', 300000, 65, '', 15000],
  ['Selo', 'Private', 'HS', '01', '2026-02-01', 300000, 70, '', 0],
  ['Selo', 'Private', 'DS', '01', '2026-02-01', 275000, 55, '', 100000],
  ['Selo', 'Private', 'OC', '01', '2026-02-01', 200000, 55, '', 15000],
  ['Lisa', 'Private', 'HS', '01', '2026-02-01', 300000, 70, '', 0],
  ['Lisa', 'Private', 'DS', '01', '2026-02-01', 275000, 55, '', 100000],
  ['Lisa', 'Private', 'OC', '01', '2026-02-01', 200000, 55, '', 15000],
  ['Fefe', 'Private', 'HS', '01', '2026-07-01', 300000, 70, '', 0],
  ['Fefe', 'Private', 'HSW', '01', '2026-07-01', 375000, 76, '', 0],
  ['Fefe', 'Private', 'HSW', '02', '2026-08-29', 425000, 76, 316774, 0],
  ['Fio', 'School Tutoring', 'HS', '01', '2026-09-24', 150000, 66.67, '', 0],
  ['Nia', 'School Tutoring', 'HS', '01', '2026-09-24', 150000, 66.67, '', 0],
  ['Zia', 'School Tutoring', 'HS', '01', '2026-09-24', 150000, 66.67, '', 0],
  ['Zia', 'Art Private', 'HS', '01', '2026-09-01', 240000, 62.5, '', 0],
  ['Zia', 'Art Semi-Private', 'HS', '02', '2026-09-01', 220000, 68.18, '', 0],
  ['Zia', 'Private', 'HS', '01', '2026-09-01', 300000, 70, '', 0],
  ['Zia', 'Private', 'DS', '01', '2026-09-01', 275000, 55, '', 100000],
  ['Elis', 'School Tutoring', 'HS', '01', '2026-09-24', 150000, 66.67, '', 0],
  ['Tessa', 'School Tutoring', 'HS', '01', '2026-09-24', 150000, 66.67, '', 0]
];
// [guru, tglBerlaku, fee untuk 1..6 murid hadir]
const SEED_GROUP = [
  ['Fenny', '2026-02-01', [200000, 200000, 200000, 250000, 300000, 300000]],
  ['Jessi', '2026-02-01', [125000, 125000, 125000, 160000, 200000, 200000]],
  ['Nikita', '2026-02-01', [200000, 200000, 200000, 250000, 300000, 300000]],
  ['Ribka', '2026-02-01', [200000, 200000, 200000, 250000, 300000, 300000]],
  ['Selo', '2026-02-01', [125000, 125000, 125000, 160000, 200000, 200000]],
  ['Lisa', '2026-05-01', [125000, 125000, 125000, 160000, 200000, 200000]],
  ['Fefe', '2026-07-01', [125000, 125000, 125000, 160000, 200000, 200000]],
  ['Zia', '2026-10-07', [null, null, null, null, null, 125000]]
];
// [id, nama, jenis, untuk, kategori, sesi, bonus, jatahIzin, harga]  harga 0 = dihitung dari tarif guru
const SEED_PAKET = [
  ['REG', 'Registration Fee', 'Registrasi', 'Semua', '', 0, 0, 0, 200000],
  ['TRIAL', 'Trial Class (Group)', 'Trial', 'Group', '', 1, 0, 0, 100000],
  ['G4', 'Group Package 4', 'Paket', 'Group', '', 4, 0, 0, 800000],
  ['G8', 'Group Package 8', 'Paket', 'Group', '', 8, 1, 1, 1600000],
  ['G12', 'Group Package 12', 'Paket', 'Group', '', 12, 2, 3, 2400000],
  ['P1', 'Private 1 Class', 'Paket', 'Private', '', 1, 0, 0, 0],
  ['P4', 'Private 4 Class', 'Paket', 'Private', '', 4, 0, 0, 0],
  ['P8', 'Private 8 Class', 'Paket', 'Private', '', 8, 0, 0, 0],
  ['ART4', 'Art Private Pack 4', 'Paket', 'Private', 'Art Private', 4, 0, 0, 880000],
  ['ARTS4', 'Art Semi-Private Pack 4', 'Paket', 'Private', 'Art Semi-Private', 4, 0, 0, 800000],
  ['ST4', 'School Tutoring Pack 4', 'Paket', 'Private', 'School Tutoring', 4, 0, 0, 600000]
];

function seedRows_() {
  const guru = SEED_GURU.map((n, i) => ({ id: 'G' + String(i + 1).padStart(2, '0'), nama: n, wa: '', status: 'Aktif', pin: '', catatan: '' }));
  const gid = n => (guru.find(g => g.nama === n) || {}).id;
  const tarif = SEED_TARIF.map((t, i) => ({
    id: 'TP' + String(i + 1).padStart(3, '0'), guruId: gid(t[0]), kategori: t[1], lokasi: t[2], versi: t[3], tglBerlaku: t[4],
    nilaiKelas: t[5], persenGuru: t[6], adminFee: 9000, feeGuruManual: t[7], classroomFee: t[8], persenReferral: 7, travelFee: 0, aktif: 'Ya'
  }));
  const group = [];
  SEED_GROUP.forEach(g => g[2].forEach((fee, i) => { if (fee != null) group.push({ id: 'TG' + String(group.length + 1).padStart(3, '0'), guruId: gid(g[0]), versi: '01', tglBerlaku: g[1], jumlahMurid: i + 1, feeGuru: fee }); }));
  const paket = SEED_PAKET.map(p => ({ id: p[0], nama: p[1], jenis: p[2], untuk: p[3], kategori: p[4], sesi: p[5], bonus: p[6], jatahIzin: p[7], harga: p[8], aktif: 'Ya' }));
  return { Guru: guru, TarifPrivat: tarif, TarifGroup: group, Paket: paket };
}

/* ---------------- ROUTER ---------------- */

function route_(req) {
  const a = req.action;
  if (a === 'ping') return { ok: true, time: new Date().toISOString() };
  if (a === 'login') return login_(req);
  if (a === 'publicSettings') return { ok: true, settings: publicSettings_() };

  const s = getSession_(req.token);
  if (!s) return { ok: false, error: 'SESSION', message: 'Sesi berakhir, silakan login kembali.' };
  if (a === 'logout') { ENV.cacheDel('s_' + req.token); return { ok: true }; }

  if (s.role === 'owner' || s.role === 'admin') {
    const allowed = s.role === 'owner' ? TABLES_OWNER : TABLES_ADMIN;
    const chk = t => { if (allowed.indexOf(t) < 0) throw new Error('Tidak punya akses ke data ' + t); };
    switch (a) {
      case 'bootstrap':      return { ok: true, data: bootstrap_(s) };
      case 'save':           chk(req.table); return { ok: true, row: ENV.lock(() => upsertMany_(req.table, [clean_(req.table, req.row, s)])[0]) };
      case 'saveMany':       chk(req.table); return { ok: true, rows: ENV.lock(() => upsertMany_(req.table, (req.rows || []).map(r => clean_(req.table, r, s)))) };
      case 'remove':         chk(req.table); ENV.lock(() => removeRow_(req.table, req.id)); return { ok: true };
      case 'saveSesi':       return { ok: true, data: ENV.lock(() => saveSesi_(s, req)) };
      case 'delSesi':        ENV.lock(() => delSesi_(s, req.id)); return { ok: true };
      case 'saveSettings':   ENV.lock(() => saveSettings_(req.settings || {})); return { ok: true, settings: getSettings_() };
      case 'changePassword': return changePassword_(s, req);
      case 'resetUser':      if (s.role !== 'owner') break; return resetUser_(req);
    }
  }
  if (s.role === 'guru') {
    switch (a) {
      case 'guruData': return { ok: true, data: guruData_(s) };
      case 'saveSesi': return { ok: true, data: ENV.lock(() => saveSesi_(s, req)) };
      case 'delSesi':  ENV.lock(() => delSesi_(s, req.id)); return { ok: true };
    }
  }
  if (s.role === 'ortu' && a === 'me') return { ok: true, data: me_(s) };
  return { ok: false, error: 'Akses ditolak' };
}

// admin tidak boleh menimpa kolom fee tarif (kalaupun dikirim)
function clean_(table, row, s) {
  const r = Object.assign({}, row || {});
  if (table === 'Guru' && r.pin === undefined) delete r.pin;
  return r;
}

/* ---------------- AUTH ---------------- */

function login_(req) {
  const role = String(req.role || '');
  const user = String(req.user || '').trim();
  const pass = String(req.pass || '').trim();
  if (!user || !pass) return { ok: false, error: 'Lengkapi data login.' };
  let profile = null, realRole = role;
  if (role === 'staff') {
    const u = readTable_('Users').find(r => String(r.username).toLowerCase() === user.toLowerCase());
    if (u && u.passwordHash === ENV.hash(pass)) { profile = { id: u.id, nama: u.nama, username: u.username }; realRole = u.role === 'owner' ? 'owner' : 'admin'; }
  } else if (role === 'guru') {
    lengkapiDataGuru();
    const g = readTable_('Guru').find(r => String(r.id).toUpperCase() === user.toUpperCase());
    if (g && String(g.pin) === pass && g.status !== 'Nonaktif') profile = { id: g.id, nama: g.nama };
  } else if (role === 'ortu') {
    lengkapiDataSiswa();
    const x = readTable_('Siswa').find(r => String(r.id).toUpperCase() === user.toUpperCase());
    if (x && String(x.kodeOrtu) === pass && x.status !== 'Berhenti') profile = { id: x.id, nama: x.nama, panggilan: x.panggilan };
  }
  if (!profile) return { ok: false, error: role === 'staff' ? 'Username atau password salah.' : role === 'guru' ? 'ID guru atau PIN salah.' : 'ID murid atau kode akses salah.' };
  const token = ENV.uuid();
  ENV.cachePut('s_' + token, JSON.stringify({ role: realRole, uid: profile.id }), SESSION_SECONDS);
  return { ok: true, token: token, role: realRole, profile: profile };
}

function getSession_(token) {
  if (!token) return null;
  const v = ENV.cacheGet('s_' + token);
  if (!v) return null;
  ENV.cachePut('s_' + token, v, SESSION_SECONDS);
  return JSON.parse(v);
}

function changePassword_(s, req) {
  const u = readTable_('Users').find(r => r.id === s.uid);
  if (!u || u.passwordHash !== ENV.hash(String(req.oldPass || ''))) return { ok: false, error: 'Password lama salah.' };
  if (String(req.newPass || '').length < 6) return { ok: false, error: 'Password baru minimal 6 karakter.' };
  u.passwordHash = ENV.hash(String(req.newPass));
  ENV.lock(() => upsertMany_('Users', [u]));
  return { ok: true };
}

function resetUser_(req) {
  const uname = String(req.username || '').trim();
  const pass = String(req.newPass || '');
  if (!uname || pass.length < 6) return { ok: false, error: 'Username wajib & password minimal 6 karakter.' };
  const u = readTable_('Users').find(r => String(r.username).toLowerCase() === uname.toLowerCase());
  const row = u ? Object.assign({}, u) : { username: uname, nama: req.nama || uname, role: req.role === 'owner' ? 'owner' : 'admin' };
  row.passwordHash = ENV.hash(pass);
  if (req.nama) row.nama = req.nama;
  ENV.lock(() => upsertMany_('Users', [row]));
  return { ok: true, users: safeUsers_() };
}
function safeUsers_() { return readTable_('Users').map(u => ({ id: u.id, username: u.username, nama: u.nama, role: u.role })); }

/* ---------------- DATA PER PERAN ---------------- */

function bootstrap_(s) {
  lengkapiDataSiswa(); lengkapiDataGuru();
  const db = {};
  const list = s.role === 'owner' ? TABLES_OWNER : TABLES_ADMIN.concat(['TarifPrivat']);
  list.forEach(t => db[t] = readTable_(t));
  if (s.role === 'admin') {
    db.TarifPrivat = db.TarifPrivat.map(r => { const o = {}; TARIF_ADMIN_COLS.forEach(k => o[k] = r[k]); return o; });
    db.TarifGroup = []; db.Gaji = [];
  } else {
    db.Users = safeUsers_();
  }
  return { db: db, settings: getSettings_(), role: s.role };
}

function guruData_(s) {
  const g = readTable_('Guru').find(r => r.id === s.uid);
  if (!g) throw new Error('Data guru tidak ditemukan');
  const db = {};
  ['Siswa', 'KelasGroup', 'Enrol', 'Pembelian', 'Sesi', 'Hadir'].forEach(t => db[t] = readTable_(t));
  const st = getSettings_();
  const L = engine_(db, st);
  const enrolAktif = db.Enrol.filter(e => e.status !== 'Selesai' && e.status !== 'Berhenti');
  const sid = {}; enrolAktif.forEach(e => sid[e.siswaId] = 1);
  const sisa = {};
  enrolAktif.forEach(e => { const x = L.enrol[e.id]; sisa[e.id] = x ? { sisa: x.sisa, kurang: x.kurang, warn: x.warn } : { sisa: 0, kurang: 0, warn: 'habis' }; });
  const cutoff = addDays_(ENV.today(), -120);
  const mySesi = db.Sesi.filter(x => x.guruId === s.uid && x.tanggal >= cutoff);
  const ids = {}; mySesi.forEach(x => ids[x.id] = 1);
  const safe = { id: g.id, nama: g.nama, wa: g.wa };
  return {
    guru: safe,
    settings: publicSettings_(),
    Guru: readTable_('Guru').filter(x => x.status !== 'Nonaktif').map(x => ({ id: x.id, nama: x.nama })),
    Siswa: db.Siswa.filter(x => sid[x.id]).map(x => ({ id: x.id, nama: x.nama, panggilan: x.panggilan, sekolah: x.sekolah })),
    KelasGroup: db.KelasGroup.filter(k => k.status !== 'Nonaktif'),
    Enrol: enrolAktif.map(e => ({ id: e.id, siswaId: e.siswaId, tipe: e.tipe, kelasId: e.kelasId, guruId: e.guruId, kategori: e.kategori, lokasi: e.lokasi, durasi: e.durasi, jadwal: e.jadwal })),
    Sesi: mySesi,
    Hadir: db.Hadir.filter(h => ids[h.sesiId]),
    sisa: sisa
  };
}

function me_(s) {
  const siswa = readTable_('Siswa').find(r => r.id === s.uid);
  if (!siswa) throw new Error('Data murid tidak ditemukan');
  const hadir = readTable_('Hadir').filter(h => h.siswaId === s.uid);
  const sids = {}; hadir.forEach(h => sids[h.sesiId] = 1);
  const safe = Object.assign({}, siswa); delete safe.kodeOrtu; delete safe.referralGuruId;
  return {
    siswa: safe,
    settings: publicSettings_(),
    Guru: readTable_('Guru').map(g => ({ id: g.id, nama: g.nama })),
    KelasGroup: readTable_('KelasGroup').map(k => ({ id: k.id, nama: k.nama, program: k.program, hari: k.hari, jam: k.jam, lokasi: k.lokasi, guruId: k.guruId })),
    Enrol: readTable_('Enrol').filter(e => e.siswaId === s.uid).map(e => { const o = Object.assign({}, e); delete o.tarifId; return o; }),
    Pembelian: readTable_('Pembelian').filter(p => p.siswaId === s.uid && p.status !== 'Batal'),
    Sesi: readTable_('Sesi').filter(x => sids[x.id]).map(x => ({ id: x.id, tanggal: x.tanggal, jam: x.jam, tipe: x.tipe, kelasId: x.kelasId, guruId: x.guruId, durasi: x.durasi, lokasi: x.lokasi, materi: x.materi, catatan: x.catatan })),
    Hadir: hadir
  };
}

/* ---------------- SESI & ABSENSI ---------------- */

const STATUS_HADIR = ['Hadir', 'Izin', 'Sakit', 'Alpa'];

function saveSesi_(s, req) {
  const src = req.sesi || {};
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(src.tanggal || ''))) throw new Error('Tanggal sesi tidak valid.');
  const enrolAll = readTable_('Enrol');
  let old = null;
  if (src.id) old = readTable_('Sesi').find(x => x.id === src.id);
  if (s.role === 'guru') {
    if (old && old.guruId !== s.uid) throw new Error('Sesi ini diinput guru lain.');
    src.guruId = s.uid;
  }
  if (!src.guruId) throw new Error('Guru pengajar wajib dipilih.');
  const sesi = {
    tanggal: src.tanggal, jam: String(src.jam || ''), tipe: src.tipe === 'Group' ? 'Group' : 'Private',
    kelasId: String(src.kelasId || ''), guruId: src.guruId, durasi: String(src.durasi || ''), lokasi: String(src.lokasi || ''),
    materi: String(src.materi || ''), catatan: String(src.catatan || ''), inputOleh: s.role + ':' + s.uid
  };
  if (old) sesi.id = old.id;
  const rows = (req.hadir || []).filter(h => STATUS_HADIR.indexOf(h.status) > -1).map(h => {
    const e = enrolAll.find(x => x.id === h.enrolId);
    if (!e) throw new Error('Pendaftaran kelas murid tidak ditemukan.');
    return { siswaId: e.siswaId, enrolId: e.id, status: h.status, catatan: String(h.catatan || '') };
  });
  if (!rows.length) throw new Error('Isi absensi minimal satu murid.');
  const saved = upsertMany_('Sesi', [sesi])[0];
  removeWhere_('Hadir', h => h.sesiId === saved.id);
  const hs = upsertMany_('Hadir', rows.map(r => Object.assign(r, { sesiId: saved.id })));
  return { sesi: saved, hadir: hs };
}

function delSesi_(s, id) {
  const old = readTable_('Sesi').find(x => x.id === id);
  if (!old) return;
  if (s.role === 'guru' && old.guruId !== s.uid) throw new Error('Sesi ini diinput guru lain.');
  removeRow_('Sesi', id);
  removeWhere_('Hadir', h => h.sesiId === id);
}

/* ---------------- SETTINGS ---------------- */

function getSettings_() {
  const o = Object.assign({}, DEFAULT_SETTINGS);
  readTable_('Pengaturan').forEach(r => { if (r.key) o[r.key] = r.value; });
  return o;
}
function publicSettings_() { const s = getSettings_(); const o = {}; PUBLIC_KEYS.forEach(k => o[k] = s[k]); return o; }

/* ---------------- ISI OTOMATIS (bila data diketik langsung di sheet) ---------------- */
const rand4_ = () => String(Math.floor(1000 + Math.random() * 9000));
function lengkapiDataSiswa() {
  return fillTable_('Siswa', 'DKC', 3, { kodeOrtu: rand4_, status: () => 'Aktif', tglDaftar: () => ENV.today() });
}
function lengkapiDataGuru() {
  return fillTable_('Guru', 'G', 2, { pin: rand4_, status: () => 'Aktif' });
}

/* ---------------- UTIL TANGGAL ---------------- */
function addDays_(iso, n) { const d = new Date(iso + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
function addMonthsIso_(iso, n) {
  const p = String(iso || '').split('-').map(Number); if (!p[0]) return '';
  let y = p[0], m = p[1] + n; while (m > 12) { m -= 12; y++; } while (m < 1) { m += 12; y--; }
  const last = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return y + '-' + String(m).padStart(2, '0') + '-' + String(Math.min(p[2] || 1, last)).padStart(2, '0');
}
function num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  let s = String(v == null ? '' : v).replace(/[^\d.,-]/g, '');
  if (!s) return 0;
  if (/^-?\d{1,3}([.,]\d{3})+$/.test(s)) s = s.replace(/[.,]/g, '');      // 1.600.000 / 1,600,000
  else if (s.indexOf(',') > -1 && s.indexOf('.') < 0) s = s.replace(',', '.'); // 66,67
  else s = s.replace(/,/g, '');
  return Number(s) || 0;
}

/* ================================================================
 *  ENGINE — dihitung ulang dari data mentah setiap kali dibutuhkan
 *  • Paket: kredit = sesi + bonus. Nilai per sesi = total bayar ÷ kredit.
 *  • Pemakaian FIFO per pendaftaran kelas (paket terlama dipakai dulu).
 *  • Hadir → sesi terpotong. Izin/Sakit → pakai jatah izin paket (tidak
 *    terpotong); lewat jatah → terpotong (hangus) sesuai pengaturan.
 *    Alpa → terpotong (hangus) sesuai pengaturan.
 *  • Pindah kelas: baris Pembelian jenis "Pindahan" (asalId = paket asal)
 *    mengurangi kredit paket asal & menambah kredit di kelas tujuan.
 *  • Revenue diakui saat sesi terpakai (bukan saat bayar) → tahu berapa
 *    yang "dipotong" tiap bulan + saldo paket yang belum terpakai.
 *  • Fee guru private  = (nilai kelas × durasi/standar − admin fee) × %guru
 *                        (atau fee manual), + transport (home service).
 *  • Fee guru group    = tabel TarifGroup sesuai jumlah murid HADIR.
 *  • Referral          = % × (nilai − admin) selama N bulan sejak murid daftar.
 * ================================================================ */
function engine_(db, st) {
  st = st || {};
  const T = n => db[n] || [];
  const byId = arr => { const m = {}; arr.forEach(r => m[r.id] = r); return m; };
  const enrolM = byId(T('Enrol')), sesiM = byId(T('Sesi')), tarifM = byId(T('TarifPrivat')), siswaM = byId(T('Siswa'));
  const std = num_(st.durasiStandar) || 45;
  const alpaPotong = st.alpaPotong !== 'Tidak';
  const izinPotong = st.izinLewatJatah !== 'Tidak';
  const guruAlpa = st.guruDibayarAlpa !== 'Tidak';
  const admGroupPct = num_(st.adminFeeGroup);
  const refGroupPct = num_(st.referralPersen);
  const refBulan = num_(st.referralBulan) || 3;
  const warnN = st.warnSisa === undefined ? 1 : num_(st.warnSisa);

  /* ---- paket ---- */
  const pkgs = {}, byEnrol = {};
  T('Pembelian').forEach(p => {
    if (p.status === 'Batal') return;
    const credits = num_(p.sesi) + num_(p.bonus);
    if (credits <= 0) return;
    const k = { id: p.id, p: p, enrolId: p.enrolId, siswaId: p.siswaId, credits: credits, izinQuota: num_(p.jatahIzin),
      perCredit: num_(p.total) / credits, usedN: 0, izinUsed: 0, hangusN: 0, movedOut: 0, moves: [], events: [], tgl: p.tglBeli || '' };
    pkgs[p.id] = k; (byEnrol[p.enrolId] = byEnrol[p.enrolId] || []).push(k);
  });
  T('Pembelian').forEach(p => {
    if (p.status === 'Batal' || !p.asalId || !pkgs[p.asalId]) return;
    pkgs[p.asalId].movedOut += num_(p.sesi) + num_(p.bonus);
    pkgs[p.asalId].moves.push(p);
  });
  Object.keys(byEnrol).forEach(e => byEnrol[e].sort((a, b) => (a.tgl).localeCompare(b.tgl) || String(a.id).localeCompare(String(b.id))));

  /* ---- absensi → pemakaian sesi ---- */
  const evs = [];
  T('Hadir').forEach(h => { const s = sesiM[h.sesiId]; if (s) evs.push({ h: h, s: s }); });
  evs.sort((a, b) => (a.s.tanggal + (a.s.jam || '')).localeCompare(b.s.tanggal + (b.s.jam || '')));
  const events = [];
  evs.forEach(x => {
    const h = x.h, s = x.s, eid = h.enrolId;
    const list = byEnrol[eid] || [];
    const avail = list.find(k => k.usedN + k.movedOut < k.credits);
    const ev = { id: h.id, sesiId: s.id, tgl: s.tanggal, jam: s.jam, siswaId: h.siswaId, enrolId: eid, status: h.status, guruId: s.guruId,
      kelasId: s.kelasId, tipe: s.tipe, durasi: num_(s.durasi), mode: '', charge: false, pkgId: '', value: 0, kurang: false };
    if (h.status === 'Hadir') { ev.mode = 'pakai'; ev.charge = true; }
    else if (h.status === 'Alpa') { ev.mode = alpaPotong ? 'hangus' : 'bebas'; ev.charge = alpaPotong; }
    else {
      const cur = avail || list[list.length - 1];
      if (cur && cur.izinUsed < cur.izinQuota) { ev.mode = 'izin'; cur.izinUsed++; ev.pkgId = cur.id; cur.events.push(ev); }
      else if (izinPotong) { ev.mode = 'hangus'; ev.charge = true; }
      else ev.mode = 'bebas';
    }
    if (ev.charge) {
      if (avail) { avail.usedN++; if (ev.mode === 'hangus') avail.hangusN++; ev.pkgId = avail.id; ev.value = avail.perCredit; avail.events.push(ev); }
      else ev.kurang = true;
    }
    events.push(ev);
  });

  /* ---- ringkasan per pendaftaran ---- */
  const enrol = {};
  T('Enrol').forEach(e => {
    const list = byEnrol[e.id] || [];
    const credits = list.reduce((a, k) => a + k.credits, 0);
    const used = list.reduce((a, k) => a + k.usedN, 0);
    const moved = list.reduce((a, k) => a + k.movedOut, 0);
    const izin = list.reduce((a, k) => a + k.izinUsed, 0);
    const kurang = events.filter(v => v.enrolId === e.id && v.kurang).length;
    const sisa = credits - used - moved - kurang;
    enrol[e.id] = { id: e.id, credits: credits, used: used, moved: moved, izin: izin, kurang: kurang, sisa: sisa, pkgs: list.map(k => k.id),
      warn: sisa <= 0 ? 'habis' : sisa <= warnN ? 'hampir' : '' };
    list.forEach(k => k.sisa = k.credits - k.usedN - k.movedOut);
  });

  /* ---- tarif ---- */
  const tarifFor = (e, guruId, tgl) => {
    let t = tarifM[e.tarifId];
    if (t && (!guruId || t.guruId === guruId)) return t;
    const cand = T('TarifPrivat').filter(r => r.guruId === guruId && r.kategori === (t ? t.kategori : e.kategori) && r.lokasi === (t ? t.lokasi : e.lokasi) && r.aktif !== 'Tidak' && (r.tglBerlaku || '') <= tgl)
      .sort((a, b) => (b.tglBerlaku || '').localeCompare(a.tglBerlaku || '') || String(b.versi).localeCompare(String(a.versi)));
    return cand[0] || null;
  };
  const groupFee = (guruId, tgl, n) => {
    if (n <= 0) return 0;
    const rows = T('TarifGroup').filter(r => r.guruId === guruId && (r.tglBerlaku || '') <= tgl);
    if (!rows.length) return 0;
    const latest = rows.reduce((a, r) => (r.tglBerlaku > a ? r.tglBerlaku : a), '');
    const cur = rows.filter(r => r.tglBerlaku === latest).sort((a, b) => num_(a.jumlahMurid) - num_(b.jumlahMurid));
    let pick = null; cur.forEach(r => { if (num_(r.jumlahMurid) <= n) pick = r; });
    return num_((pick || cur[0]).feeGuru);
  };
  const refInfo = (siswaId, tgl) => {
    const x = siswaM[siswaId];
    if (!x || !(x.referralDari || x.referralGuruId) || !x.tglDaftar) return null;
    const end = addMonthsIso_(x.tglDaftar, refBulan);
    if (tgl < x.tglDaftar || tgl >= end) return null;
    return { nama: x.referralDari || '', guruId: x.referralGuruId || '' };
  };

  /* ---- keuangan per kejadian ---- */
  const fin = [];      // per murid per sesi (private: fee guru di sini)
  const groupSesi = {}; // per sesi group (fee guru group di sini)
  events.forEach(ev => {
    const e = enrolM[ev.enrolId] || {};
    const isGroup = (e.tipe || ev.tipe) === 'Group';
    const f = { ev: ev, tipe: isGroup ? 'Group' : 'Private', kategori: isGroup ? 'Group Class' : (e.kategori || 'Private'), lokasi: isGroup ? '' : (e.lokasi || ''),
      nilai: ev.value, feeGuru: 0, travel: 0, admin: 0, classroom: 0, referral: 0, refNama: '', refGuruId: '', guruId: ev.guruId, tarifMissing: false };
    if (isGroup) {
      const g = groupSesi[ev.sesiId] || (groupSesi[ev.sesiId] = { sesiId: ev.sesiId, tgl: ev.tgl, guruId: ev.guruId, kelasId: ev.kelasId, n: 0, nilai: 0, fee: 0 });
      if (ev.status === 'Hadir') g.n++;
      g.nilai += ev.value;
      if (ev.value) {
        f.admin = Math.round(ev.value * admGroupPct / 100);
        const r = refInfo(ev.siswaId, ev.tgl);
        if (r) { f.referral = Math.round((ev.value - f.admin) * refGroupPct / 100); f.refNama = r.nama; f.refGuruId = r.guruId; }
      }
    } else {
      const teach = ev.status === 'Hadir' || (ev.status === 'Alpa' && ev.charge && guruAlpa);
      const t = tarifFor(e, ev.guruId, ev.tgl);
      if (!t) f.tarifMissing = teach;
      if (t && (teach || ev.charge)) {
        const dur = ev.durasi || num_(e.durasi) || std;
        const fac = dur / std;
        const cv = num_(t.nilaiKelas) * fac;
        const admRaw = String(t.adminFee || '');
        const adm = /%/.test(admRaw) ? Math.round(cv * num_(admRaw) / 100) : num_(admRaw);
        if (ev.charge) { f.admin = adm; f.classroom = num_(t.classroomFee); }
        if (teach) {
          f.feeGuru = num_(t.feeGuruManual) ? Math.round(num_(t.feeGuruManual) * fac) : Math.round((cv - adm) * num_(t.persenGuru) / 100);
          f.travel = num_(e.travelFee !== '' && e.travelFee != null ? e.travelFee : t.travelFee);
        }
        if (ev.charge) {
          const r = refInfo(ev.siswaId, ev.tgl);
          if (r) { f.referral = Math.round((cv - adm) * num_(t.persenReferral) / 100); f.refNama = r.nama; f.refGuruId = r.guruId; }
        }
      }
    }
    fin.push(f);
  });
  Object.keys(groupSesi).forEach(id => { const g = groupSesi[id]; g.fee = groupFee(g.guruId, g.tgl, g.n); });

  return { pkgs: pkgs, byEnrol: byEnrol, events: events, enrol: enrol, fin: fin, groupSesi: groupSesi };
}

/* ---- rekap bulanan (dipakai halaman Keuangan & Gaji) ---- */
function rekapBulan_(db, st, L, periode) {
  const inP = d => String(d || '').slice(0, 7) === periode;
  const end = periode + '-31';
  const out = { periode: periode, kasMasuk: 0, kasPerJenis: {}, diakui: 0, diakuiSesi: 0, diakuiNonSesi: 0, hangus: 0,
    feeGuru: 0, travel: 0, admin: 0, classroom: 0, referral: 0, margin: 0, saldoPaket: 0, perKategori: {}, guru: {}, referral_: [] };
  (db.Pembelian || []).forEach(p => {
    if (p.status !== 'Lunas' || p.asalId || !inP(p.tglBayar)) return;
    out.kasMasuk += num_(p.total);
    out.kasPerJenis[p.jenis || 'Lainnya'] = (out.kasPerJenis[p.jenis || 'Lainnya'] || 0) + num_(p.total);
    if (num_(p.sesi) + num_(p.bonus) <= 0) { out.diakuiNonSesi += num_(p.total); }
  });
  const kat = k => out.perKategori[k] || (out.perKategori[k] = { nilai: 0, sesi: 0, feeGuru: 0, travel: 0, admin: 0, classroom: 0, referral: 0 });
  const gr = id => out.guru[id] || (out.guru[id] = { privatSesi: 0, privatFee: 0, travel: 0, groupSesi: 0, groupFee: 0, referral: 0, rincian: [] });
  L.fin.forEach(f => {
    if (!inP(f.ev.tgl)) return;
    const k = kat(f.kategori);
    k.nilai += f.nilai; if (f.ev.charge) k.sesi++;
    k.feeGuru += f.feeGuru; k.travel += f.travel; k.admin += f.admin; k.classroom += f.classroom; k.referral += f.referral;
    out.diakuiSesi += f.nilai;
    if (f.ev.mode === 'hangus') out.hangus += f.nilai;
    out.feeGuru += f.feeGuru; out.travel += f.travel; out.admin += f.admin; out.classroom += f.classroom; out.referral += f.referral;
    if (f.tipe === 'Private' && (f.feeGuru || f.travel || f.tarifMissing)) {
      const g = gr(f.guruId); g.privatSesi++; g.privatFee += f.feeGuru; g.travel += f.travel;
      g.rincian.push({ tgl: f.ev.tgl, jenis: f.kategori, lokasi: f.lokasi, siswaId: f.ev.siswaId, status: f.ev.status, fee: f.feeGuru, travel: f.travel, durasi: f.ev.durasi, missing: f.tarifMissing });
    }
    if (f.referral) {
      if (f.refGuruId) { const g = gr(f.refGuruId); g.referral += f.referral; g.rincian.push({ tgl: f.ev.tgl, jenis: 'Referral', siswaId: f.ev.siswaId, fee: f.referral }); }
      out.referral_.push({ tgl: f.ev.tgl, siswaId: f.ev.siswaId, nama: f.refNama, guruId: f.refGuruId, nilai: f.referral });
    }
  });
  Object.keys(L.groupSesi).forEach(id => {
    const g = L.groupSesi[id]; if (!inP(g.tgl)) return;
    const k = kat('Group Class'); k.feeGuru += g.fee; out.feeGuru += g.fee;
    const r = gr(g.guruId); r.groupSesi++; r.groupFee += g.fee;
    r.rincian.push({ tgl: g.tgl, jenis: 'Group Class', kelasId: g.kelasId, n: g.n, fee: g.fee, sesiId: id });
  });
  out.diakui = out.diakuiSesi + out.diakuiNonSesi;
  out.margin = out.diakui - out.feeGuru - out.travel - out.referral;
  // saldo paket (belum terpakai) per akhir bulan
  Object.keys(L.pkgs).forEach(id => {
    const k = L.pkgs[id]; if ((k.tgl || '') > end) return;
    const used = k.events.filter(v => v.charge && v.tgl <= end).length;
    const moved = k.moves.filter(m => (m.tglBeli || '') <= end).reduce((a, m) => a + num_(m.sesi) + num_(m.bonus), 0);
    const sisa = Math.max(0, k.credits - used - moved);
    out.saldoPaket += sisa * k.perCredit;
  });
  return out;
}

/* ---------------- ID otomatis (dipakai adapter penyimpanan) ---------------- */
function newId_(name, existing) {
  const seq = (prefix, digits) => {
    let max = 0;
    existing.forEach(r => { const m = String(r.id).match(new RegExp('^' + prefix + '(\\d+)$')); if (m && +m[1] > max) max = +m[1]; });
    return prefix + String(max + 1).padStart(digits, '0');
  };
  if (name === 'Siswa') return seq('DKC', 3);
  if (name === 'Guru') return seq('G', 2);
  if (name === 'KelasGroup') return seq('KG', 2);
  return name.slice(0, 2).toUpperCase() + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}
function autoFields_(name, r, existing) {
  if (name === 'Pembelian' && !r.noInvoice) {
    const ym = String(r.tglBeli || ENV.today()).slice(0, 7).replace('-', '');
    let max = 0;
    existing.forEach(x => { const m = String(x.noInvoice || '').match(new RegExp('^INV-' + ym + '-(\\d+)$')); if (m && +m[1] > max) max = +m[1]; });
    r.noInvoice = 'INV-' + ym + '-' + String(max + 1).padStart(3, '0');
  }
  if (name === 'Siswa') { if (!r.kodeOrtu) r.kodeOrtu = rand4_(); if (!r.status) r.status = 'Aktif'; if (!r.tglDaftar) r.tglDaftar = ENV.today(); }
  if (name === 'Guru') { if (!r.pin) r.pin = rand4_(); if (!r.status) r.status = 'Aktif'; }
}
