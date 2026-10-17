/**
 * ============================================================
 *  BACKEND  —  Sistem SPPD (Surat Perintah Tugas & Perjalanan Dinas)
 *  MA Al-Ma'tuq
 * ============================================================
 *  Cara pasang:
 *   1. Buka Google Sheet kosong  →  Extensions / Ekstensi  →  Apps Script.
 *   2. Hapus isi Code.gs, tempel seluruh kode ini, simpan.
 *   3. Jalankan fungsi  setupAwal  satu kali (beri izin akses).
 *   4. Deploy  →  New deployment  →  Web app
 *        Execute as : Me
 *        Who has access : Anyone
 *   5. Salin URL Web App (berakhiran /exec) ke halaman login aplikasi.
 *
 *  Kata sandi awal: admin123  (segera ganti di menu Pengaturan).
 *  Kop surat, stempel & tanda tangan disimpan di sheet "Aset"
 *  (di-upload dari aplikasi, tidak perlu Google Drive).
 * ============================================================
 */

const SH_SPPD = 'SPPD';
const SH_PEGAWAI = 'Pegawai';
const SH_SETTING = 'Pengaturan';
const SH_ASET = 'Aset';

const DEFAULT_PASSWORD = 'admin123';
const TOKEN_TTL = 21600;        // 6 jam (detik)
const CHUNK = 40000;            // batas aman 1 sel = 50.000 karakter

const SPPD_COLS = [
  'id', 'dibuat', 'diubah', 'nomor', 'urut', 'tahun',
  'pegawaiId', 'nama', 'nip', 'pangkat', 'jabatan', 'gaji', 'tingkat', 'asalMadrasah',
  'kegiatan', 'tglBerangkat', 'tglKembali', 'pukul', 'tempatTujuan',
  'maksud', 'alat', 'tempatBerangkat', 'pembebanan', 'mataAnggaran',
  'keteranganLain', 'tglSurat', 'tempatSurat', 'catatan'
];
const PEGAWAI_COLS = ['id', 'nama', 'nip', 'pangkat', 'jabatan', 'gaji', 'tingkat'];

const DEFAULT_SETTINGS = {
  nama_kepala: 'Irwansyah Ramdani, S.S., M.Pd.',
  nip_kepala: '',
  jabatan_kepala: 'Kepala Madrasah',
  nama_madrasah: 'MA AL-MA’TUQ',
  pejabat_pemberi: 'Kepala Madrasah Aliyah Al-Ma’tuq',
  tempat_surat: 'Cisaat',
  template_nomor: 'Mas.10.02/{urut}/PP.00.6/021/{bulan}/{tahun}',
  tempat_berangkat_default: 'MAS Al-Ma’tuq',
  pembebanan_default: 'MAS Al-Ma’tuq',
  alat_default: 'Kendaraan Pribadi',
  ttd_tinggi_mm: '22',
  stempel_ukuran_mm: '30',
  stempel_geser_x_mm: '-14',
  stempel_geser_y_mm: '0',
  aset_versi: '0'
};

/* ---------------------------------------------------------------
 *  SETUP
 * ------------------------------------------------------------- */
function setupAwal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SH_SPPD, SPPD_COLS);
  ensureSheet_(ss, SH_PEGAWAI, PEGAWAI_COLS);
  ensureSheet_(ss, SH_ASET, ['kunci', 'urutan', 'data']);
  let st = ss.getSheetByName(SH_SETTING);
  if (!st) {
    st = ss.insertSheet(SH_SETTING);
    st.getRange(1, 1, 1, 2).setValues([['kunci', 'nilai']]).setFontWeight('bold');
    st.getRange('A:B').setNumberFormat('@');
    const rows = Object.keys(DEFAULT_SETTINGS).map(k => [k, DEFAULT_SETTINGS[k]]);
    st.getRange(2, 1, rows.length, 2).setValues(rows);
    st.setFrozenRows(1);
    st.setColumnWidth(1, 220); st.setColumnWidth(2, 420);
  }
  const first = ss.getSheetByName('Sheet1') || ss.getSheetByName('Sheet 1') || ss.getSheetByName('Lembar 1');
  if (first && ss.getSheets().length > 1) { try { ss.deleteSheet(first); } catch (e) {} }
  return 'Setup selesai. Silakan Deploy sebagai Web App.';
}

function ensureSheet_(ss, name, cols) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#e6f2ea');
    sh.setFrozenRows(1);
  }
  // semua kolom sebagai teks supaya tanggal/nomor tidak diubah otomatis oleh Sheets
  sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
  return sh;
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { setupAwal(); sh = ss.getSheetByName(name); }
  return sh;
}

/* ---------------------------------------------------------------
 *  ENTRY POINT
 * ------------------------------------------------------------- */
function doGet() {
  return out_({ ok: true, app: 'SPPD MA Al-Matuq', info: 'API aktif. Gunakan metode POST.' });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return out_({ ok: false, error: 'Permintaan tidak valid.' });
  }
  try {
    if (req.action === 'ping') return out_({ ok: true });
    if (req.action === 'login') return out_(login_(req));
    if (!checkToken_(req.token)) return out_({ ok: false, error: 'Sesi berakhir, silakan login ulang.', auth: false });

    const writes = ['saveSppd', 'deleteSppd', 'savePegawai', 'deletePegawai', 'saveSettings', 'uploadAset', 'deleteAset', 'changePassword', 'nextNumber'];
    let lock = null;
    if (writes.indexOf(req.action) >= 0) { lock = LockService.getScriptLock(); lock.waitLock(20000); }
    try {
      switch (req.action) {
        case 'bootstrap': return out_(bootstrap_());
        case 'getAset': return out_({ ok: true, aset: getAset_() });
        case 'saveSppd': return out_(saveSppd_(req.data));
        case 'deleteSppd': return out_(deleteRow_(SH_SPPD, req.id));
        case 'savePegawai': return out_(savePegawai_(req.data));
        case 'deletePegawai': return out_(deleteRow_(SH_PEGAWAI, req.id));
        case 'saveSettings': return out_(saveSettings_(req.data));
        case 'uploadAset': return out_(uploadAset_(req.jenis, req.dataUrl));
        case 'deleteAset': return out_(deleteAset_(req.jenis));
        case 'changePassword': return out_(changePassword_(req.lama, req.baru));
        case 'nextNumber': return out_(nextNumber_(req.tahun, req.bulan));
        default: return out_({ ok: false, error: 'Aksi tidak dikenal: ' + req.action });
      }
    } finally {
      if (lock) lock.releaseLock();
    }
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------------------------------------------------------
 *  AUTH
 * ------------------------------------------------------------- */
function hash_(s) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'sppd-almatuq|' + s, Utilities.Charset.UTF_8);
  return raw.map(b => ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2)).join('');
}
function storedHash_() {
  return PropertiesService.getScriptProperties().getProperty('PASSWORD_HASH') || hash_(DEFAULT_PASSWORD);
}
function login_(req) {
  if (hash_(String(req.password || '')) !== storedHash_()) {
    Utilities.sleep(800);
    return { ok: false, error: 'Kata sandi salah.' };
  }
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('tok_' + token, '1', TOKEN_TTL);
  const isDefault = !PropertiesService.getScriptProperties().getProperty('PASSWORD_HASH');
  return { ok: true, token: token, passwordDefault: isDefault };
}
function checkToken_(t) {
  return !!t && CacheService.getScriptCache().get('tok_' + t) === '1';
}
function changePassword_(lama, baru) {
  if (hash_(String(lama || '')) !== storedHash_()) return { ok: false, error: 'Kata sandi lama salah.' };
  if (!baru || String(baru).length < 6) return { ok: false, error: 'Kata sandi baru minimal 6 karakter.' };
  PropertiesService.getScriptProperties().setProperty('PASSWORD_HASH', hash_(String(baru)));
  return { ok: true };
}

/* ---------------------------------------------------------------
 *  UTIL BACA / TULIS
 * ------------------------------------------------------------- */
function clean_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Jakarta', 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}

function readAll_(name, cols) {
  const sh = sheet_(name);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const vals = sh.getRange(2, 1, n, cols.length).getValues();
  return vals.filter(r => r[0] !== '').map(r => {
    const o = {};
    cols.forEach((c, i) => o[c] = clean_(r[i]));
    return o;
  });
}

function findRow_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function deleteRow_(name, id) {
  const sh = sheet_(name);
  const r = findRow_(sh, id);
  if (r < 0) return { ok: false, error: 'Data tidak ditemukan.' };
  sh.deleteRow(r);
  return { ok: true };
}

function upsert_(name, cols, data, prefix) {
  const sh = sheet_(name);
  const now = Utilities.formatDate(new Date(), 'Asia/Jakarta', "yyyy-MM-dd'T'HH:mm:ss");
  data = data || {};
  let r = data.id ? findRow_(sh, data.id) : -1;
  if (r < 0) {
    data.id = prefix + Utilities.getUuid().slice(0, 8);
    if (cols.indexOf('dibuat') >= 0) data.dibuat = now;
    r = sh.getLastRow() + 1;
  } else if (cols.indexOf('dibuat') >= 0) {
    data.dibuat = clean_(sh.getRange(r, cols.indexOf('dibuat') + 1).getValue());
  }
  if (cols.indexOf('diubah') >= 0) data.diubah = now;
  const row = cols.map(c => data[c] === undefined || data[c] === null ? '' : String(data[c]));
  sh.getRange(r, 1, 1, cols.length).setNumberFormat('@').setValues([row]);
  return data;
}

/* ---------------------------------------------------------------
 *  SPPD
 * ------------------------------------------------------------- */
function saveSppd_(d) {
  if (!d || !d.nama) return { ok: false, error: 'Nama pegawai wajib diisi.' };
  if (!d.tglBerangkat) return { ok: false, error: 'Tanggal berangkat wajib diisi.' };
  const saved = upsert_(SH_SPPD, SPPD_COLS, d, 'S');
  return { ok: true, data: saved };
}

/** Nomor urut berikutnya pada tahun tertentu → format sesuai template. */
function nextNumber_(tahun, bulan) {
  const list = readAll_(SH_SPPD, SPPD_COLS);
  let max = 0;
  list.forEach(s => {
    if (String(s.tahun) === String(tahun)) {
      const n = parseInt(s.urut, 10);
      if (!isNaN(n) && n > max) max = n;
    }
  });
  const urut = ('00' + (max + 1)).slice(-3);
  const st = getSettings_();
  const romawi = ['', 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII'][parseInt(bulan, 10)] || '';
  const nomor = String(st.template_nomor || DEFAULT_SETTINGS.template_nomor)
    .replace(/\{urut\}/g, urut).replace(/\{bulan\}/g, romawi).replace(/\{tahun\}/g, tahun);
  return { ok: true, urut: urut, nomor: nomor };
}

/* ---------------------------------------------------------------
 *  PEGAWAI
 * ------------------------------------------------------------- */
function savePegawai_(d) {
  if (!d || !d.nama) return { ok: false, error: 'Nama pegawai wajib diisi.' };
  return { ok: true, data: upsert_(SH_PEGAWAI, PEGAWAI_COLS, d, 'P') };
}

/* ---------------------------------------------------------------
 *  PENGATURAN
 * ------------------------------------------------------------- */
function getSettings_() {
  const sh = sheet_(SH_SETTING);
  const o = {};
  Object.keys(DEFAULT_SETTINGS).forEach(k => o[k] = DEFAULT_SETTINGS[k]);
  const n = sh.getLastRow() - 1;
  if (n > 0) sh.getRange(2, 1, n, 2).getValues().forEach(r => { if (r[0]) o[String(r[0])] = clean_(r[1]); });
  return o;
}

function setSetting_(key, value) {
  const sh = sheet_(SH_SETTING);
  const n = sh.getLastRow() - 1;
  if (n > 0) {
    const keys = sh.getRange(2, 1, n, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0]) === key) { sh.getRange(i + 2, 2).setNumberFormat('@').setValue(String(value)); return; }
    }
  }
  sh.getRange(sh.getLastRow() + 1, 1, 1, 2).setNumberFormat('@').setValues([[key, String(value)]]);
}

function saveSettings_(d) {
  Object.keys(d || {}).forEach(k => {
    if (k === 'aset_versi') return;           // hanya diubah oleh upload aset
    if (k in DEFAULT_SETTINGS) setSetting_(k, d[k]);
  });
  return { ok: true, settings: getSettings_() };
}

/* ---------------------------------------------------------------
 *  ASET  (kop, stempel, tanda tangan)  — disimpan di sheet "Aset"
 * ------------------------------------------------------------- */
const JENIS_ASET = ['kop', 'stempel', 'ttd'];

function getAset_() {
  const sh = sheet_(SH_ASET);
  const n = sh.getLastRow() - 1;
  const res = { kop: '', stempel: '', ttd: '' };
  if (n < 1) return res;
  const rows = sh.getRange(2, 1, n, 3).getValues();
  const buf = {};
  rows.forEach(r => {
    const k = String(r[0]);
    if (!buf[k]) buf[k] = [];
    buf[k][parseInt(r[1], 10)] = String(r[2]);
  });
  JENIS_ASET.forEach(k => { if (buf[k]) res[k] = buf[k].join(''); });
  return res;
}

function clearAset_(jenis) {
  const sh = sheet_(SH_ASET);
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const keys = sh.getRange(2, 1, n, 1).getValues();
  for (let i = keys.length - 1; i >= 0; i--) {
    if (String(keys[i][0]) === jenis) sh.deleteRow(i + 2);
  }
}

function uploadAset_(jenis, dataUrl) {
  if (JENIS_ASET.indexOf(jenis) < 0) return { ok: false, error: 'Jenis aset tidak valid.' };
  if (!/^data:image\/(png|jpeg|jpg|webp);base64,/.test(String(dataUrl || ''))) return { ok: false, error: 'File harus berupa gambar PNG/JPG.' };
  if (dataUrl.length > 900000) return { ok: false, error: 'Gambar terlalu besar. Perkecil ukuran gambar.' };
  clearAset_(jenis);
  const sh = sheet_(SH_ASET);
  const rows = [];
  for (let i = 0, k = 0; i < dataUrl.length; i += CHUNK, k++) rows.push([jenis, String(k), dataUrl.substr(i, CHUNK)]);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setNumberFormat('@').setValues(rows);
  const ver = String(Date.now());
  setSetting_('aset_versi', ver);
  return { ok: true, versi: ver };
}

function deleteAset_(jenis) {
  if (JENIS_ASET.indexOf(jenis) < 0) return { ok: false, error: 'Jenis aset tidak valid.' };
  clearAset_(jenis);
  const ver = String(Date.now());
  setSetting_('aset_versi', ver);
  return { ok: true, versi: ver };
}

/* ---------------------------------------------------------------
 *  BOOTSTRAP — semua data awal dalam satu panggilan
 * ------------------------------------------------------------- */
function bootstrap_() {
  const sppd = readAll_(SH_SPPD, SPPD_COLS).sort((a, b) => String(b.dibuat).localeCompare(String(a.dibuat)));
  return {
    ok: true,
    settings: getSettings_(),
    pegawai: readAll_(SH_PEGAWAI, PEGAWAI_COLS).sort((a, b) => a.nama.localeCompare(b.nama)),
    sppd: sppd
  };
}
