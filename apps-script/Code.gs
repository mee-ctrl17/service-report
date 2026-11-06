/**
 * Maintenance History Report — Code.gs (v3, API untuk GitHub Pages)
 *
 * Tampilan (index.html) di-host di GitHub Pages, script ini hanya melayani data lewat API.
 *
 * PENGATURAN (Project Settings > Script Properties):
 *   APP_TOKEN    = kode akses tim, wajib diisi. Diminta sekali di tiap perangkat.
 *   APP_PASSWORD = password admin untuk edit dan hapus. Default "admin123" jika kosong.
 *
 * DEPLOY: Terapkan > Kelola penerapan > pensil > Versi baru,
 *   Jalankan sebagai: Saya, Yang memiliki akses: Siapa saja (Anyone).
 */

const CONFIG = {
  SHEET_DATA: 'Data',
  SHEET_CUSTOMER: 'Customer',
  SHEET_SPAREPART: 'Sparepart',
  SHEET_TEKNISI: 'Teknisi',
  FOLDER_FOTO: 'Dokumentasi_Maintenance',
  FOLDER_TTD: 'TandaTangan_Maintenance',
  DEFAULT_TEKNISI: ['Faisal R', 'Agus S', 'Lukman A'],
  DEFAULT_PASSWORD: 'admin123'
};

const HEADERS = [
  'Tanggal', 'Nama Customer', 'Tipe Mesin', 'Serial Number', 'Jenis Maintenance',
  'Teknisi', 'Problem', 'Solusi / Tindakan', 'Status Mesin', 'Status Pengerjaan',
  'Link Foto', 'Sparepart', 'No. Report', 'Link Tanda Tangan', 'Nama Penerima'
];
const NUM_COLS = HEADERS.length;

const BULAN_ID = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
const BULAN_EN = ['january', 'february', 'march', 'april', 'may', 'june',
  'july', 'august', 'september', 'october', 'november', 'december'];

// ======================= API =======================

/** Daftar aksi yang boleh dipanggil dari halaman. */
const API = {
  ping: () => ({ pong: true }),
  getInitialData: () => getInitialData_(),
  getSparepartList: tipe => getSparepartList_(tipe),
  verifyPassword: pwd => verifyPassword_(pwd),
  saveRecord: data => saveRecord_(data),
  deleteRecord: (row, sn, pwd) => deleteRecord_(row, sn, pwd)
};

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Cek cepat dari browser: buka link /exec, harus muncul {"ok":true,...}. */
function doGet() {
  return json_({ ok: true, app: 'Maintenance History Report API', version: 3 });
}

/**
 * Halaman mengirim POST dengan body JSON: { action, args, token }.
 * Content-Type dari halaman adalah text/plain agar tidak terkena preflight CORS.
 */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const token = PropertiesService.getScriptProperties().getProperty('APP_TOKEN');
    if (!token) {
      return json_({ ok: false, code: 'SETUP', error: 'Kode akses belum diatur. Isi APP_TOKEN di Script Properties.' });
    }
    if (str_(body.token) !== token) {
      return json_({ ok: false, code: 'AUTH', error: 'Kode akses salah atau sudah diganti.' });
    }
    const fn = API[body.action];
    if (!fn) return json_({ ok: false, error: 'Aksi tidak dikenal: ' + body.action });
    const args = Array.isArray(body.args) ? body.args : [];
    return json_({ ok: true, data: fn.apply(null, args) });
  } catch (err) {
    return json_({ ok: false, error: (err && err.message) ? err.message : String(err) });
  }
}

// ======================= HELPER =======================

function str_(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}

function tz_() {
  return Session.getScriptTimeZone() || 'Asia/Jakarta';
}

function toISO_(d) {
  return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd');
}

function formatIndo_(d) {
  const p = toISO_(d).split('-');
  return p[2] + ' ' + BULAN_ID[parseInt(p[1], 10) - 1] + ' ' + p[0];
}

/** Membaca tanggal dari Date, "2026-11-05", "05 November 2026", "05 Mei 2026", atau "05/11/2026". */
function parseTanggal_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  const s = str_(v);
  if (!s) return null;

  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], 12);

  m = s.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/);
  if (m) {
    const nama = m[2].toLowerCase();
    let idx = BULAN_ID.findIndex(b => b.toLowerCase() === nama);
    if (idx < 0) idx = BULAN_EN.indexOf(nama);
    if (idx >= 0) return new Date(+m[3], idx, +m[1], 12);
  }

  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1], 12);

  return null;
}

/** Mencegah teks yang diawali "=" / "+" / "@" dibaca Sheets sebagai rumus. */
function sanitize_(v) {
  const s = str_(v);
  return /^[=+@]/.test(s) ? "'" + s : s;
}

function getDataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(CONFIG.SHEET_DATA) || ss.getSheets()[0];
}

function ensureHeaders_(sheet) {
  if (sheet.getMaxColumns() < NUM_COLS) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), NUM_COLS - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, NUM_COLS).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    return;
  }
  const cur = sheet.getRange(1, 1, 1, NUM_COLS).getValues()[0];
  let changed = false;
  for (let i = 0; i < NUM_COLS; i++) {
    if (!str_(cur[i])) { cur[i] = HEADERS[i]; changed = true; }
  }
  if (changed) sheet.getRange(1, 1, 1, NUM_COLS).setValues([cur]);
}

function getFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function uploadBase64_(fileObj, folder) {
  const blob = Utilities.newBlob(
    Utilities.base64Decode(fileObj.bytes),
    fileObj.mimeType || 'application/octet-stream',
    fileObj.filename || ('file_' + Date.now())
  );
  const file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    console.warn('Tidak bisa mengatur sharing file: ' + e);
  }
  return file.getUrl();
}

function uploadFiles_(filesData, folderName) {
  if (!filesData || !filesData.length) return [];
  const folder = getFolder_(folderName);
  return filesData.map(f => uploadBase64_(f, folder));
}

// ======================= PASSWORD =======================

function getPassword_() {
  return PropertiesService.getScriptProperties().getProperty('APP_PASSWORD') || CONFIG.DEFAULT_PASSWORD;
}

function checkPassword_(pwd) {
  if (str_(pwd) !== getPassword_()) throw new Error('Password salah.');
}

/** Dipanggil dari halaman sebelum masuk mode edit. */
function verifyPassword_(pwd) {
  checkPassword_(pwd);
  return true;
}

/** Alternatif mengganti password/kode akses: ubah teksnya, lalu Run sekali dari editor. */
function aturPasswordDanKodeAkses() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('APP_PASSWORD', 'GANTI_PASSWORD_ADMIN');
  props.setProperty('APP_TOKEN', 'GANTI_KODE_AKSES_TIM');
}

// ======================= BACA DATA =======================

/** Satu panggilan untuk memuat semua data awal halaman. */
function getInitialData_() {
  return {
    records: getData_(),
    customers: getCustomerList_(),
    teknisi: getTeknisiList_()
  };
}

function getData_() {
  const sheet = getDataSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const width = Math.min(Math.max(sheet.getLastColumn(), 12), NUM_COLS);
  const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  const out = [];

  values.forEach((r, i) => {
    if (r.every(c => str_(c) === '')) return;
    const d = parseTanggal_(r[0]);
    out.push({
      row: i + 2, // nomor baris ASLI di Sheet
      tanggalISO: d ? toISO_(d) : '',
      tanggal: d ? formatIndo_(d) : str_(r[0]),
      ts: d ? d.getTime() : 0,
      customer: str_(r[1]),
      tipe: str_(r[2]),
      sn: str_(r[3]),
      jenis: str_(r[4]),
      teknisi: str_(r[5]),
      problem: str_(r[6]),
      solusi: str_(r[7]),
      statusMesin: str_(r[8]),
      statusPengerjaan: str_(r[9]),
      foto: str_(r[10]).split(',').map(s => s.trim()).filter(Boolean),
      sparepart: str_(r[11]),
      noReport: str_(r[12]),
      ttd: str_(r[13]),
      penerima: str_(r[14])
    });
  });

  // Terbaru di atas; jika tanggal sama, baris yang lebih baru di atas
  out.sort((a, b) => (b.ts - a.ts) || (b.row - a.row));
  return out;
}

function getCustomerList_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_CUSTOMER);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  return sheet.getRange(2, 1, lastRow - 1, 3).getValues()
    .map(r => [str_(r[0]), str_(r[1]), str_(r[2])])
    .filter(r => r[0] || r[1] || r[2]);
}

function getTeknisiList_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_TEKNISI);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_TEKNISI);
    sheet.getRange(1, 1, CONFIG.DEFAULT_TEKNISI.length + 1, 1)
      .setValues([['Nama Teknisi']].concat(CONFIG.DEFAULT_TEKNISI.map(n => [n])));
    sheet.setFrozenRows(1);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  return sheet.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(r => str_(r[0]))
    .filter(Boolean);
}

function getSparepartList_(tipeMesin) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SPAREPART);
  if (!sheet || !tipeMesin) return [];

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol < 1) return [];

  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const norm = v => str_(v).toLowerCase().replace(/[\s.]+/g, '');
  const key = norm(tipeMesin);
  const col = data[0].findIndex(h => norm(h) === key);
  if (col < 0) return [];

  const parts = [];
  for (let row = 1; row < data.length; row++) {
    const p = str_(data[row][col]);
    if (p && parts.indexOf(p) < 0) parts.push(p);
  }
  return parts;
}

// ======================= SIMPAN / HAPUS =======================

function generateNoReport_(tanggal) {
  const ym = Utilities.formatDate(tanggal, tz_(), 'yyyyMM');
  const props = PropertiesService.getScriptProperties();
  const key = 'SR_COUNTER_' + ym;
  const n = parseInt(props.getProperty(key) || '0', 10) + 1;
  props.setProperty(key, String(n));
  return 'SR-' + ym + '-' + ('000' + n).slice(-4);
}

/**
 * Tambah data baru (data.row kosong) atau edit (data.row berisi nomor baris asli).
 * Edit wajib menyertakan data.password dan data.originalSn untuk validasi.
 */
function saveRecord_(data) {
  if (!data) throw new Error('Data kosong.');
  const isEdit = !!data.row;
  if (isEdit) checkPassword_(data.password);

  const wajib = {
    customer: 'Nama Customer', tipeMesin: 'Tipe Mesin', serialNumber: 'Serial Number',
    kategori: 'Jenis Maintenance', teknisi: 'Teknisi',
    statusMesin: 'Status Mesin', statusPengerjaan: 'Status Pengerjaan'
  };
  Object.keys(wajib).forEach(k => {
    if (!str_(data[k])) throw new Error('Kolom "' + wajib[k] + '" wajib diisi.');
  });

  const tanggal = parseTanggal_(data.tanggal);
  if (!tanggal) throw new Error('Tanggal tidak valid.');

  // Upload dilakukan sebelum lock supaya lock tidak tertahan lama
  const fotoBaru = uploadFiles_(data.filesData, CONFIG.FOLDER_FOTO);
  let ttdBaru = '';
  if (data.signature && data.signature.indexOf(',') > -1) {
    ttdBaru = uploadBase64_({
      bytes: data.signature.split(',')[1],
      mimeType: 'image/png',
      filename: 'TTD_' + str_(data.serialNumber) + '_' + Date.now() + '.png'
    }, getFolder_(CONFIG.FOLDER_TTD));
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getDataSheet_();
    ensureHeaders_(sheet);

    let row = null;
    let old = null;
    if (isEdit) {
      row = parseInt(data.row, 10);
      if (!(row >= 2 && row <= sheet.getLastRow())) {
        throw new Error('Baris data tidak ditemukan. Klik Refresh lalu coba lagi.');
      }
      old = sheet.getRange(row, 1, 1, NUM_COLS).getValues()[0];
      if (str_(old[3]) !== str_(data.originalSn)) {
        throw new Error('Data di baris ini sudah berubah (mungkin ada yang menghapus data lain). Klik Refresh lalu coba lagi.');
      }
    }

    const fotoLama = old ? str_(old[10]).split(',').map(s => s.trim()).filter(Boolean) : [];
    const foto = fotoLama.concat(fotoBaru);
    const noReport = (old && str_(old[12])) || generateNoReport_(tanggal);
    const ttd = ttdBaru || (old ? str_(old[13]) : '');

    const values = [
      tanggal,
      sanitize_(data.customer),
      sanitize_(data.tipeMesin),
      sanitize_(data.serialNumber),
      sanitize_(data.kategori),
      sanitize_(data.teknisi),
      sanitize_(data.problem),
      sanitize_(data.solusi),
      sanitize_(data.statusMesin),
      sanitize_(data.statusPengerjaan),
      foto.join(','),
      sanitize_(data.sparepart) || '-',
      noReport,
      ttd,
      sanitize_(data.penerima)
    ];

    if (isEdit) {
      sheet.getRange(row, 1, 1, NUM_COLS).setValues([values]);
    } else {
      sheet.appendRow(values);
      row = sheet.getLastRow();
    }
    sheet.getRange(row, 1).setNumberFormat('dd/mm/yyyy');

    autoSaveCustomerMaster_(data.customer, data.tipeMesin, data.serialNumber);
    SpreadsheetApp.flush();
    return { row: row, noReport: noReport };
  } finally {
    lock.releaseLock();
  }
}

function deleteRecord_(row, expectedSn, pwd) {
  checkPassword_(pwd);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getDataSheet_();
    row = parseInt(row, 10);
    if (!(row >= 2 && row <= sheet.getLastRow())) {
      throw new Error('Baris data tidak ditemukan. Klik Refresh lalu coba lagi.');
    }
    const sn = str_(sheet.getRange(row, 4).getValue());
    if (sn !== str_(expectedSn)) {
      throw new Error('Data di baris ini sudah berubah. Klik Refresh lalu coba lagi.');
    }
    sheet.deleteRow(row);
    return true;
  } finally {
    lock.releaseLock();
  }
}

function autoSaveCustomerMaster_(customer, tipeMesin, serialNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_CUSTOMER);
  const sn = str_(serialNumber);
  if (!sheet || !sn) return;

  const lastRow = sheet.getLastRow();
  const existing = lastRow > 1 ? sheet.getRange(2, 3, lastRow - 1, 1).getValues() : [];
  const ada = existing.some(r => str_(r[0]).toLowerCase() === sn.toLowerCase());
  if (!ada) sheet.appendRow([str_(customer), str_(tipeMesin), sn]);
}

// ======================= UTILITAS SEKALI JALAN =======================

/**
 * Jalankan SEKALI dari editor (pilih fungsi ini lalu klik Run).
 * Mengubah tanggal lama yang tersimpan sebagai teks ("05 Mei 2026") menjadi tanggal asli.
 */
function migrasiTanggalLama() {
  const sheet = getDataSheet_();
  ensureHeaders_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const range = sheet.getRange(2, 1, lastRow - 1, 1);
  const vals = range.getValues();
  let diubah = 0;
  const gagal = [];

  const baru = vals.map((r, i) => {
    if (r[0] instanceof Date || str_(r[0]) === '') return [r[0]];
    const d = parseTanggal_(r[0]);
    if (d) { diubah++; return [d]; }
    gagal.push(i + 2);
    return [r[0]];
  });

  range.setValues(baru);
  range.setNumberFormat('dd/mm/yyyy');
  console.log('Tanggal diubah: ' + diubah + ' baris.' +
    (gagal.length ? ' Tidak terbaca di baris: ' + gagal.join(', ') : ''));
}
