/**
 * PILKETOS SMP - Backend Google Apps Script (versi siap-run)
 *
 * CARA PAKAI SINGKAT
 *  1. Tempel semua kode ini di Extensions > Apps Script pada Google Sheet kosong.
 *  2. Pilih fungsi  siapkanSimulasi  di dropdown, klik Run, izinkan akses.
 *     -> otomatis membuat tab Pemilih (90 siswa simulasi), Suara, dan Rekap.
 *  3. Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
 *  4. Salin URL /exec, tempel di pilketos-smp.html (baris const API='').
 *
 * Sheet "Pemilih": A=NISN siswa (10 angka) / NIP guru (18 angka), B=Nama, C=Kelas (guru: "Guru"), D=Status (otomatis: SUDAH)
 * Sheet "Suara"  : A=Waktu, B=Paslon OSIS (1-3), C=Ketua MPK (1-3). TANPA NISN, suara rahasia
 * Sheet "Rekap"  : hitungan otomatis untuk panitia
 */

// ===== PENGATURAN =====
const JUMLAH_PASLON    = 3;             // paslon Ketua & Wakil OSIS
const JUMLAH_MPK       = 3;             // calon Ketua MPK
const KUNCI_PANITIA    = 'panitia-2026'; // GANTI sebelum pemilihan sungguhan!
const PEMILIHAN_DIBUKA = true;           // false = login & voting ditolak
const HASIL_TERBUKA    = false;          // true = siswa boleh lihat hasil
const JUMLAH_SIMULASI  = 90;             // jumlah siswa simulasi

// NIP guru (18 angka). Nama di sheet Pemilih bisa diganti manual.
const DAFTAR_NIP = [
  '197506042005012011',
  '196610091994032006',
  '196704101990032011',
  '197506101999032007',
  '196710011994121003',
  '197202152007012008',
  '197903232008012015',
  '197906052010011022',
  '198110012010012018',
  '197805042008011012',
  '197212192007012004',
  '198309272014062007',
  '198602272020121006',
  '199102072020122018',
  '199612162022211007',
  '199602292022212005',
  '199309152022212006',
  '198807302023211006',
  '198903092023212021',
  '199812312024212012',
  '199907142024211002',
  '199504042024212016',
  '199903202024211005',
  '199812242024212007',
  '199603072024212008',
  '199710032024212008',
  '199707272024212017',
  '199805222027211001',
  '197307172025212017',
  '197605292025212007',
  '199207122025212113',
  '198302142025211071'
];
// ======================

/* ---------- MENU & SETUP ---------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Pilketos')
    .addItem('Siapkan simulasi (90 siswa + guru)', 'siapkanSimulasi')
    .addItem('Tambah daftar guru (NIP)', 'tambahGuru')
    .addItem('Kosongkan suara & status', 'kosongkanSuaraDanStatus')
    .addToUi();
}

/** Jalankan SEKALI: membuat semua tab dan data simulasi. Aman dijalankan ulang. */
function siapkanSimulasi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const p = ensure_(ss, 'Pemilih', ['NISN', 'Nama', 'Kelas', 'Status']);
  const s = ensure_(ss, 'Suara', ['Waktu', 'Paslon OSIS', 'Ketua MPK']);
  s.getRange(1, 1, 1, 3).setValues([['Waktu', 'Paslon OSIS', 'Ketua MPK']]);
  p.getRange('A:A').setNumberFormat('@'); // NISN sebagai teks (angka 0 depan aman)

  // isi pemilih simulasi hanya kalau masih kosong
  if (p.getLastRow() < 2) {
    const kelas = ['7A', '7B', '7C'];
    const rows = [];
    for (let i = 1; i <= JUMLAH_SIMULASI; i++) {
      rows.push([String(1000000000 + i), 'Siswa ' + ('0' + i).slice(-2), kelas[(i - 1) % kelas.length], '']);
    }
    p.getRange(2, 1, rows.length, 4).setValues(rows);
  }
  tambahGuru_(ss, p);
  kosongkanSuaraDanStatus();
  buatRekap_(ss);

  // hapus tab kosong bawaan
  ss.getSheets().forEach(function (sh) {
    if (['Sheet1', 'Lembar1'].indexOf(sh.getName()) >= 0 && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  Logger.log('Siap! NISN simulasi 1000000001 s/d ' + (1000000000 + JUMLAH_SIMULASI) + '. Kode panitia: ' + KUNCI_PANITIA);
  Logger.log('Langkah berikut: Deploy > New deployment > Web app. Setelah itu jalankan tampilkanUrl untuk melihat URL.');
}

/**
 * Menambah guru (NIP) ke sheet Pemilih tanpa menghapus suara. Aman dijalankan ulang:
 * NIP yang sudah ada dilewati. Sesudah itu deploy versi baru Web App.
 */
function tambahGuru() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const p = ensure_(ss, 'Pemilih', ['NISN', 'Nama', 'Kelas', 'Status']);
  p.getRange('A:A').setNumberFormat('@');
  const baru = tambahGuru_(ss, p);
  buatRekap_(ss);
  Logger.log(baru + ' guru ditambahkan. Total NIP terdaftar: ' + DAFTAR_NIP.length + '.');
}

function tambahGuru_(ss, p) {
  const ada = {};
  const n = p.getLastRow() - 1;
  if (n > 0) p.getRange(2, 1, n, 1).getValues().forEach(function (r) { ada[nisn_(r[0])] = true; });
  const rows = [];
  DAFTAR_NIP.forEach(function (nip, i) {
    if (!ada[nisn_(nip)]) rows.push([nip, 'Guru ' + ('0' + (i + 1)).slice(-2), 'Guru', '']);
  });
  if (rows.length) p.getRange(p.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  return rows.length;
}

/** Reset: hapus semua suara dan kosongkan kolom Status. Dipakai setelah simulasi. */
function kosongkanSuaraDanStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName('Suara');
  if (s && s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1);
  const p = ss.getSheetByName('Pemilih');
  if (p && p.getLastRow() > 1) p.getRange(2, 4, p.getLastRow() - 1, 1).clearContent();
}

/** Cetak URL Web App (setelah di-deploy) ke log. */
function tampilkanUrl() {
  Logger.log(ScriptApp.getService().getUrl() || 'Belum ada deployment Web App.');
}

function ensure_(ss, nama, header) {
  let sh = ss.getSheetByName(nama) || ss.insertSheet(nama);
  if (sh.getLastRow() === 0) sh.appendRow(header);
  sh.setFrozenRows(1);
  return sh;
}

function buatRekap_(ss) {
  const r = ss.getSheetByName('Rekap') || ss.insertSheet('Rekap');
  r.clear();
  r.getRange(1, 1).setValue('REKAP PILKETOS (otomatis)');
  r.getRange(1, 1).setFontWeight('bold');
  let row = 3;
  row = blokRekap_(r, row, 'Ketua & Wakil OSIS', 'Paslon ', JUMLAH_PASLON, 'B');
  row = blokRekap_(r, row + 1, 'Ketua MPK', 'Calon ', JUMLAH_MPK, 'C');
  r.getRange(row + 1, 1).setValue('Jumlah pemilih');
  r.getRange(row + 1, 2).setFormula('=COUNTA(Pemilih!A2:A)');
  r.getRange(row + 2, 1).setValue('Sudah memilih');
  r.getRange(row + 2, 2).setFormula('=COUNTIF(Pemilih!D2:D,"SUDAH")');
  r.getRange(row + 3, 1).setValue('Partisipasi');
  r.getRange(row + 3, 2).setFormula('=IFERROR(B' + (row + 2) + '/B' + (row + 1) + ',0)');
  r.getRange(row + 3, 2).setNumberFormat('0%');
  r.getRange(row + 4, 1).setValue('Pemilih guru');
  r.getRange(row + 4, 2).setFormula('=COUNTIF(Pemilih!C2:C,"Guru")');
  r.getRange(row + 5, 1).setValue('Guru sudah memilih');
  r.getRange(row + 5, 2).setFormula('=COUNTIFS(Pemilih!C2:C,"Guru",Pemilih!D2:D,"SUDAH")');
}

// tulis satu blok rekap; return baris kosong berikutnya
function blokRekap_(r, row, judul, awalan, jumlah, kolSuara) {
  r.getRange(row, 1).setValue(judul);
  r.getRange(row, 2).setValue('Suara');
  r.getRange(row, 1, 1, 2).setFontWeight('bold');
  for (let k = 1; k <= jumlah; k++) {
    r.getRange(row + k, 1).setValue(awalan + k);
    r.getRange(row + k, 2).setFormula('=COUNTIF(Suara!' + kolSuara + '2:' + kolSuara + ',' + k + ')');
  }
  r.getRange(row + jumlah + 1, 1).setValue('Total');
  r.getRange(row + jumlah + 1, 2).setFormula('=SUM(B' + (row + 1) + ':B' + (row + jumlah) + ')');
  return row + jumlah + 2;
}

/* ---------- API WEB APP ---------- */

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
// NISN (<=10 angka) dilengkapi nol di depan jadi 10; NIP (18 angka) dipakai apa adanya
function nisn_(v) {
  const d = String(v == null ? '' : v).trim().replace(/\D/g, '');
  return d.length > 10 ? d : d.padStart(10, '0');
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'login')   return json_(login_(nisn_(p.nisn)));
    if (p.action === 'results') return json_(results_(String(p.key || '')));
    return json_({ ok: true, msg: 'Pilketos API aktif' });
  } catch (err) {
    return json_({ ok: false, error: 'SERVER' });
  }
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === 'vote') return json_(vote_(nisn_(d.nisn), Number(d.paslon), Number(d.mpk)));
  } catch (err) {}
  return json_({ ok: false, error: 'SERVER' });
}

// cari baris pemilih; return {row, data} atau null
function find_(sheet, nisn) {
  const n = sheet.getLastRow() - 1;
  if (n < 1) return null;
  const vals = sheet.getRange(2, 1, n, 4).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (nisn_(vals[i][0]) === nisn) return { row: i + 2, data: vals[i] };
  }
  return null;
}

function login_(nisn) {
  if (!PEMILIHAN_DIBUKA) return { ok: false, error: 'DITUTUP' };
  const f = find_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pemilih'), nisn);
  if (!f) return { ok: false, error: 'TIDAK_TERDAFTAR' };
  return { ok: true, nama: String(f.data[1]), kelas: String(f.data[2]), voted: String(f.data[3]).toUpperCase() === 'SUDAH' };
}

function vote_(nisn, paslon, mpk) {
  if (!PEMILIHAN_DIBUKA) return { ok: false, error: 'DITUTUP' };
  if (!(paslon >= 1 && paslon <= JUMLAH_PASLON) || paslon % 1 !== 0) return { ok: false, error: 'PASLON_INVALID' };
  if (!(mpk >= 1 && mpk <= JUMLAH_MPK) || mpk % 1 !== 0) return { ok: false, error: 'MPK_INVALID' };
  const lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (err) { return { ok: false, error: 'SIBUK' }; }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Pemilih');
    const f = find_(sh, nisn);
    if (!f) return { ok: false, error: 'TIDAK_TERDAFTAR' };
    if (String(f.data[3]).toUpperCase() === 'SUDAH') return { ok: false, error: 'SUDAH_MEMILIH' };
    // tulis suara dulu (tanpa NISN), baru tandai pemilih
    ss.getSheetByName('Suara').appendRow([new Date(), paslon, mpk]);
    sh.getRange(f.row, 4).setValue('SUDAH');
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function results_(key) {
  if (!(HASIL_TERBUKA || key === KUNCI_PANITIA)) return { ok: true, locked: true };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const votes = new Array(JUMLAH_PASLON).fill(0);
  const votesMpk = new Array(JUMLAH_MPK).fill(0);
  const s = ss.getSheetByName('Suara');
  const n = s.getLastRow() - 1;
  if (n > 0) s.getRange(2, 2, n, 2).getValues().forEach(function (r) {
    const k = Number(r[0]), m = Number(r[1]);
    if (k >= 1 && k <= JUMLAH_PASLON) votes[k - 1]++;
    if (m >= 1 && m <= JUMLAH_MPK) votesMpk[m - 1]++;
  });
  const p = ss.getSheetByName('Pemilih');
  const t = p.getLastRow() - 1;
  let sudah = 0;
  if (t > 0) p.getRange(2, 4, t, 1).getValues().forEach(function (r) {
    if (String(r[0]).toUpperCase() === 'SUDAH') sudah++;
  });
  return { ok: true, locked: false, votes: votes, votesMpk: votesMpk, pemilih: Math.max(t, 0), sudah: sudah };
}
