/**
 * PILKETOS SMP - Backend Google Apps Script (versi PIN)
 *
 * CARA PAKAI SINGKAT
 *  1. Tempel semua kode ini di Extensions > Apps Script pada Google Sheet kosong.
 *  2. Pilih fungsi  siapkanSimulasi  di dropdown, klik Run, izinkan akses.
 *     -> otomatis membuat tab Pemilih (90 siswa simulasi + guru, lengkap dengan PIN), Suara, dan Rekap.
 *  3. Menu Pilketos > Buat kartu PIN (cetak). Bagikan PIN ke tiap pemilih, lalu HAPUS tab "Kartu PIN".
 *  4. Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
 *     Setiap kode diubah: Deploy > Manage deployments > edit > Version: New version.
 *  5. Salin URL /exec, tempel di index.html (baris const API='').
 *
 * Sheet "Pemilih": A=NISN siswa (10 angka) / NIP guru (18 angka), B=Nama, C=Kelas (guru: "Guru"),
 *                  D=Status (otomatis: SUDAH), E=PIN (6 angka, otomatis)
 * Sheet "Suara"  : A=Waktu, B=Paslon OSIS (1-3), C=Ketua MPK (1-3). TANPA NISN, suara rahasia
 * Sheet "Rekap"  : hitungan otomatis untuk panitia
 * API: semua lewat POST (login, vote, results). PIN dan kode panitia tidak muncul di URL.
 */

// ===== PENGATURAN =====
const JUMLAH_PASLON    = 3;             // paslon Ketua & Wakil OSIS (harus sama dengan jumlah kandidat di index.html)
const JUMLAH_MPK       = 3;             // calon Ketua MPK (harus sama dengan jumlah kandidat di index.html)
const KUNCI_PANITIA    = 'panitia-2026'; // GANTI sebelum pemilihan sungguhan!
const PEMILIHAN_DIBUKA = true;           // false = login & voting ditolak
const HASIL_TERBUKA    = false;          // true = siswa boleh lihat hasil
const JUMLAH_SIMULASI  = 90;             // jumlah siswa simulasi
const MAKS_GAGAL       = 5;              // salah PIN/kode berturut-turut sebelum dikunci
const KUNCI_MENIT      = 10;             // lama kunci (menit)

// Guru: [NIP 18 angka, Nama]. Urutan bebas. Jalankan tambahGuru setelah mengubah daftar ini.
const DAFTAR_GURU = [
  ['197506042005012011', 'Nurhidayah, S.Pd.'],
  ['196610091994032006', 'Dra. Mahmudah'],
  ['196704101990032011', 'Lis Setyaningsih, S.Pd.Fis.'],
  ['197506101999032007', 'Titik Wijayanti, M.Pd.'],
  ['196710011994121003', 'Supomo, S.Pd'],
  ['197202152007012008', 'Rochmatun, S.Pd.'],
  ['197903232008012015', 'Khamidah, S.Pd.'],
  ['197906052010011022', 'Moch. Solehfudin, S.Kom'],
  ['198110012010012018', 'Dwi Kartika Rahmawati, S.Pd.'],
  ['197805042008011012', 'Benny Eko Listanto, S.Pd.'],
  ['197212192007012004', 'Udiartik, S.Pd.'],
  ['198309272014062007', 'Peny Lukito Sari, S.Pd.'],
  ['198602272020121006', 'Edy Kurniawan, S.Pd.'],
  ['199102072020122018', 'Rina Wijayanti, S.Pd.'],
  ['199612162022211007', 'Isyahida Fatwa Gusti, S.Pd'],
  ['199602292022212005', 'Fitronatul Mardhiyah, S.Pd'],
  ['199309152022212006', 'Indah Dwi Astuti, S.Pd.'],
  ['198807302023211006', 'Abdul Muis, S.Pd.'],
  ['198903092023212021', 'Ratna Hadi Astuti, S.Pd.'],
  ['199812312024212012', 'Diah Hesti Safitri, S.Pd.'],
  ['199907142024211002', 'Wildan Azza Assegaf, S.Pd.'],
  ['199504042024212016', 'Nur Laili Mudrikah, S.Pd.'],
  ['199903202024211005', 'Izal Nur Widagdo, S.Pd.'],
  ['199812242024212007', 'Elvina Murni Alsuci, S.Pd.'],
  ['199603072024212008', 'Diyah Arum Puspitasari, S.Pd.'],
  ['199710032024212008', 'Tri Oktaviyani, S.Pd'],
  ['199707272024212017', 'Wiwik Nur Laili, S.Pd.'],
  ['199805222027211001', 'Yusril Ainul Yakin, S.Pd.'],
  ['197307172025212017', 'Imro\'ah'],
  ['197605292025212007', 'Rochmatun'],
  ['199207122025212113', 'Himatul Ulya, S.E'],
  ['198302142025211071', 'Sudrajad']
];
// ======================

/* ---------- MENU & SETUP ---------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Pilketos')
    .addItem('Siapkan simulasi (90 siswa + guru)', 'siapkanSimulasi')
    .addItem('Tambah daftar guru (NIP)', 'tambahGuru')
    .addItem('Buat PIN pemilih (yang belum punya)', 'buatPin')
    .addItem('Buat kartu PIN (cetak)', 'buatKartuPin')
    .addItem('Kosongkan suara & status', 'kosongkanSuaraDanStatus')
    .addToUi();
}

/**
 * Membuat semua tab dan data simulasi. Jika sheet Suara sudah berisi suara,
 * fungsi ini meminta konfirmasi karena semua suara dan status akan dihapus.
 */
function siapkanSimulasi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s0 = ss.getSheetByName('Suara');
  if (s0 && s0.getLastRow() > 1) {
    const ui = SpreadsheetApp.getUi();
    const jawab = ui.alert('Suara sudah ada',
      'Sheet Suara berisi ' + (s0.getLastRow() - 1) + ' suara. Simulasi akan MENGHAPUS semua suara dan status. Lanjutkan?',
      ui.ButtonSet.YES_NO);
    if (jawab !== ui.Button.YES) { Logger.log('Dibatalkan. Suara tidak diubah.'); return; }
  }
  const p = ensure_(ss, 'Pemilih', ['NISN', 'Nama', 'Kelas', 'Status', 'PIN']);
  const s = ensure_(ss, 'Suara', ['Waktu', 'Paslon OSIS', 'Ketua MPK']);
  s.getRange(1, 1, 1, 3).setValues([['Waktu', 'Paslon OSIS', 'Ketua MPK']]);
  p.getRange('A:A').setNumberFormat('@'); // NISN sebagai teks (angka 0 depan aman)
  siapkanKolomPin_(p);

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
  buatPin_(p);
  kosongkanSuaraDanStatus();
  buatRekap_(ss);

  // hapus tab kosong bawaan
  ss.getSheets().forEach(function (sh) {
    if (['Sheet1', 'Lembar1'].indexOf(sh.getName()) >= 0 && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  Logger.log('Siap! NISN simulasi 1000000001 s/d ' + (1000000000 + JUMLAH_SIMULASI) + '. PIN ada di kolom E sheet Pemilih. Kode panitia: ' + KUNCI_PANITIA);
  Logger.log('Langkah berikut: Deploy > New deployment > Web app. Setelah itu jalankan tampilkanUrl untuk melihat URL.');
}

/**
 * Menambah guru (NIP) ke sheet Pemilih tanpa menghapus suara. Aman dijalankan ulang:
 * NIP yang sudah ada dilewati, nama placeholder lama ("Guru 01") diganti nama asli.
 * Guru baru otomatis mendapat PIN.
 */
function tambahGuru() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const p = ensure_(ss, 'Pemilih', ['NISN', 'Nama', 'Kelas', 'Status', 'PIN']);
  p.getRange('A:A').setNumberFormat('@');
  const baru = tambahGuru_(ss, p);
  const pinBaru = buatPin_(p);
  buatRekap_(ss);
  Logger.log(baru + ' guru ditambahkan, ' + pinBaru + ' PIN dibuat. Total NIP terdaftar: ' + DAFTAR_GURU.length + '.');
}

function tambahGuru_(ss, p) {
  const ada = {};
  const n = p.getLastRow() - 1;
  if (n > 0) p.getRange(2, 1, n, 2).getValues().forEach(function (r, i) { ada[nisn_(r[0])] = { row: i + 2, nama: String(r[1]) }; });
  const rows = [];
  DAFTAR_GURU.forEach(function (g) {
    const x = ada[nisn_(g[0])];
    if (!x) rows.push([g[0], g[1], 'Guru', '']);
    else if (/^Guru \d+$/.test(x.nama)) p.getRange(x.row, 2).setValue(g[1]); // ganti nama placeholder lama
  });
  if (rows.length) p.getRange(p.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  return rows.length;
}

/* ---------- PIN ---------- */

function siapkanKolomPin_(p) {
  p.getRange(1, 5).setValue('PIN');
  p.getRange('E:E').setNumberFormat('@'); // PIN sebagai teks (angka 0 depan aman)
}

// PIN 6 angka acak dari UUID
function pin_() {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid());
  let n = 0;
  for (let i = 0; i < 4; i++) n = n * 256 + (b[i] & 255);
  return ('000000' + (n % 1000000)).slice(-6);
}

/** Isi PIN untuk baris yang belum punya. PIN yang sudah ada tidak berubah. Return jumlah PIN baru. */
function buatPin_(p) {
  siapkanKolomPin_(p);
  const n = p.getLastRow() - 1;
  if (n < 1) return 0;
  const all = p.getRange(2, 1, n, 5).getValues();
  const pakai = {};
  all.forEach(function (r) { const v = String(r[4]).trim(); if (v) pakai[v] = true; });
  let baru = 0;
  const out = all.map(function (r) {
    const ada = String(r[4]).trim();
    if (ada || !String(r[0]).trim()) return [ada];
    let x;
    do { x = pin_(); } while (pakai[x]);
    pakai[x] = true; baru++;
    return [x];
  });
  if (baru) p.getRange(2, 5, n, 1).setValues(out);
  return baru;
}

function buatPin() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const p = ensure_(ss, 'Pemilih', ['NISN', 'Nama', 'Kelas', 'Status', 'PIN']);
  const baru = buatPin_(p);
  Logger.log(baru + ' PIN baru dibuat.');
  ss.toast(baru + ' PIN baru dibuat.', 'Pilketos');
}

/** Membuat tab "Kartu PIN" (urut kelas lalu nama) untuk dicetak dan dipotong. HAPUS tab ini setelah dicetak. */
function buatKartuPin() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const p = ss.getSheetByName('Pemilih');
  if (!p || p.getLastRow() < 2) { Logger.log('Sheet Pemilih kosong.'); return; }
  buatPin_(p);
  const v = p.getRange(2, 1, p.getLastRow() - 1, 5).getValues().filter(function (r) { return String(r[0]).trim(); });
  v.sort(function (a, b) {
    return String(a[2]).localeCompare(String(b[2])) || String(a[1]).localeCompare(String(b[1]));
  });
  const k = ss.getSheetByName('Kartu PIN') || ss.insertSheet('Kartu PIN');
  k.clear();
  k.getRange('A:D').setNumberFormat('@');
  const rows = [['Kelas', 'Nama', 'NISN / NIP', 'PIN']].concat(v.map(function (r) {
    return [String(r[2]), String(r[1]), String(r[0]), String(r[4])];
  }));
  k.getRange(1, 1, rows.length, 4).setValues(rows);
  k.getRange(1, 1, 1, 4).setFontWeight('bold');
  k.setFrozenRows(1);
  ss.setActiveSheet(k);
  Logger.log('Kartu PIN dibuat (' + v.length + ' pemilih). Cetak, lalu HAPUS tab "Kartu PIN".');
}

/** Reset: hapus semua suara dan kosongkan kolom Status. PIN tidak berubah. */
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

// Batas percobaan salah (disimpan di cache, hilang sendiri setelah KUNCI_MENIT)
function gagal_(id) { return Number(CacheService.getScriptCache().get('gagal:' + id) || 0); }
function catatGagal_(id) { CacheService.getScriptCache().put('gagal:' + id, String(gagal_(id) + 1), KUNCI_MENIT * 60); }
function hapusGagal_(id) { CacheService.getScriptCache().remove('gagal:' + id); }

// Semua aksi lewat POST: PIN dan kode panitia tidak ikut di URL.
function doGet() {
  return json_({ ok: true, msg: 'Pilketos API aktif' });
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const pin = String(d.pin == null ? '' : d.pin).replace(/\D/g, '');
    if (d.action === 'login')   return json_(login_(nisn_(d.nisn), pin));
    if (d.action === 'vote')    return json_(vote_(nisn_(d.nisn), pin, Number(d.paslon), Number(d.mpk)));
    if (d.action === 'results') return json_(results_(String(d.key || '')));
  } catch (err) {}
  return json_({ ok: false, error: 'SERVER' });
}

// cari baris pemilih; return {row, data} atau null
function find_(sheet, nisn) {
  const n = sheet.getLastRow() - 1;
  if (n < 1) return null;
  const vals = sheet.getRange(2, 1, n, 5).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (nisn_(vals[i][0]) === nisn) return { row: i + 2, data: vals[i] };
  }
  return null;
}

// cek NISN/NIP + PIN. Return {error} atau {f}. Nomor tak terdaftar dan PIN salah memberi error sama.
function cek_(sheet, nisn, pin) {
  if (gagal_(nisn) >= MAKS_GAGAL) return { error: 'TERKUNCI' };
  const f = find_(sheet, nisn);
  const benar = f && /^\d{6}$/.test(pin) && String(f.data[4]).trim() === pin;
  if (!benar) { catatGagal_(nisn); return { error: 'LOGIN_GAGAL' }; }
  hapusGagal_(nisn);
  return { f: f };
}

function login_(nisn, pin) {
  if (!PEMILIHAN_DIBUKA) return { ok: false, error: 'DITUTUP' };
  const c = cek_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pemilih'), nisn, pin);
  if (c.error) return { ok: false, error: c.error };
  const d = c.f.data;
  return { ok: true, nama: String(d[1]), kelas: String(d[2]), voted: String(d[3]).toUpperCase() === 'SUDAH', paslon: JUMLAH_PASLON, mpk: JUMLAH_MPK };
}

function vote_(nisn, pin, paslon, mpk) {
  if (!PEMILIHAN_DIBUKA) return { ok: false, error: 'DITUTUP' };
  if (!(paslon >= 1 && paslon <= JUMLAH_PASLON) || paslon % 1 !== 0) return { ok: false, error: 'PASLON_INVALID' };
  if (!(mpk >= 1 && mpk <= JUMLAH_MPK) || mpk % 1 !== 0) return { ok: false, error: 'MPK_INVALID' };
  const lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (err) { return { ok: false, error: 'SIBUK' }; }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Pemilih');
    const c = cek_(sh, nisn, pin);
    if (c.error) return { ok: false, error: c.error };
    const f = c.f;
    if (String(f.data[3]).toUpperCase() === 'SUDAH') return { ok: false, error: 'SUDAH_MEMILIH' };
    // tandai pemilih dulu, baru tulis suara (tanpa NISN). Jika gagal menulis, tanda dibatalkan.
    sh.getRange(f.row, 4).setValue('SUDAH');
    try {
      ss.getSheetByName('Suara').appendRow([new Date(), paslon, mpk]);
      SpreadsheetApp.flush();
    } catch (err) {
      sh.getRange(f.row, 4).clearContent();
      return { ok: false, error: 'SERVER' };
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function results_(key) {
  if (!HASIL_TERBUKA) {
    if (key === '') return { ok: true, locked: true };
    if (gagal_('panitia') >= MAKS_GAGAL) return { ok: false, error: 'TERKUNCI' };
    if (key !== KUNCI_PANITIA) { catatGagal_('panitia'); return { ok: true, locked: true }; }
    hapusGagal_('panitia');
  }
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
