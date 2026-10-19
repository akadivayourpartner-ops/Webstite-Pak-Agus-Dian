/**
 * PORTAL GURU WALI - BACKEND API (Google Apps Script + Google Sheets)
 * Frontend (index.html) di-hosting di GitHub Pages dan memanggil URL Web App ini.
 * Jalankan setupPortalDatabase() sekali untuk membuat semua sheet.
 */

var SHEET_NAMES = {
  CONFIG: 'Config',
  SISWA: 'Master Siswa',
  PRESENSI: 'Presensi',
  HAFALAN: 'Hafalan',
  KEGIATAN: 'Kegiatan',
  CURHAT: 'Curhat'
};

var HEADERS = {
  'Config': ['Kunci', 'Nilai', 'Keterangan'],
  'Master Siswa': ['No', 'Nama Siswa', 'Kelas'],
  'Presensi': ['Bulan', 'Pertemuan', 'Nama Siswa', 'Kelas', 'Status', 'Waktu Input'],
  'Hafalan': ['Nama Siswa', 'Kelas', 'Juz 30', 'Bacaan Shalat', 'Terakhir Update'],
  'Kegiatan': ['Tanggal', 'Topik / Nama Kegiatan', 'Keterangan', 'Waktu Input'],
  'Curhat': ['Waktu', 'Nama Siswa', 'Kelas', 'Pesan']
};

var CONFIG_DEFAULT = [
  ['NamaGuru', 'Ibu Lucky Salilana', 'Nama Guru Wali'],
  ['NamaKelompok', 'Siswa Binaan Ibu Lucky Salilana', 'Nama kelompok binaan'],
  ['Sekolah', 'SMP Unggul Bunga Bangsa', 'Satuan pendidikan'],
  ['TahunAjaran', '2026/2027', 'Periode tahun ajaran'],
  ['Password', 'unggulbungabangsa', 'Password akses Kotak Curhat (khusus Guru Wali)'],
  ['BatasHadir', '80', 'Batas minimal % hadir untuk status "Baik"']
];

var SISWA_DEFAULT = [
  [1, 'Siswa Contoh 1', 'VIII.A'],
  [2, 'Siswa Contoh 2', 'VIII.B'],
  [3, 'Siswa Contoh 3', 'VIII.C'],
  [4, 'Siswa Contoh 4', 'IX.A'],
  [5, 'Siswa Contoh 5', 'IX.B']
];

var JUZ30 = [
  'An-Nas', 'Al-Falaq', 'Al-Ikhlas', 'Al-Lahab', 'An-Nasr', 'Al-Kafirun', 'Al-Kautsar',
  "Al-Ma'un", 'Quraisy', 'Al-Fil', 'Al-Humazah', "Al-'Asr", 'At-Takatsur', "Al-Qari'ah",
  "Al-'Adiyat", 'Az-Zalzalah', 'Al-Bayyinah', 'Al-Qadr', "Al-'Alaq", 'At-Tin', 'Al-Insyirah',
  'Ad-Dhuha', 'Al-Lail', 'Asy-Syams', 'Al-Balad', 'Al-Fajr', 'Al-Ghasyiyah', "Al-A'la",
  'At-Thariq', 'Al-Buruj', 'Al-Insyiqaq', 'Al-Muthaffifin', 'Al-Infithar', 'At-Takwir',
  "'Abasa", "An-Nazi'at", "An-Naba'"
];

var SHALAT = [
  'Niat Shalat', 'Doa Iftitah', "Bacaan Ruku'", "Bacaan I'tidal", 'Bacaan Sujud',
  'Duduk di Antara Dua Sujud', 'Tasyahud (Awal & Akhir)', 'Doa Qunut', 'Dzikir Setelah Shalat'
];

var BULAN = ['Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember',
  'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni'];

/* ======================= SETUP DATABASE ======================= */

function setupPortalDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var names = Object.keys(HEADERS);

  names.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);

    if (sh.getLastRow() === 0) {
      var h = HEADERS[name];
      sh.getRange(1, 1, 1, h.length).setValues([h])
        .setFontWeight('bold').setBackground('#1d4ed8').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }

    if (name === 'Config' && sh.getLastRow() === 1) {
      sh.getRange(2, 1, CONFIG_DEFAULT.length, 3).setValues(CONFIG_DEFAULT);
    }
    if (name === 'Master Siswa' && sh.getLastRow() === 1) {
      sh.getRange(2, 1, SISWA_DEFAULT.length, 3).setValues(SISWA_DEFAULT);
    }
    if (name === 'Kegiatan') sh.getRange('A:A').setNumberFormat('@');
    sh.autoResizeColumns(1, HEADERS[name].length);
  });

  var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Sheet 1') || ss.getSheetByName('Lembar 1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);

  SpreadsheetApp.getUi().alert('Database siap. Sheet dibuat: ' + names.join(', ') +
    '.\n\nEdit sheet "Config" dan "Master Siswa", lalu Deploy sebagai Web App.');
}

/* ======================= HELPER ======================= */

function sheet_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" belum ada. Jalankan setupPortalDatabase() dahulu.');
  return sh;
}

function rows_(name) {
  var sh = sheet_(name);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var n = HEADERS[name].length;
  return sh.getRange(2, 1, last - 1, n).getValues().filter(function (r) {
    return r.join('') !== '';
  });
}

function str_(v, pattern) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), pattern || 'yyyy-MM-dd');
  }
  return String(v === null || v === undefined ? '' : v);
}

function now_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function config_() {
  var cfg = {};
  rows_('Config').forEach(function (r) { cfg[String(r[0]).trim()] = String(r[1]); });
  return cfg;
}

function siswa_() {
  return rows_('Master Siswa')
    .filter(function (r) { return String(r[1]).trim() !== ''; })
    .map(function (r) { return { nama: String(r[1]).trim(), kelas: String(r[2]).trim() }; });
}

function splitList_(v) {
  return String(v || '').split(';').map(function (x) { return x.trim(); }).filter(Boolean);
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ======================= WEB APP ======================= */

var API_FUNCTIONS = {
  getInitData: getInitData, getPresensi: getPresensi, savePresensi: savePresensi,
  getHafalan: getHafalan, saveHafalan: saveHafalan, getRekapHafalan: getRekapHafalan,
  getRekapAbsen: getRekapAbsen, getKegiatan: getKegiatan, saveKegiatan: saveKegiatan,
  saveCurhat: saveCurhat, getCurhat: getCurhat
};

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return json_({ ok: true, data: 'API Portal Guru Wali aktif.' });
}

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    var fn = API_FUNCTIONS[req.fn];
    if (!fn) throw new Error('Fungsi tidak dikenal.');
    var out = fn.apply(null, req.args || []);
    return json_({ ok: true, data: out === undefined ? null : out });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function getInitData() {
  var cfg = config_();
  return {
    guru: cfg.NamaGuru || '',
    kelompok: cfg.NamaKelompok || '',
    sekolah: cfg.Sekolah || '',
    tahunAjaran: cfg.TahunAjaran || '',
    siswa: siswa_(),
    juz30: JUZ30,
    shalat: SHALAT,
    bulan: BULAN
  };
}

/* ---------- Presensi ---------- */

function getPresensi(bulan, sesi) {
  var map = {};
  rows_('Presensi').forEach(function (r) {
    if (String(r[0]) === String(bulan) && String(r[1]) === String(sesi)) {
      map[String(r[2])] = String(r[4]);
    }
  });
  return map;
}

function savePresensi(bulan, sesi, records) {
  return withLock_(function () {
    var sh = sheet_('Presensi');
    var last = sh.getLastRow();
    var keep = [];
    if (last > 1) {
      keep = sh.getRange(2, 1, last - 1, 6).getValues().filter(function (r) {
        return !(String(r[0]) === String(bulan) && String(r[1]) === String(sesi)) && r.join('') !== '';
      });
      sh.getRange(2, 1, last - 1, 6).clearContent();
    }
    var kelasMap = {};
    siswa_().forEach(function (s) { kelasMap[s.nama] = s.kelas; });
    var valid = { H: 1, S: 1, I: 1, A: 1, C: 1 };
    var waktu = now_();
    var saved = 0;
    records.forEach(function (r) {
      if (kelasMap[r.nama] !== undefined && valid[r.status]) {
        keep.push([String(bulan), Number(sesi), r.nama, kelasMap[r.nama], r.status, waktu]);
        saved++;
      }
    });
    if (keep.length) sh.getRange(2, 1, keep.length, 6).setValues(keep);
    return saved;
  });
}

/* ---------- Hafalan ---------- */

function findHafalanRow_(sh, nama) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var names = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < names.length; i++) {
    if (String(names[i][0]) === nama) return i + 2;
  }
  return 0;
}

function getHafalan(nama) {
  var res = { juz: [], shalat: [] };
  rows_('Hafalan').forEach(function (r) {
    if (String(r[0]) === String(nama)) {
      res.juz = splitList_(r[2]);
      res.shalat = splitList_(r[3]);
    }
  });
  return res;
}

function saveHafalan(nama, juz, shalat) {
  return withLock_(function () {
    var kelas = null;
    siswa_().forEach(function (s) { if (s.nama === nama) kelas = s.kelas; });
    if (kelas === null) throw new Error('Siswa tidak ditemukan di Master Siswa.');
    var juzOk = juz.filter(function (x) { return JUZ30.indexOf(x) > -1; });
    var shalatOk = shalat.filter(function (x) { return SHALAT.indexOf(x) > -1; });
    var sh = sheet_('Hafalan');
    var row = findHafalanRow_(sh, nama);
    var data = [[nama, kelas, juzOk.join(';'), shalatOk.join(';'), now_()]];
    if (row) {
      sh.getRange(row, 1, 1, 5).setValues(data);
    } else {
      sh.getRange(sh.getLastRow() + 1, 1, 1, 5).setValues(data);
    }
    return { juz: juzOk.length, shalat: shalatOk.length };
  });
}

function getRekapHafalan() {
  var map = {};
  rows_('Hafalan').forEach(function (r) {
    map[String(r[0])] = { juz: splitList_(r[2]), shalat: splitList_(r[3]) };
  });
  return siswa_().map(function (s) {
    var d = map[s.nama] || { juz: [], shalat: [] };
    return {
      nama: s.nama, kelas: s.kelas,
      juz: d.juz, shalat: d.shalat,
      totalJuz: JUZ30.length, totalShalat: SHALAT.length
    };
  });
}

/* ---------- Rekap Absen ---------- */

function getRekapAbsen() {
  var cfg = config_();
  var batas = Number(cfg.BatasHadir) || 80;
  var count = {};
  rows_('Presensi').forEach(function (r) {
    var n = String(r[2]);
    var st = String(r[4]);
    if (!count[n]) count[n] = { H: 0, S: 0, I: 0, A: 0, C: 0 };
    if (count[n][st] !== undefined) count[n][st]++;
  });
  return siswa_().map(function (s) {
    var c = count[s.nama] || { H: 0, S: 0, I: 0, A: 0, C: 0 };
    var total = c.H + c.S + c.I + c.A + c.C;
    var persen = total > 0 ? Math.round((c.H / total) * 1000) / 10 : 0;
    var status = total === 0 ? 'Belum ada data' : (persen >= batas ? 'Baik' : 'Perlu Perhatian');
    return { nama: s.nama, kelas: s.kelas, H: c.H, S: c.S, I: c.I, A: c.A, C: c.C, persen: persen, status: status };
  });
}

/* ---------- Kegiatan ---------- */

function getKegiatan() {
  return rows_('Kegiatan').map(function (r) {
    return { tanggal: str_(r[0]), topik: str_(r[1]), ket: str_(r[2]) };
  }).sort(function (a, b) { return a.tanggal < b.tanggal ? 1 : (a.tanggal > b.tanggal ? -1 : 0); });
}

function saveKegiatan(tanggal, topik, ket) {
  topik = String(topik || '').trim();
  if (!tanggal || !topik) throw new Error('Tanggal dan Topik wajib diisi.');
  return withLock_(function () {
    var sh = sheet_('Kegiatan');
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange(sh.getLastRow() + 1, 1, 1, 4).setValues([[
      String(tanggal), topik.substring(0, 200), String(ket || '').substring(0, 1000), now_()
    ]]);
    return true;
  });
}

/* ---------- Curhat ---------- */

function saveCurhat(nama, pesan) {
  pesan = String(pesan || '').trim();
  if (!nama) throw new Error('Pilih nama siswa terlebih dahulu.');
  if (!pesan) throw new Error('Pesan tidak boleh kosong.');
  return withLock_(function () {
    var kelas = null;
    siswa_().forEach(function (s) { if (s.nama === nama) kelas = s.kelas; });
    if (kelas === null) throw new Error('Siswa tidak ditemukan di Master Siswa.');
    var sh = sheet_('Curhat');
    sh.getRange(sh.getLastRow() + 1, 1, 1, 4).setValues([[now_(), nama, kelas, pesan.substring(0, 2000)]]);
    return true;
  });
}

function getCurhat(password) {
  var cfg = config_();
  if (!cfg.Password || String(password) !== cfg.Password) {
    throw new Error('Password salah.');
  }
  return rows_('Curhat').map(function (r) {
    return { waktu: str_(r[0], 'yyyy-MM-dd HH:mm'), nama: str_(r[1]), kelas: str_(r[2]), pesan: str_(r[3]) };
  }).reverse();
}
