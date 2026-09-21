/*******************************************************************
 * LEGACY ENGINE — pengganti 3 file:
 *   Tarik Data.gs / Tarik Data Remed.gs / Tarik Data Remed2.gs
 * -----------------------------------------------------------------
 * Isi ketiganya identik, cuma beda reportid. Jadi disatuin di sini.
 * Perilakunya sengaja dibikin sama persis kayak aslinya supaya
 * hasil di sheet gak berubah — cuma bug-nya yang dibenerin.
 *
 * SETUP:
 *   1. simpanTokenLegacy()  -> tempel token, terus kosongin lagi
 *   2. HAPUS 3 file lama biar gak bentrok nama fungsi
 *******************************************************************/

const LEGACY = {
  BASE_URL  : 'https://academy.amartha.com/webservice/rest/server.php',
  TOKEN_KEY : 'WSTOKEN_LEGACY',
  SS_ID     : '191zlOveljc2ADVVFPkR8r_OIwdANxed_uUGbDDAdyME'
};

/** Satu-satunya tempat mapping report -> sheet. */
const LEGACY_REPORT = {
  'RAW'        : { reportId: 119, label: 'Reguler'    },
  'Remedial'   : { reportId: 130, label: 'Remedial 1' },
  'Remedial 2' : { reportId: 136, label: 'Remedial 2' }
};


/*******************************************************************
 * TOKEN
 *******************************************************************/

function simpanTokenLegacy() {
  PropertiesService.getScriptProperties()
    .setProperty(LEGACY.TOKEN_KEY, 'TOKEN_LAMA');
  Logger.log('Tersimpan. Sekarang kosongin lagi string di fungsi ini.');
}

function tokenLegacy_() {
  const t = PropertiesService.getScriptProperties().getProperty(LEGACY.TOKEN_KEY);
  if (!t || t.indexOf('TEMPEL') === 0) {
    throw new Error('Token legacy belum di-set. Jalanin simpanTokenLegacy() dulu.');
  }
  return t;
}


/*******************************************************************
 * UTIL
 *******************************************************************/

/** Baris terakhir yang beneran ada datanya, diukur dari kolom A.
 *  Perlu karena kolom M berisi formula sampai jauh ke bawah,
 *  bikin getLastRow() ngasih angka yang menyesatkan. */
function legacyLastDataRow_(sh) {
  const n = sh.getLastRow();
  if (n < 2) return 1;
  const colA = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = colA.length - 1; i >= 0; i--) {
    if (String(colA[i][0]).trim() !== '') return i + 2;
  }
  return 1;
}

function md5_(str) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str, Utilities.Charset.UTF_8)
    .map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}


/*******************************************************************
 * ENGINE
 *******************************************************************/

function tarikLegacy_(sheetName) {
  const cfg = LEGACY_REPORT[sheetName];
  if (!cfg) throw new Error('Sheet "' + sheetName + '" gak terdaftar di LEGACY_REPORT.');

  const url = LEGACY.BASE_URL
            + '?wstoken=' + tokenLegacy_()
            + '&wsfunction=block_configurable_reports_get_report_data'
            + '&moodlewsrestformat=json'
            + '&reportid=' + cfg.reportId;

  // --- fetch + retry ---
  let raw = null, lastErr = null;
  for (let i = 1; i <= 3; i++) {
    try {
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
      const json = JSON.parse(res.getContentText());
      if (json.exception) throw new Error(json.errorcode + ': ' + json.message);
      if (!json.data) throw new Error('Response gak punya field .data');
      raw = json.data;
      break;
    } catch (e) {
      lastErr = e;
      Logger.log('  retry ' + i + '/3 — ' + e.message);
      Utilities.sleep(2000 * i);
    }
  }
  if (raw === null) throw lastErr;

  // --- decode ---
  const rows = (typeof raw === 'string') ? JSON.parse(raw) : raw;
  if (!Array.isArray(rows) || rows.length === 0) {
    Logger.log('  Data kosong, sheet gak disentuh.');
    return 0;
  }

  // --- dirty check via hash (gak baca sheet sama sekali) ---
  const props = PropertiesService.getScriptProperties();
  const hashKey = 'HASH_LEGACY_' + sheetName;
  const hashBaru = md5_(JSON.stringify(rows));
  if (props.getProperty(hashKey) === hashBaru) {
    Logger.log('  Data belum berubah, dilewat.');
    return -1;
  }

  // --- siapkan sheet ---
  const sh = SpreadsheetApp.openById(LEGACY.SS_ID).getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet "' + sheetName + '" gak ketemu.');

  const headers = Object.keys(rows[0]);
  const nKol = headers.length;

  const headerLama = sh.getLastColumn() > 0
    ? sh.getRange(1, 1, 1, nKol).getValues()[0] : [];

  // JANGAN sheet.clear() — itu ngehapus formula kolom M juga.
  // Cukup tulis ulang barisan header di kolom yang dipakai report.
  if (JSON.stringify(headers) !== JSON.stringify(headerLama)) {
    Logger.log('  Header berubah, header ditulis ulang (kolom M dst aman).');
    sh.getRange(1, 1, 1, nKol).setValues([headers]);
  }

  // --- bersihkan data lama (kolom A..nKol saja) ---
  const lastRow = legacyLastDataRow_(sh);
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, nKol).clearContent();

  // --- tulis data baru ---
  // ?? bukan || supaya angka 0 gak berubah jadi string kosong
  const matrix = rows.map(r => headers.map(h => {
    const v = r[h] ?? '';
    if (typeof v !== 'string') return v;
    const s = v.replace(/<[^>]*>/g, '').trim();
    if (/^-?[\d,]+\.?\d*$/.test(s) && s !== '') {
      const n = Number(s.replace(/,/g, ''));
      if (!isNaN(n)) return n;
    }
    return s;
  }));

  const CHUNK = 5000;
  for (let i = 0; i < matrix.length; i += CHUNK) {
    const bagian = matrix.slice(i, i + CHUNK);
    sh.getRange(2 + i, 1, bagian.length, nKol).setValues(bagian);
    SpreadsheetApp.flush();
  }

  props.setProperty(hashKey, hashBaru);
  Logger.log('  ' + matrix.length + ' baris ditulis.');
  return matrix.length;
}


/*******************************************************************
 * WRAPPER
 * Nama fungsi dipertahankan supaya Orkestrator.gs & pemanggil lain
 * gak perlu diubah.
 *******************************************************************/

function updateDataRapot(sheetName)       { return tarikLegacy_(sheetName || 'RAW'); }
function updateDataRapotRemed(sheetName)  { return tarikLegacy_(sheetName || 'Remedial'); }
function updateDataRapotRemed2(sheetName) { return tarikLegacy_(sheetName || 'Remedial 2'); }

/** Paksa tarik ulang walau hash-nya sama. Buat kalau sheet keburu diedit manual. */
function resetHashLegacy() {
  const props = PropertiesService.getScriptProperties();
  Object.keys(LEGACY_REPORT).forEach(s => props.deleteProperty('HASH_LEGACY_' + s));
  Logger.log('Hash direset. Tarikan berikutnya bakal nulis ulang semua.');
}
