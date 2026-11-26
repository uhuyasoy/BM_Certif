// Menu, log, dan trigger harian Rapot Sertifikasi.
const PIPE = {
  SS_ID: '191zlOveljc2ADVVFPkR8r_OIwdANxed_uUGbDDAdyME',
  SHEET_LOG: '_LOG',
  MAX_LOG: 400,
  JAM_HARIAN: 5,
  TZ: 'Asia/Jakarta'
};

const LANGKAH = [
  { nama: 'RAW', jalankan: () => updateDataRapot(), sheet: 'RAW' },
  { nama: 'Remedial', jalankan: () => updateDataRapotRemed(), sheet: 'Remedial' },
  { nama: 'Resertifikasi', jalankan: () => updateDataRapotRemed2(), sheet: 'Resertifikasi' }
];

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Rapot')
    .addItem('Jalankan semua', 'menuJalankanSemua')
    .addSubMenu(ui.createMenu('Update per sheet')
      .addItem('RAW', 'menuUpdateRaw')
      .addItem('Remedial', 'menuUpdateRemedial')
      .addItem('Resertifikasi', 'menuUpdateResertifikasi'))
    .addItem('Update slot map', 'menuUpdateSlotMap')
    .addItem('Buka log', 'menuBukaLog')
    .addSubMenu(ui.createMenu('Trigger')
      .addItem('Pasang trigger harian', 'pasangTriggerHarian')
      .addItem('Hapus trigger', 'hapusTrigger')
      .addItem('Status trigger', 'statusTrigger'))
    .addToUi();
}

function ui_() {
  try { return SpreadsheetApp.getUi(); } catch (e) { return null; }
}

function toast_(pesan) {
  try { SpreadsheetApp.openById(PIPE.SS_ID).toast(pesan, 'Rapot', 5); } catch (e) {}
}

function konfirmasi_(judul, pesan) {
  const u = ui_();
  if (!u) return true;
  return u.alert(judul, pesan, u.ButtonSet.YES_NO) === u.Button.YES;
}

function menuJalankanSemua() {
  if (!konfirmasi_('Jalankan semua?', 'Sinkronkan RAW, Remedial, dan Resertifikasi?')) return;
  laporkan_(jalankanSemua());
}

function menuUpdateSheet_(namaSheet) {
  if (!konfirmasi_('Konfirmasi update', 'Update sheet ' + namaSheet + '?')) return;
  const langkah = LANGKAH.find(item => item.sheet === namaSheet);
  const runId = 'MANUAL-' + Utilities.formatDate(new Date(), PIPE.TZ, 'yyyyMMdd-HHmmss');
  laporkan_(jalankanLangkah_([langkah], runId));
}

function menuUpdateRaw() { menuUpdateSheet_('RAW'); }
function menuUpdateRemedial() { menuUpdateSheet_('Remedial'); }
function menuUpdateResertifikasi() { menuUpdateSheet_('Resertifikasi'); }

function menuUpdateSlotMap() {
  const runId = 'SLOTMAP-' + Utilities.formatDate(new Date(), PIPE.TZ, 'yyyyMMdd-HHmmss');
  laporkan_(jalankanLangkah_([{ nama: 'Slot map', jalankan: updateSlotMap, sheet: RAPOT_REPORT.SLOT_SHEET }], runId));
}

function menuBukaLog() {
  const ss = SpreadsheetApp.openById(PIPE.SS_ID);
  const sh = sheetLog_();
  ss.setActiveSheet(sh);
  sh.setActiveRange(sh.getRange(Math.max(2, sh.getLastRow() - 20), 1));
}

function laporkan_(hasil) {
  const u = ui_();
  if (!u) return;
  const gagal = hasil.filter(h => h.status === 'GAGAL').length;
  u.alert(gagal ? 'Selesai dengan ' + gagal + ' kegagalan' : 'Selesai',
          ringkasanTeks_(hasil), u.ButtonSet.OK);
}


/*******************************************************************
 * LOG
 *******************************************************************/

function sheetLog_() {
  const ss = SpreadsheetApp.openById(PIPE.SS_ID);
  let sh = ss.getSheetByName(PIPE.SHEET_LOG);
  if (!sh) {
    sh = ss.insertSheet(PIPE.SHEET_LOG);
    sh.getRange(1, 1, 1, 7).setValues([[
      'waktu', 'run_id', 'langkah', 'status', 'durasi (dtk)', 'baris', 'keterangan'
    ]]).setFontWeight('bold').setBackground('#f1f3f4');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 150);
    sh.setColumnWidth(3, 150);
    sh.setColumnWidth(7, 420);
  }
  return sh;
}

function catat_(runId, langkah, status, durasi, baris, ket) {
  try {
    const sh = sheetLog_();
    sh.appendRow([
      Utilities.formatDate(new Date(), PIPE.TZ, 'yyyy-MM-dd HH:mm:ss'),
      runId, langkah, status, durasi, baris, String(ket || '').substring(0, 500)
    ]);
    const warna = { OK: '#e6f4ea', GAGAL: '#fce8e6', MULAI: '#e8f0fe' }[status];
    if (warna) sh.getRange(sh.getLastRow(), 1, 1, 7).setBackground(warna);
  } catch (e) {
    Logger.log('Gagal nulis log: ' + e.message);
  }
}

function pangkasLog_() {
  try {
    const sh = sheetLog_();
    const lebih = sh.getLastRow() - 1 - PIPE.MAX_LOG;
    if (lebih > 0) sh.deleteRows(2, lebih);
  } catch (e) {}
}

/** Baris data terakhir, diukur dari kolom A (kolom M penuh formula). */
function barisData_(namaSheet) {
  if (!namaSheet) return '';
  try {
    const sh = SpreadsheetApp.openById(PIPE.SS_ID).getSheetByName(namaSheet);
    if (!sh || sh.getLastRow() < 2) return 0;
    return Math.max(0, barisTerakhirRapot_(sh) - 1);
  } catch (e) { return ''; }
}


/*******************************************************************
 * RUNNER
 *******************************************************************/

function jalankanLangkah_(daftar, runId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    catat_(runId, '-', 'GAGAL', 0, '', 'Ada proses lain yang masih jalan.');
    return [{ nama: '-', status: 'GAGAL', ket: 'Proses lain masih jalan' }];
  }

  const hasil = [];
  const t0 = Date.now();
  catat_(runId, '-', 'MULAI', 0, '', daftar.length + ' langkah');
  toast_('Mulai: ' + daftar.length + ' langkah');

  try {
    daftar.forEach(lk => {
      const sebelum = barisData_(lk.sheet);
      const tMulai = Date.now();
      toast_(lk.nama + '...');

      try {
        const kembali = lk.jalankan();
        const durasi = ((Date.now() - tMulai) / 1000).toFixed(1);

        let info;
        if (lk.sheet && sebelum !== '') {
          info = sebelum + ' -> ' + barisData_(lk.sheet);
        } else {
          info = (typeof kembali === 'number') ? String(kembali) : '';
        }
        catat_(runId, lk.nama, 'OK', durasi, info, '');
        hasil.push({ nama: lk.nama, status: 'OK', durasi: durasi, ket: info });

      } catch (e) {
        const durasi = ((Date.now() - tMulai) / 1000).toFixed(1);
        catat_(runId, lk.nama, 'GAGAL', durasi, '', e.message);
        hasil.push({ nama: lk.nama, status: 'GAGAL', durasi: durasi, ket: e.message });
        Logger.log('GAGAL ' + lk.nama + ': ' + e.stack);
      }
    });

  } finally {
    lock.releaseLock();
  }

  const total = ((Date.now() - t0) / 1000).toFixed(1);
  const gagal = hasil.filter(h => h.status === 'GAGAL').length;
  catat_(runId, '-', gagal ? 'GAGAL' : 'OK', total, '',
         'Selesai. ' + (hasil.length - gagal) + '/' + hasil.length + ' berhasil.');
  pangkasLog_();
  toast_(gagal ? gagal + ' langkah gagal, cek _LOG' : 'Semua selesai (' + total + ' dtk)');

  return hasil;
}

/** PASANG INI DI TRIGGER HARIAN. */
function jalankanSemua() {
  const runId = Utilities.formatDate(new Date(), PIPE.TZ, 'yyyyMMdd-HHmmss');
  return jalankanLangkah_(LANGKAH, runId);
}

function ringkasanTeks_(hasil) {
  return hasil.map(h => {
    const ikon = { OK: '[OK]', GAGAL: '[X]' }[h.status] || '[?]';
    return ikon + ' ' + h.nama +
           (h.durasi ? ' (' + h.durasi + ' dtk)' : '') +
           (h.ket ? ' - ' + h.ket : '');
  }).join('\n');
}


/*******************************************************************
 * TRIGGER
 *******************************************************************/

function pasangTriggerHarian() {
  hapusTrigger();
  ScriptApp.newTrigger('jalankanSemua')
    .timeBased().everyDays(1)
    .atHour(PIPE.JAM_HARIAN)
    .inTimezone(PIPE.TZ)
    .create();

  const pesan = 'Trigger harian dipasang jam ' + PIPE.JAM_HARIAN + ':00 WIB.';
  catat_('SETUP', 'Trigger', 'OK', 0, '', pesan);
  const u = ui_();
  if (u) u.alert(pesan); else Logger.log(pesan);
}

function hapusTrigger() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'jalankanSemua') { ScriptApp.deleteTrigger(t); n++; }
  });
  if (n) catat_('SETUP', 'Trigger', 'OK', 0, '', n + ' trigger dihapus');
  return n;
}

function statusTrigger() {
  const list = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'jalankanSemua');
  const pesan = list.length
    ? list.length + ' trigger aktif untuk jalankanSemua (jam ' + PIPE.JAM_HARIAN + ':00 WIB).'
    : 'Belum ada trigger terpasang.';
  const u = ui_();
  if (u) u.alert(pesan); else Logger.log(pesan);
}
