/*******************************************************************
 * ORKESTRATOR — Rapot Sertifikasi  (versi final)
 * -----------------------------------------------------------------
 * Menggantikan Call Update.gs sepenuhnya.
 * Ngatur urutan, nyatet log, bikin menu, pasang trigger harian.
 *
 * STRUKTUR PROJECT SETELAH MIGRASI:
 *   LegacyEngine.gs  -> tarik dari Configurable Reports (119/130/136)
 *   QBBaru.gs        -> tarik QB baru via API murni
 *   Orkestrator.gs   -> file ini
 *
 * SETUP:
 *   1. Hapus Call Update.gs + 3 file Tarik Data*.gs
 *   2. Refresh spreadsheet -> menu "Rapot" muncul
 *   3. Menu > Trigger > Pasang trigger harian
 *******************************************************************/

const PIPE = {
  SS_ID     : '191zlOveljc2ADVVFPkR8r_OIwdANxed_uUGbDDAdyME',
  SHEET_LOG : '_LOG',
  MAX_LOG   : 400,
  JAM_HARIAN: 5,
  TZ        : 'Asia/Jakarta'
};

/**
 * URUTAN WAJIB: legacy dulu, QB Baru terakhir.
 * Legacy ngosongin kolom A..L di sheet targetnya, termasuk baris
 * punya QB Baru. Kalau urutannya kebalik, data 2026 hilang.
 *
 * Langkah terakhir mulihin formula kolom M di baris legacy —
 * baris engine sendiri sengaja dilewat karena udah diisi teks.
 */
const LANGKAH = [
  { aktif: true, nama: 'Legacy RAW',        fn: 'updateDataRapot',       args: ['RAW'],        sheet: 'RAW' },
  { aktif: true, nama: 'Legacy Remedial',   fn: 'updateDataRapotRemed',  args: ['Remedial'],   sheet: 'Remedial' },
  { aktif: true, nama: 'Legacy Remedial 2', fn: 'updateDataRapotRemed2', args: ['Remedial 2'], sheet: 'Remedial 2' },
  { aktif: true, nama: 'QB Baru (API)',     fn: 'tarikDataBaru',         args: [],             sheet: null },
  { aktif: true, nama: 'Rapikan formula M', fn: 'rapikanFormulaMSemua',  args: [],             sheet: null }
];

/** Dua langkah penutup: QB Baru + rapikan formula M. */
function langkahPenutup_() {
  return LANGKAH.slice(-2);
}


/*******************************************************************
 * MENU
 *******************************************************************/

function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Rapot')
      .addItem('Jalankan semua', 'menuJalankanSemua')
      .addSeparator()
      .addSubMenu(ui.createMenu('Update per sheet')
        .addItem('RAW', 'menuUpdateRaw')
        .addItem('Remedial', 'menuUpdateRemedial')
        .addItem('Remedial 2', 'menuUpdateRemedial2')
        .addSeparator()
        .addItem('QB Baru saja', 'menuQbBaruSaja'))
      .addSeparator()
      .addItem('Update slot map', 'menuUpdateSlotMap')
      .addItem('Rapikan formula kolom M', 'menuRapikanFormulaM')
      .addItem('Cek koneksi Moodle', 'menuCekKoneksi')
      .addSeparator()
      .addItem('Buka log', 'menuBukaLog')
      .addSubMenu(ui.createMenu('Trigger')
        .addItem('Pasang trigger harian', 'pasangTriggerHarian')
        .addItem('Hapus trigger', 'hapusTrigger')
        .addItem('Status trigger', 'statusTrigger'))
      .addToUi();
  } catch (e) { /* dipanggil dari trigger, gak ada UI */ }
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
  if (!konfirmasi_('Jalankan semua?',
      'Semua langkah dijalankan berurutan. Bisa makan 2-5 menit.\n' +
      'Jangan tutup spreadsheet sampai selesai.')) return;
  laporkan_(jalankanSemua());
}

/**
 * Update satu sheet legacy, lalu QB Baru + rapikan formula M.
 * Wajib begitu: legacy ngehapus baris QB Baru di sheet itu.
 */
function updateSheetLegacy_(namaLangkah) {
  const lk = LANGKAH.filter(x => x.nama === namaLangkah)[0];
  if (!lk) { const u = ui_(); if (u) u.alert('Langkah "' + namaLangkah + '" gak ketemu.'); return; }

  if (!konfirmasi_('Konfirmasi update',
      'Update sheet "' + lk.sheet + '"?\n\n' +
      'QB Baru ikut dijalankan ulang setelahnya supaya data 2026 gak hilang.')) {
    const u = ui_(); if (u) u.alert('Update dibatalin.');
    return;
  }

  const runId = 'MANUAL-' + Utilities.formatDate(new Date(), PIPE.TZ, 'HHmmss');
  laporkan_(jalankanLangkah_([lk].concat(langkahPenutup_()), runId));
}

function menuUpdateRaw()       { updateSheetLegacy_('Legacy RAW'); }
function menuUpdateRemedial()  { updateSheetLegacy_('Legacy Remedial'); }
function menuUpdateRemedial2() { updateSheetLegacy_('Legacy Remedial 2'); }

function menuQbBaruSaja() {
  if (!konfirmasi_('Jalankan QB Baru saja?',
      'Legacy TIDAK dijalankan. Aman selama legacy gak jalan sesudahnya.')) return;
  const runId = 'MANUAL-QB-' + Utilities.formatDate(new Date(), PIPE.TZ, 'HHmmss');
  laporkan_(jalankanLangkah_(langkahPenutup_(), runId));
}

function menuRapikanFormulaM() {
  const u = ui_();
  if (!konfirmasi_('Rapikan formula kolom M?',
      'Formula di M2 disalin ke semua baris LEGACY.\n' +
      'Baris QB Baru dilewat (Competency-nya diisi engine).')) return;
  try {
    const n = rapikanFormulaMSemua();
    catat_('SETUP', 'Rapikan formula M', 'OK', 0, n, 'manual');
    if (u) u.alert('Formula M dipulihkan di ' + n + ' baris legacy.');
  } catch (e) {
    catat_('SETUP', 'Rapikan formula M', 'GAGAL', 0, '', e.message);
    if (u) u.alert('Gagal', e.message, u.ButtonSet.OK); else Logger.log(e.message);
  }
}

function menuUpdateSlotMap() {
  const u = ui_();
  try {
    if (typeof updateSlotMap === 'function') {
      updateSlotMap();
    } else if (typeof ENGINE !== 'undefined' && ENGINE.REPORT_SLOTMAP) {
      isiSlotMapOtomatis(ENGINE.REPORT_SLOTMAP);
    } else {
      throw new Error('Isi dulu ENGINE.REPORT_SLOTMAP di QBBaru.gs');
    }
    catat_('SETUP', 'Slot map', 'OK', 0, '', 'diperbarui manual');
    if (u) u.alert('Slot map ter-update. Cek sheet _SLOTMAP.');
  } catch (e) {
    catat_('SETUP', 'Slot map', 'GAGAL', 0, '', e.message);
    if (u) u.alert('Gagal', e.message, u.ButtonSet.OK); else Logger.log(e.message);
  }
}

function menuCekKoneksi() {
  const u = ui_();
  try {
    cekKoneksi();
    if (u) u.alert('Cek koneksi jalan. Hasil lengkapnya di Execution log.');
  } catch (e) {
    if (u) u.alert('Gagal', e.message, u.ButtonSet.OK);
  }
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
    const warna = { OK: '#e6f4ea', GAGAL: '#fce8e6', SKIP: '#fff3e0',
                    MULAI: '#e8f0fe', LEWAT: '#f5f5f5' }[status];
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
    const colA = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (let i = colA.length - 1; i >= 0; i--) {
      if (String(colA[i][0]).trim() !== '') return i + 1;
    }
    return 0;
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
      if (!lk.aktif) {
        catat_(runId, lk.nama, 'SKIP', 0, '', 'aktif = false');
        hasil.push({ nama: lk.nama, status: 'SKIP', ket: 'dinonaktifkan' });
        return;
      }

      const fn = (typeof this[lk.fn] === 'function') ? this[lk.fn] : null;
      if (!fn) {
        catat_(runId, lk.nama, 'SKIP', 0, '', 'Fungsi "' + lk.fn + '" gak ketemu');
        hasil.push({ nama: lk.nama, status: 'SKIP', ket: 'fungsi ' + lk.fn + ' gak ada' });
        return;
      }

      const sebelum = barisData_(lk.sheet);
      const tMulai = Date.now();
      toast_(lk.nama + '...');

      try {
        const kembali = fn.apply(null, lk.args || []);
        const durasi = ((Date.now() - tMulai) / 1000).toFixed(1);

        // LegacyEngine balikin -1 kalau data gak berubah
        if (kembali === -1) {
          catat_(runId, lk.nama, 'LEWAT', durasi, sebelum, 'Data belum berubah');
          hasil.push({ nama: lk.nama, status: 'LEWAT', durasi: durasi, ket: 'data belum berubah' });
          return;
        }

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
    const ikon = { OK: '[OK]', GAGAL: '[X]', SKIP: '[-]', LEWAT: '[=]' }[h.status] || '[?]';
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
