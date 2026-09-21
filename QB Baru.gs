/*******************************************************************
 * QB BARU — Sertifikasi struktur Question Bank baru (API only)
 * -----------------------------------------------------------------
 * Narik data lewat web service Moodle murni, tanpa Configurable
 * Reports (kecuali sekali buat isi _SLOTMAP).
 *
 * Config ada di sheet _AKTIVITAS dan _SLOTMAP — gak perlu buka
 * Apps Script buat nambah aktivitas ujian baru.
 *
 * KOLOM M:
 *   - Baris QB Baru  -> diisi engine (teks, dari struktur QB)
 *   - Baris legacy   -> gak disentuh, formula tetap jalan
 *
 * Dijalankan lewat Orkestrator.gs (menu "Rapot"), bukan langsung.
 *
 * URUTAN SETUP (sekali aja):
 *   1. simpanToken()   -> tempel token service baru, lalu kosongin
 *   2. setupConfig()   -> bikin sheet _AKTIVITAS & _SLOTMAP
 *   3. cekKoneksi()    -> pastikan semua wsfunction keizinin
 *   4. updateSlotMap() -> isi _SLOTMAP dari report slot map
 *   5. isi _AKTIVITAS  -> daftar quiz yang mau ditarik
 *******************************************************************/

const ENGINE = {
  BASE_URL   : 'https://academy.amartha.com/webservice/rest/server.php',
  TOKEN_KEY  : 'WSTOKEN_V2',
  SS_ID      : '191zlOveljc2ADVVFPkR8r_OIwdANxed_uUGbDDAdyME',

  REPORT_SLOTMAP : 223,   // report kecil khusus slot -> topik/competency

  SHEET_AKT  : '_AKTIVITAS',
  SHEET_MAP  : '_SLOTMAP',

  TZ         : 'Asia/Jakarta',
  BATCH      : 25,        // request per fetchAll
  KOLOM      : 13,        // A..M — engine ikut ngisi Competency
  KOL_M      : 13         // posisi kolom Competency
};


/*******************************************************************
 * TOKEN
 *******************************************************************/

/** Jalanin sekali, habis itu KOSONGIN lagi string tokennya. */
function simpanToken() {
  PropertiesService.getScriptProperties()
    .setProperty(ENGINE.TOKEN_KEY, 'TEMPEL_TOKEN_BARU_DISINI');
  Logger.log('Token tersimpan. Sekarang kosongin lagi string di fungsi ini.');
}

function getToken_() {
  const t = PropertiesService.getScriptProperties().getProperty(ENGINE.TOKEN_KEY);
  if (!t || t.indexOf('TEMPEL') === 0) {
    throw new Error('Token belum di-set. Jalanin simpanToken() dulu.');
  }
  return t;
}


/*******************************************************************
 * SETUP CONFIG
 *******************************************************************/

/** Bikin sheet config kalau belum ada. Aman dijalanin berkali-kali. */
function setupConfig() {
  const ss = SpreadsheetApp.openById(ENGINE.SS_ID);

  if (!ss.getSheetByName(ENGINE.SHEET_AKT)) {
    const sh = ss.insertSheet(ENGINE.SHEET_AKT);
    sh.getRange(1, 1, 1, 6).setValues([[
      'aktif', 'quiz_id', 'group_ids', 'target_sheet', 'nama_ujian', 'catatan'
    ]]).setFontWeight('bold').setBackground('#f1f3f4');
    sh.getRange(2, 1, 4, 6).setValues([
      [true, 1165, '3622', 'RAW',      'Ujian Sertifikasi 2026',                 'Reguler'],
      [true, 1167, '3622', 'RAW',      'Ujian Sertifikasi 2026 (Resertifikasi)', 'Resert'],
      [true, 1146, '3622', 'RAW',      'Ujian Sertifikasi 2026 (Resert Hide)',   'Resert hide'],
      [true, 1166, '3622', 'Remedial', 'Ujian Sertifikasi 2026 (Remedial)',      'Remedial 1']
    ]);
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, 6);
  }

  if (!ss.getSheetByName(ENGINE.SHEET_MAP)) {
    const sh = ss.insertSheet(ENGINE.SHEET_MAP);
    sh.getRange(1, 1, 1, 4).setValues([[
      'quiz_id', 'slot', 'topik', 'competency'
    ]]).setFontWeight('bold').setBackground('#f1f3f4');
    sh.setFrozenRows(1);
  }

  Logger.log('Sheet config siap: ' + ENGINE.SHEET_AKT + ', ' + ENGINE.SHEET_MAP);
}


/*******************************************************************
 * API HELPER
 *******************************************************************/

function buildUrl_(fn, params) {
  const p = ['wstoken=' + getToken_(),
             'wsfunction=' + fn,
             'moodlewsrestformat=json'];
  Object.keys(params || {}).forEach(k => {
    const v = params[k];
    if (Array.isArray(v)) {
      v.forEach((item, i) => p.push(k + '[' + i + ']=' + encodeURIComponent(item)));
    } else {
      p.push(k + '=' + encodeURIComponent(v));
    }
  });
  return ENGINE.BASE_URL + '?' + p.join('&');
}

function api_(fn, params) {
  const res = UrlFetchApp.fetch(buildUrl_(fn, params), { muteHttpExceptions: true });
  const json = JSON.parse(res.getContentText());
  if (json && json.exception) {
    throw new Error(fn + ' -> ' + json.errorcode + ': ' + json.message);
  }
  return json;
}

/** Batch request. Balikin array hasil, elemen null kalau gagal. */
function apiAll_(reqs) {
  const out = [];
  for (let i = 0; i < reqs.length; i += ENGINE.BATCH) {
    const chunk = reqs.slice(i, i + ENGINE.BATCH);
    const urls = chunk.map(r => ({
      url: buildUrl_(r.fn, r.params),
      muteHttpExceptions: true
    }));
    let responses;
    try {
      responses = UrlFetchApp.fetchAll(urls);
    } catch (e) {
      Logger.log('  batch gagal, retry sekali: ' + e.message);
      Utilities.sleep(3000);
      responses = UrlFetchApp.fetchAll(urls);
    }
    responses.forEach(res => {
      try {
        const json = JSON.parse(res.getContentText());
        out.push(json && json.exception ? null : json);
      } catch (e) {
        out.push(null);
      }
    });
    Utilities.sleep(300);
  }
  return out;
}

/** Cek Moodle nyambung + wsfunction mana yang udah keizinin. */
function cekKoneksi() {
  const info = api_('core_webservice_get_site_info', {});
  Logger.log('Moodle ' + info.release + ' | user: ' + info.username);

  const wajib = [
    'mod_quiz_get_user_attempts',
    'mod_quiz_get_attempt_review',
    'core_group_get_group_members',
    'core_user_get_users_by_field'
  ];
  const opsional = [
    'core_group_get_groups',                      // nama group, ada fallback ke ID
    'block_configurable_reports_get_report_data'  // buat updateSlotMap
  ];

  const ada = (info.functions || []).map(f => f.name);
  let kurang = 0;

  Logger.log('--- WAJIB ---');
  wajib.forEach(f => {
    const ok = ada.indexOf(f) > -1;
    if (!ok) kurang++;
    Logger.log((ok ? 'OK    ' : 'BELUM ') + f);
  });

  Logger.log('--- OPSIONAL ---');
  opsional.forEach(f => Logger.log((ada.indexOf(f) > -1 ? 'OK    ' : 'BELUM ') + f));

  Logger.log(kurang === 0
    ? 'Semua fungsi wajib siap.'
    : kurang + ' fungsi wajib belum keizinin. Tambahin di External services.');
}


/*******************************************************************
 * SLOT MAP
 *******************************************************************/

/**
 * Isi _SLOTMAP dari custom report kecil.
 * Report harus balikin kolom: quiz_id, slot, topik, competency
 */
function isiSlotMapOtomatis(reportId) {
  if (!reportId) throw new Error('Kasih reportId-nya, contoh: isiSlotMapOtomatis(223)');

  const raw = api_('block_configurable_reports_get_report_data', { reportid: reportId });
  const rows = (typeof raw.data === 'string') ? JSON.parse(raw.data) : raw.data;
  if (!Array.isArray(rows) || !rows.length) throw new Error('Report kosong.');

  const bersih = s => String(s == null ? '' : s).replace(/<[^>]*>/g, '').trim();
  const data = rows.map(r => [
    Number(bersih(r.quiz_id)),
    Number(bersih(r.slot)),
    bersih(r.topik),
    bersih(r.competency)
  ]).filter(d => d[0] && d[1]);

  if (!data.length) throw new Error('Gak ada baris valid. Cek nama kolom di report.');

  const sh = SpreadsheetApp.openById(ENGINE.SS_ID).getSheetByName(ENGINE.SHEET_MAP);
  if (!sh) throw new Error('Sheet ' + ENGINE.SHEET_MAP + ' gak ada. Jalanin setupConfig().');

  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).clearContent();
  sh.getRange(2, 1, data.length, 4).setValues(data);

  const perQuiz = {};
  data.forEach(d => { perQuiz[d[0]] = (perQuiz[d[0]] || 0) + 1; });
  Logger.log('Slot map terisi: ' + data.length + ' baris');
  Object.keys(perQuiz).forEach(q => Logger.log('  quiz ' + q + ': ' + perQuiz[q] + ' slot'));
}

/** Dipanggil dari menu Orkestrator. */
function updateSlotMap() {
  if (!ENGINE.REPORT_SLOTMAP) {
    throw new Error('Isi dulu ENGINE.REPORT_SLOTMAP dengan ID report slot map.');
  }
  isiSlotMapOtomatis(ENGINE.REPORT_SLOTMAP);
}

function bacaSlotMap_() {
  const sh = SpreadsheetApp.openById(ENGINE.SS_ID).getSheetByName(ENGINE.SHEET_MAP);
  if (!sh || sh.getLastRow() < 2) throw new Error('_SLOTMAP masih kosong.');

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  const map = {};   // quizId -> { slot -> {topik, comp} }
  rows.forEach(r => {
    const qid = Number(r[0]), slot = Number(r[1]);
    if (!qid || !slot) return;
    if (!map[qid]) map[qid] = {};
    map[qid][slot] = { topik: String(r[2]).trim(), comp: String(r[3]).trim() };
  });
  return map;
}


/*******************************************************************
 * CONFIG AKTIVITAS
 *******************************************************************/

function bacaAktivitas_() {
  const sh = SpreadsheetApp.openById(ENGINE.SS_ID).getSheetByName(ENGINE.SHEET_AKT);
  if (!sh || sh.getLastRow() < 2) throw new Error('_AKTIVITAS masih kosong.');

  return sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues()
    .filter(r => r[0] === true || String(r[0]).toUpperCase() === 'TRUE')
    .map(r => ({
      quizId    : Number(r[1]),
      groupIds  : String(r[2]).split(/[,;|]/).map(s => Number(s.trim())).filter(Boolean),
      sheet     : String(r[3]).trim(),
      namaUjian : String(r[4]).trim()
    }))
    .filter(a => a.quizId && a.groupIds.length && a.sheet && a.namaUjian);
}

/** Semua nama_ujian yang dikelola engine, termasuk baris yang aktif = FALSE.
 *  Dipakai buat mastiin baris non-aktif gak dikira punya legacy. */
function semuaNamaUjianEngine_() {
  const sh = SpreadsheetApp.openById(ENGINE.SS_ID).getSheetByName(ENGINE.SHEET_AKT);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 5, sh.getLastRow() - 1, 1).getValues()
    .map(r => String(r[0]).trim())
    .filter(Boolean);
}


/*******************************************************************
 * TARIK DATA PER AKTIVITAS
 *******************************************************************/

function tarikAktivitas_(akt, slotMap) {
  const peta = slotMap[akt.quizId];
  if (!peta) {
    Logger.log('  LEWAT: slot map quiz ' + akt.quizId + ' belum ada di _SLOTMAP.');
    return [];
  }

  // --- 1. anggota group (group tertentu saja, sesuai config) ---
  const namaGroup = {};
  try {
    const info = api_('core_group_get_groups', { groupids: akt.groupIds });
    (info || []).forEach(g => { namaGroup[Number(g.id)] = g.name; });
  } catch (e) {
    Logger.log('  core_group_get_groups gagal (' + e.message + '), pakai ID group.');
  }

  const userGroup = {};   // userid -> nama group
  akt.groupIds.forEach(gid => {
    const res = api_('core_group_get_group_members', { groupids: [gid] });
    const anggota = (res[0] && res[0].userids) || [];
    const label = namaGroup[gid] || String(gid);
    anggota.forEach(uid => { if (!userGroup[uid]) userGroup[uid] = label; });
  });

  const userIds = Object.keys(userGroup).map(Number);
  if (!userIds.length) { Logger.log('  Group kosong.'); return []; }

  // --- 2. detail user ---
  const users = {};
  for (let i = 0; i < userIds.length; i += 50) {
    const batch = userIds.slice(i, i + 50);
    const res = api_('core_user_get_users_by_field', { field: 'id', values: batch });
    res.forEach(u => { users[u.id] = { nik: u.username, nama: u.fullname }; });
  }

  // --- 3. attempt per user ---
  const reqAtt = userIds.map(uid => ({
    fn: 'mod_quiz_get_user_attempts',
    params: { quizid: akt.quizId, userid: uid, status: 'finished' }
  }));
  const hasilAtt = apiAll_(reqAtt);

  const attemptTerpilih = [];
  hasilAtt.forEach((res, i) => {
    const list = (res && res.attempts) || [];
    if (!list.length) return;
    const last = list.reduce((a, b) => (Number(b.attempt) > Number(a.attempt) ? b : a));
    attemptTerpilih.push({ userId: userIds[i], attempt: last });
  });

  if (!attemptTerpilih.length) { Logger.log('  Belum ada attempt selesai.'); return []; }

  // --- 4. review per attempt ---
  const reqRev = attemptTerpilih.map(a => ({
    fn: 'mod_quiz_get_attempt_review',
    params: { attemptid: a.attempt.id, page: -1 }
  }));
  const hasilRev = apiAll_(reqRev);

  // --- 5. agregasi per topik ---
  const baris = [];
  let slotTakDikenal = 0;

  hasilRev.forEach((rev, i) => {
    if (!rev || !rev.questions) return;
    const a   = attemptTerpilih[i];
    const usr = users[a.userId];
    if (!usr) return;

    const agg = {};       // topik -> {total, skor, max, comp, urut}
    let quizMax = 0;

    rev.questions.forEach(q => {
      const maxmark = parseFloat(q.maxmark) || 0;
      quizMax += maxmark;

      const m = peta[Number(q.slot)];
      if (!m) { slotTakDikenal++; return; }

      const mark = (q.mark === null || q.mark === '' || q.mark === undefined)
                   ? 0 : (parseFloat(q.mark) || 0);

      if (!agg[m.topik]) {
        agg[m.topik] = { total: 0, skor: 0, max: 0, comp: m.comp, urut: Number(q.slot) };
      }
      agg[m.topik].total += 1;
      agg[m.topik].skor  += mark;
      agg[m.topik].max   += maxmark;
    });

    const sumgrades  = parseFloat(a.attempt.sumgrades) || 0;
    const nilaiAkhir = quizMax > 0 ? Math.round(sumgrades / quizMax * 10000) / 100 : 0;
    const tf         = Number(a.attempt.timefinish) * 1000;
    const tgl        = Utilities.formatDate(new Date(tf), ENGINE.TZ, 'yyyy-MM-dd');
    const jam        = Utilities.formatDate(new Date(tf), ENGINE.TZ, 'HH:mm:ss');

    Object.keys(agg)
      .sort((x, y) => agg[x].urut - agg[y].urut)
      .forEach(topik => {
        const d = agg[topik];
        baris.push([
          a.userId,
          usr.nik,
          usr.nama,
          akt.namaUjian,
          userGroup[a.userId],
          topik,
          d.total,
          Math.round(d.skor * 100) / 100,
          d.max > 0 ? Math.round(d.skor / d.max * 10000) / 100 : 0,
          nilaiAkhir,
          tgl,
          jam,
          d.comp                       // kolom M diisi engine
        ]);
      });
  });

  if (slotTakDikenal) {
    Logger.log('  PERHATIAN: ' + slotTakDikenal +
               ' slot gak ada di _SLOTMAP, skornya gak masuk topik manapun.');
  }
  Logger.log('  quiz ' + akt.quizId + ': ' + attemptTerpilih.length +
             ' peserta, ' + baris.length + ' baris');
  return baris;
}


/*******************************************************************
 * TULIS KE SHEET
 *******************************************************************/

/** Baris terakhir yang beneran ada datanya, diukur dari kolom A.
 *  Perlu karena kolom M berisi formula sampai jauh ke bawah,
 *  bikin getLastRow() ngasih angka yang menyesatkan. */
function lastDataRow_(sh) {
  const n = sh.getLastRow();
  if (n < 2) return 1;
  const colA = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = colA.length - 1; i >= 0; i--) {
    if (String(colA[i][0]).trim() !== '') return i + 2;
  }
  return 1;
}

/**
 * Hapus baris milik engine (kolom D cocok), lalu tulis baris baru
 * di bawah data terakhir.
 *
 * Penghapusan dilakukan per BLOK kontigu, bukan per baris:
 * 119 kali deleteRow itu lambat dan boros kuota.
 * Baris legacy gak disentuh selama nama_ujian-nya beda.
 */
function tulisKeSheet_(namaSheet, barisBaru, namaUjianDikelola) {
  const sh = SpreadsheetApp.openById(ENGINE.SS_ID).getSheetByName(namaSheet);
  if (!sh) { Logger.log('Sheet "' + namaSheet + '" gak ketemu, dilewat.'); return 0; }

  // --- kumpulkan baris milik engine ---
  const lastRow = lastDataRow_(sh);
  let dihapus = 0;

  if (lastRow > 1) {
    const kolomD = sh.getRange(2, 4, lastRow - 1, 1).getValues();
    const milik = [];
    for (let i = 0; i < kolomD.length; i++) {
      if (namaUjianDikelola.indexOf(String(kolomD[i][0]).trim()) > -1) milik.push(i + 2);
    }

    // gabung jadi blok kontigu, hapus dari bawah supaya index gak geser
    const blok = [];
    milik.forEach(r => {
      const t = blok[blok.length - 1];
      if (t && r === t.mulai + t.jumlah) t.jumlah++;
      else blok.push({ mulai: r, jumlah: 1 });
    });
    for (let i = blok.length - 1; i >= 0; i--) {
      sh.deleteRows(blok[i].mulai, blok[i].jumlah);
      dihapus += blok[i].jumlah;
    }
    if (dihapus) {
      Logger.log(namaSheet + ': ' + dihapus + ' baris lama dihapus (' +
                 blok.length + ' blok).');
    }
  }

  if (!barisBaru.length) {
    Logger.log(namaSheet + ': gak ada baris baru.');
    return 0;
  }

  // --- tulis baris baru di bawah data terakhir ---
  const mulai = lastDataRow_(sh) + 1;
  const CHUNK = 5000;
  for (let i = 0; i < barisBaru.length; i += CHUNK) {
    const bagian = barisBaru.slice(i, i + CHUNK);
    sh.getRange(mulai + i, 1, bagian.length, ENGINE.KOLOM).setValues(bagian);
    SpreadsheetApp.flush();
  }

  Logger.log(namaSheet + ': ' + barisBaru.length + ' baris ditulis mulai row ' + mulai);
  return barisBaru.length;
}


/*******************************************************************
 * PEMULIH FORMULA KOLOM M
 * -----------------------------------------------------------------
 * Baris engine ditulis sebagai teks, dan penghapusan baris bikin
 * stok baris ber-formula di bawah berkurang tiap run. Fungsi ini
 * nyalin ulang formula M dari baris referensi ke SELURUH baris
 * legacy (yang nama_ujian-nya bukan punya engine).
 *******************************************************************/

function rapikanFormulaM(namaSheet) {
  const sh = SpreadsheetApp.openById(ENGINE.SS_ID).getSheetByName(namaSheet);
  if (!sh) throw new Error('Sheet "' + namaSheet + '" gak ketemu.');

  const rumus = sh.getRange(2, ENGINE.KOL_M).getFormula();
  if (!rumus) {
    throw new Error('Baris 2 kolom M bukan formula. ' +
                    'Isi dulu M2 dengan formula acuan, baru jalanin lagi.');
  }

  const lastRow = lastDataRow_(sh);
  if (lastRow < 3) { Logger.log('Belum ada data buat dirapikan.'); return 0; }

  const namaEngine = semuaNamaUjianEngine_();
  const kolomD = sh.getRange(2, 4, lastRow - 1, 1).getValues();

  // baris legacy = yang nama_ujian-nya BUKAN punya engine
  const legacy = [];
  for (let i = 0; i < kolomD.length; i++) {
    const d = String(kolomD[i][0]).trim();
    if (d && namaEngine.indexOf(d) === -1) legacy.push(i + 2);
  }
  if (!legacy.length) { Logger.log('Gak ada baris legacy.'); return 0; }

  // salin formula per blok kontigu (copyTo otomatis nyesuaiin referensi relatif)
  const blok = [];
  legacy.forEach(r => {
    const t = blok[blok.length - 1];
    if (t && r === t.mulai + t.jumlah) t.jumlah++;
    else blok.push({ mulai: r, jumlah: 1 });
  });

  const sumber = sh.getRange(2, ENGINE.KOL_M);
  let total = 0;
  blok.forEach(b => {
    sumber.copyTo(sh.getRange(b.mulai, ENGINE.KOL_M, b.jumlah, 1),
                  SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
    total += b.jumlah;
  });

  SpreadsheetApp.flush();
  Logger.log(namaSheet + ': formula M dipulihkan di ' + total +
             ' baris legacy (' + blok.length + ' blok).');
  return total;
}

/** Rapikan formula M di semua sheet target yang dipakai engine. */
function rapikanFormulaMSemua() {
  const target = {};
  bacaAktivitas_().forEach(a => { target[a.sheet] = true; });

  let total = 0;
  Object.keys(target).forEach(nama => {
    try {
      total += rapikanFormulaM(nama);
    } catch (e) {
      Logger.log(nama + ': ' + e.message);
    }
  });
  return total;
}


/*******************************************************************
 * RUNNER
 * Dipanggil Orkestrator.gs. Jangan bikin jalankanSemua() di sini —
 * nama itu sudah dipakai Orkestrator dan bakal bentrok.
 *******************************************************************/

function tarikDataBaru() {
  const t0 = Date.now();
  const slotMap = bacaSlotMap_();
  const daftar  = bacaAktivitas_();

  Logger.log('=== ' + daftar.length + ' aktivitas aktif ===');

  const perSheet = {};
  const ujianPerSheet = {};

  daftar.forEach(akt => {
    Logger.log('> quiz ' + akt.quizId + ' -> ' + akt.sheet);
    if (!perSheet[akt.sheet]) { perSheet[akt.sheet] = []; ujianPerSheet[akt.sheet] = []; }
    ujianPerSheet[akt.sheet].push(akt.namaUjian);

    try {
      perSheet[akt.sheet] = perSheet[akt.sheet].concat(tarikAktivitas_(akt, slotMap));
    } catch (e) {
      Logger.log('  ERROR quiz ' + akt.quizId + ': ' + e.message);
    }
  });

  let total = 0;
  Object.keys(perSheet).forEach(nama => {
    total += tulisKeSheet_(nama, perSheet[nama], ujianPerSheet[nama]);
  });

  Logger.log('Selesai: ' + total + ' baris, ' +
             ((Date.now() - t0) / 1000).toFixed(1) + ' detik.');
  return total;
}
