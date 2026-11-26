// Sinkronisasi Configurable Reports Moodle ke spreadsheet Rapot.
const RAPOT_REPORT = {
  URL: 'https://academy.amartha.com/webservice/rest/server.php',
  TOKEN: 'WSTOKEN_LEGACY',
  SLOT_TOKEN: 'WSTOKEN_V2',
  CHUNK_SIZE: 5000,
  TARGETS: { RAW: [119], Remedial: [130], Resertifikasi: [136, 220] },
  SLOT_ID: 223,
  SLOT_SHEET: '_SLOTMAP'
};

const RAPOT_HEADER = [
  'nik', 'nama', 'nama_ujian', 'nama_group', 'topik', 'total_soal',
  'skor_per_topik', 'persentase', 'nilai_akhir', 'tanggal', 'waktu',
  'last_sync_time', 'Competency', 'Competency Add'
];

function requestReport_(reportId, tokenKey) {
  const token = PropertiesService.getScriptProperties().getProperty(tokenKey);
  if (!token || /^(TEMPEL|TOKEN_LAMA$)/i.test(token)) {
    throw new Error('Isi Script Property ' + tokenKey + ' dengan token Moodle.');
  }

  const url = RAPOT_REPORT.URL + '?wstoken=' + encodeURIComponent(token)
    + '&wsfunction=block_configurable_reports_get_report_data'
    + '&moodlewsrestformat=json&reportid=' + reportId;

  let errorMessage = '';
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      let response;
      try {
        response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      } catch (e) {
        throw new Error('Request ke Moodle gagal.');
      }
      if (response.getResponseCode() !== 200) {
        throw new Error('HTTP ' + response.getResponseCode());
      }
      const payload = JSON.parse(response.getContentText());
      if (payload.exception) {
        throw new Error(payload.errorcode || 'Moodle web service error');
      }
      const rows = typeof payload.data === 'string' ? JSON.parse(payload.data) : payload.data;
      if (!Array.isArray(rows)) throw new Error('Field data harus berupa array.');
      if (rows.some(row => !row || typeof row !== 'object' || Array.isArray(row))) {
        throw new Error('Format baris report harus berupa objek dengan nama kolom.');
      }
      return rows;
    } catch (e) {
      errorMessage = e.message;
      if (attempt < 3) Utilities.sleep(attempt * 2000);
    }
  }
  throw new Error('Report ' + reportId + ': ' + errorMessage);
}

function reportKey_(name) {
  const key = name.trim().toLowerCase().replace(/\s+/g, '_');
  return key === 'report_generated_at' ? 'last_sync_time' : key;
}

function reportText_(value) {
  return value == null ? '' : String(value).replace(/<[^>]*>/g, '').trim();
}

function reportValue_(value, key) {
  if (typeof value !== 'string') return value == null ? '' : value;
  const clean = reportText_(value);
  if (key === 'nik' && /^0\d/.test(clean)) return clean;
  if (/^-?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?$/.test(clean)) {
    return Number(clean.replace(/,/g, ''));
  }
  return clean.startsWith('=') ? "'" + clean : clean;
}

function buildReportTable_(reports) {
  const headerByKey = new Map(RAPOT_HEADER.map(name => [reportKey_(name), name]));
  const competencyWithData = new Set();
  const data = [];
  reports.forEach(({ id, rows }) => {
    rows.forEach((row, index) => {
      const values = new Map();
      Object.keys(row).forEach(name => {
        const key = reportKey_(name);
        if (!headerByKey.has(key)) headerByKey.set(key, name);
        const value = reportValue_(row[name], key);
        values.set(key, value);
        if ((key === 'competency' || key === 'competency_add') && value !== '') {
          competencyWithData.add(key);
        }
      });
      for (const key of ['nik', 'nama_ujian', 'topik']) {
        if (!values.has(key)) {
          throw new Error('Report ' + id + ', baris ' + (index + 1) + ': kolom ' + key + ' tidak ditemukan.');
        }
      }
      data.push(values);
    });
  });

  const keys = Array.from(headerByKey.keys());
  return {
    headers: Array.from(headerByKey.values()),
    rows: data.map(row => keys.map(key => row.has(key) ? row.get(key) : '')),
    competencyWithData
  };
}

function pastikanUkuranSheet_(sheet, rows, columns) {
  const maxRows = sheet.getMaxRows();
  const maxColumns = sheet.getMaxColumns();
  if (rows > maxRows) sheet.insertRowsAfter(maxRows, rows - maxRows);
  if (columns > maxColumns) sheet.insertColumnsAfter(maxColumns, columns - maxColumns);
}

function barisTerakhirRapot_(sheet) {
  const last = sheet.getLastRow();
  if (!last) return 0;
  const values = sheet.getRange(1, 1, last, 1).getValues();
  for (let index = values.length - 1; index >= 0; index--) {
    if (String(values[index][0]).trim() !== '') return index + 1;
  }
  return 0;
}

function tulisReport_(sheet, headers, rows, oldLastRow, formulaColumns = {}) {
  const width = headers.length;
  const end = rows.length + 1;
  pastikanUkuranSheet_(sheet, Math.max(2, end), width);
  sheet.getRange(1, 1, 1, width).setValues([headers]);

  // Tulis kolom data tanpa menimpa formula acuan di M2/N2.
  const spans = [];
  let start = 1;
  for (let column = 1; column <= width + 1; column++) {
    if (column <= width && !formulaColumns[column]) continue;
    if (column > start) spans.push({ start, width: column - start });
    start = column + 1;
  }
  spans.forEach(span => {
    for (let index = 0; index < rows.length; index += RAPOT_REPORT.CHUNK_SIZE) {
      const chunk = rows.slice(index, index + RAPOT_REPORT.CHUNK_SIZE);
      const values = chunk.map(row => row.slice(span.start - 1, span.start - 1 + span.width));
      sheet.getRange(index + 2, span.start, values.length, span.width).setValues(values);
    }
    if (oldLastRow > end) {
      sheet.getRange(end + 1, span.start, oldLastRow - end, span.width).clearContent();
    }
  });

  Object.keys(formulaColumns).forEach(key => {
    const column = Number(key);
    const formula = formulaColumns[key];
    if (/\bARRAYFORMULA\s*\(/i.test(formula)) return;
    if (rows.length > 1) {
      sheet.getRange(3, column, rows.length - 1, 1)
        .setFormulasR1C1(Array.from({ length: rows.length - 1 }, () => [formula]));
    }
    const clearFrom = Math.max(3, end + 1);
    if (oldLastRow >= clearFrom) {
      sheet.getRange(clearFrom, column, oldLastRow - clearFrom + 1, 1).clearContent();
    }
  });
  SpreadsheetApp.flush();
}

function tarikReportKeSheet(namaSheet) {
  const ids = RAPOT_REPORT.TARGETS[namaSheet];
  if (!ids) throw new Error('Sheet ' + namaSheet + ' belum punya konfigurasi report.');

  // Ambil dan validasi semua sumber sebelum tab tujuan disentuh.
  const reports = ids.map(id => ({ id, rows: requestReport_(id, RAPOT_REPORT.TOKEN) }));
  const table = buildReportTable_(reports);
  const ss = SpreadsheetApp.openById(PIPE.SS_ID);
  const sheet = ss.getSheetByName(namaSheet) || ss.insertSheet(namaSheet);
  const last = barisTerakhirRapot_(sheet);
  const formulas = {};
  if (sheet.getLastColumn() >= 13 && sheet.getMaxRows() >= 2) {
    const count = Math.min(2, sheet.getMaxColumns() - 12);
    const names = sheet.getRange(1, 13, 1, count).getValues()[0];
    const templates = sheet.getRange(2, 13, 1, count).getFormulasR1C1()[0];
    templates.forEach((formula, index) => {
      const key = reportKey_(String(names[index]));
      if (formula && key === reportKey_(RAPOT_HEADER[12 + index]) && !table.competencyWithData.has(key)) {
        formulas[13 + index] = formula;
      }
    });
  }

  tulisReport_(sheet, table.headers, table.rows, last, formulas);
  Logger.log(namaSheet + ': ' + reports.map(report => report.id + '=' + report.rows.length).join(', '));
  return table.rows.length;
}

// Pertahankan nama fungsi yang dipakai pemanggil rapot sebelumnya.
function updateDataRapot() { return tarikReportKeSheet('RAW'); }
function updateDataRapotRemed() { return tarikReportKeSheet('Remedial'); }
function updateDataRapotRemed2() { return tarikReportKeSheet('Resertifikasi'); }

function updateSlotMap() {
  return isiSlotMapOtomatis(RAPOT_REPORT.SLOT_ID);
}

function isiSlotMapOtomatis(reportId) {
  const props = PropertiesService.getScriptProperties();
  const tokenKey = props.getProperty(RAPOT_REPORT.SLOT_TOKEN)
    ? RAPOT_REPORT.SLOT_TOKEN : RAPOT_REPORT.TOKEN;
  const rows = requestReport_(reportId, tokenKey);
  const headers = ['quiz_id', 'slot', 'topik', 'competency'];
  const data = rows.map((row, index) => {
    const values = new Map(Object.keys(row).map(name => [reportKey_(name), reportText_(row[name])]));
    const quizId = Number(values.get('quiz_id'));
    const slot = Number(values.get('slot'));
    if (!Number.isInteger(quizId) || quizId < 1 || !Number.isInteger(slot) || slot < 1
        || !values.get('topik') || !values.has('competency')) {
      throw new Error('Report ' + reportId + ', baris ' + (index + 1) + ': mapping slot tidak valid.');
    }
    return [quizId, slot, reportValue_(values.get('topik'), 'topik'),
      reportValue_(values.get('competency'), 'competency')];
  });
  if (!data.length) throw new Error('Report ' + reportId + ': mapping slot kosong.');

  const ss = SpreadsheetApp.openById(PIPE.SS_ID);
  const sheet = ss.getSheetByName(RAPOT_REPORT.SLOT_SHEET) || ss.insertSheet(RAPOT_REPORT.SLOT_SHEET);
  tulisReport_(sheet, headers, data, barisTerakhirRapot_(sheet));
  return data.length;
}
