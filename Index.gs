/**
 * ABSENSI ONLINE 1 KELAS/LEMBAGA
 * Backend Google Apps Script - terhubung ke Google Sheet.
 * Jalankan setupAplikasi() sekali sebelum deployment.
 */

const _KT_BRAND = Object.freeze([
  'QXBsaWthc2kgaW5pIGRpYnVhdCBiZXJzYW1h',
  'S2FtcHVzIFR1dG9yaWFs',
  'aHR0cHM6Ly95dDMuZ29vZ2xldXNlcmNvbnRlbnQuY29tL1djRDVJcTNyZGk0T1JidEFvMnBLXzROSF9EVmN3ZWJnUE9RUUthazQtVUQySEY5MVdUZ3plODVKcEFFYWMzcmduQjlUMVFseXJGYz1zOTAwLWMtay1jMHgwMGZmZmZmZi1uby1yag==',
  'aHR0cHM6Ly93d3cueW91dHViZS5jb20vY2hhbm5lbC9VQ2VVb0ZwOXg0OFcxRnJLOU1ha1F0N1E='
]);

const APP = Object.freeze({
  VERSION: '1.2.0',
  TIMEZONE: 'Asia/Jakarta',
  SHEETS: {
    SETTINGS: 'Pengaturan',
    PARTICIPANTS: 'Peserta',
    ATTENDANCE: 'Absensi'
  },
  DEFAULT_PIN: '123456',
  STATUS: ['H', 'S', 'I', 'A']
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Absensi Online')
    .addItem('Siapkan / Perbaiki Aplikasi', 'setupAplikasi')
    .addItem('Ganti PIN Pengelola', 'gantiPinDariSpreadsheet')
    .addSeparator()
    .addItem('Buka Petunjuk Deployment', 'tampilkanPetunjuk')
    .addToUi();
}

function doGet() {
  try {
    getSS_();
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Absensi Online')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    const message = escapeHtml_(error && error.message ? error.message : String(error));
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html><html><head><base target="_top"><meta name="viewport" content="width=device-width,initial-scale=1">
      <style>body{font-family:Arial,sans-serif;background:#f4f7fb;padding:32px;color:#172033}.box{max-width:700px;margin:auto;background:#fff;padding:28px;border-radius:20px;box-shadow:0 18px 50px rgba(15,23,42,.12)}code{background:#eef2ff;padding:3px 7px;border-radius:7px}</style>
      </head><body><div class="box"><h2>Aplikasi belum disiapkan</h2><p>${message}</p><p>Buka Google Sheet sumber, pilih <b>Ekstensi → Apps Script</b>, lalu jalankan fungsi <code>setupAplikasi</code> satu kali.</p></div></body></html>
    `).setTitle('Setup Absensi Online');
  }
}

/** Jalankan sekali dari editor Apps Script yang terikat pada Google Sheet. */
function setupAplikasi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Buka Apps Script melalui Google Sheet agar Spreadsheet ID dapat disimpan.');

  const props = PropertiesService.getScriptProperties();
  props.setProperty('SPREADSHEET_ID', ss.getId());
  ss.setSpreadsheetTimeZone(APP.TIMEZONE);

  setupSettingsSheet_(ss);
  setupParticipantsSheet_(ss);
  setupAttendanceSheet_(ss);
  repairAttendanceDates_(ss);
  setupSecurity_();

  SpreadsheetApp.flush();
  try {
    SpreadsheetApp.getUi().alert(
      'Setup berhasil',
      'Aplikasi siap. PIN awal: 123456. Segera ganti PIN melalui menu “Absensi Online”.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (_) {}
  return 'Setup berhasil. PIN awal: 123456';
}

function tampilkanPetunjuk() {
  SpreadsheetApp.getUi().alert(
    'Deployment',
    '1. Klik Deploy → New deployment.\n2. Pilih Web app.\n3. Execute as: Me.\n4. Who has access: Anyone.\n5. Salin URL /exec.\n6. Buka aplikasi dan ganti PIN awal.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function gantiPinDariSpreadsheet() {
  const ui = SpreadsheetApp.getUi();
  const oldResult = ui.prompt('PIN lama', 'Masukkan PIN pengelola saat ini:', ui.ButtonSet.OK_CANCEL);
  if (oldResult.getSelectedButton() !== ui.Button.OK) return;
  const newResult = ui.prompt('PIN baru', 'Masukkan PIN baru minimal 6 karakter:', ui.ButtonSet.OK_CANCEL);
  if (newResult.getSelectedButton() !== ui.Button.OK) return;
  changePin(oldResult.getResponseText(), newResult.getResponseText());
  ui.alert('PIN berhasil diubah.');
}

function ping() {
  return { ok: true, version: APP.VERSION };
}

function getBootstrap(pin) {
  requirePin_(pin);
  const today = Utilities.formatDate(new Date(), APP.TIMEZONE, 'yyyy-MM-dd');
  const month = today.slice(0, 7);
  return {
    version: APP.VERSION,
    settings: getSettings_(),
    participants: getParticipants_(),
    attendance: getAttendanceByDate_(today),
    dashboard: buildDashboard_(month),
    branding: getProtectedBranding_(),
    today,
    month,
    defaultPin: String(pin) === APP.DEFAULT_PIN
  };
}

function getAttendance(pin, dateText) {
  requirePin_(pin);
  const date = normalizeDate_(dateText);
  return getAttendanceByDate_(date);
}

function saveAttendance(pin, payload) {
  requirePin_(pin);
  if (!payload || !payload.date || !Array.isArray(payload.records)) {
    throw new Error('Data absensi tidak lengkap.');
  }

  const date = normalizeDate_(payload.date);
  const participantMap = Object.fromEntries(getParticipants_(true).map(item => [item.id, item]));
  const cleaned = payload.records
    .map(record => ({
      participantId: sanitizeText_(record.participantId, 80),
      status: sanitizeText_(record.status, 1).toUpperCase(),
      note: sanitizeText_(record.note || '', 250)
    }))
    .filter(record => participantMap[record.participantId] && APP.STATUS.includes(record.status));

  if (!cleaned.length) throw new Error('Tidak ada status absensi yang dapat disimpan.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getSS_().getSheetByName(APP.SHEETS.ATTENDANCE);
    const values = sheet.getDataRange().getValues();
    const rowMap = {};
    for (let i = 1; i < values.length; i++) {
      if (dateKeyFromCell_(values[i][1]) === date) rowMap[String(values[i][2])] = i + 1;
    }

    const now = new Date();
    const appendRows = [];
    cleaned.forEach(record => {
      const participant = participantMap[record.participantId];
      const row = [
        Utilities.getUuid(),
        date,
        participant.id,
        participant.number,
        participant.name,
        record.status,
        record.note,
        now
      ];
      if (rowMap[participant.id]) {
        row[0] = String(values[rowMap[participant.id] - 1][0]) || row[0];
        sheet.getRange(rowMap[participant.id], 1, 1, row.length).setValues([row]);
      } else {
        appendRows.push(row);
      }
    });
    if (appendRows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, appendRows[0].length).setValues(appendRows);
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  return {
    success: true,
    message: `Absensi tanggal ${formatDateIndonesia_(date)} berhasil disimpan.`,
    attendance: getAttendanceByDate_(date),
    dashboard: buildDashboard_(date.slice(0, 7)),
    report: buildMonthlyReport_(date.slice(0, 7))
  };
}

function getMonthlyReport(pin, monthText) {
  requirePin_(pin);
  const month = normalizeMonth_(monthText);
  return buildMonthlyReport_(month);
}

function saveParticipant(pin, participant) {
  requirePin_(pin);
  if (!participant) throw new Error('Data peserta tidak ditemukan.');

  const id = sanitizeText_(participant.id || '', 80);
  const number = sanitizeText_(participant.number || '', 60);
  const name = sanitizeText_(participant.name || '', 120);
  const gender = sanitizeText_(participant.gender || '', 1).toUpperCase();

  if (!name) throw new Error('Nama peserta wajib diisi.');
  if (gender && !['L', 'P'].includes(gender)) throw new Error('Jenis kelamin harus L atau P.');

  const sheet = getSS_().getSheetByName(APP.SHEETS.PARTICIPANTS);
  const values = sheet.getDataRange().getValues();
  const now = new Date();

  if (id) {
    const index = values.findIndex((row, i) => i > 0 && String(row[0]) === id);
    if (index < 0) throw new Error('Peserta yang akan diedit tidak ditemukan.');
    sheet.getRange(index + 1, 2, 1, 5).setValues([[number, name, gender, true, now]]);
  } else {
    sheet.appendRow([Utilities.getUuid(), number, name, gender, true, now]);
  }

  return { success: true, participants: getParticipants_() };
}

function bulkAddParticipants(pin, rawText) {
  requirePin_(pin);
  const text = String(rawText || '').trim();
  if (!text) throw new Error('Tempelkan daftar peserta terlebih dahulu.');

  const lines = text.split(/\r?\n/).map(line => line.trim()).filter(Boolean);
  const rows = [];
  const now = new Date();

  lines.forEach((line, index) => {
    const parts = line.split(/\t|;|,/).map(item => item.trim());
    let number = '';
    let name = '';
    let gender = '';

    if (parts.length >= 2) {
      number = sanitizeText_(parts[0], 60);
      name = sanitizeText_(parts[1], 120);
      gender = sanitizeText_(parts[2] || '', 1).toUpperCase();
    } else {
      name = sanitizeText_(parts[0], 120);
    }

    if (!name) return;
    if (gender && !['L', 'P'].includes(gender)) gender = '';
    rows.push([Utilities.getUuid(), number, name, gender, true, now]);
  });

  if (!rows.length) throw new Error('Daftar peserta tidak dapat dibaca.');
  const sheet = getSS_().getSheetByName(APP.SHEETS.PARTICIPANTS);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return { success: true, added: rows.length, participants: getParticipants_() };
}

function deleteParticipant(pin, participantId) {
  requirePin_(pin);
  const id = sanitizeText_(participantId, 80);
  const sheet = getSS_().getSheetByName(APP.SHEETS.PARTICIPANTS);
  const values = sheet.getDataRange().getValues();
  const index = values.findIndex((row, i) => i > 0 && String(row[0]) === id);
  if (index < 0) throw new Error('Peserta tidak ditemukan.');
  sheet.deleteRow(index + 1);
  return { success: true, participants: getParticipants_() };
}

function saveSettings(pin, settings) {
  requirePin_(pin);
  if (!settings || typeof settings !== 'object') throw new Error('Pengaturan tidak valid.');

  const allowed = [
    'institutionName', 'unitName', 'address', 'logoUrl', 'academicYear',
    'leaderTitle', 'leaderName', 'leaderId',
    'coordinatorTitle', 'coordinatorName', 'coordinatorId',
    'footerText', 'theme'
  ];
  const current = getSettings_();
  allowed.forEach(key => {
    if (Object.prototype.hasOwnProperty.call(settings, key)) {
      current[key] = sanitizeText_(settings[key], key === 'logoUrl' ? 500 : 250);
    }
  });
  writeSettings_(current);
  return { success: true, settings: getSettings_() };
}

function changePin(oldPin, newPin) {
  requirePin_(oldPin);
  const candidate = String(newPin || '').trim();
  if (candidate.length < 6) throw new Error('PIN baru minimal 6 karakter.');
  if (candidate.length > 40) throw new Error('PIN baru maksimal 40 karakter.');
  setPin_(candidate);
  return { success: true };
}

/**
 * Branding kolaborator tidak disimpan di Sheet agar tidak dapat diubah melalui menu aplikasi.
 * Pengkodean dan signature bersifat deterrent untuk edit awam, bukan DRM absolut.
 */
function getProtectedBranding_() {
  const line = decodeBrandValue_(_KT_BRAND[0]);
  const name = decodeBrandValue_(_KT_BRAND[1]);
  const logoUrl = decodeBrandValue_(_KT_BRAND[2]);
  const youtubeUrl = decodeBrandValue_(_KT_BRAND[3]);
  const payload = [line, name, logoUrl, youtubeUrl, APP.VERSION].join('|');
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    payload,
    Utilities.Charset.UTF_8
  );
  const signature = digest.map(byte => ('0' + ((byte + 256) % 256).toString(16)).slice(-2)).join('').slice(0, 24);
  return { line, name, logoUrl, youtubeUrl, signature };
}

function decodeBrandValue_(encoded) {
  return Utilities.newBlob(Utilities.base64Decode(String(encoded || ''))).getDataAsString('UTF-8');
}

// ==================== DATA BUILDERS ====================

function getSettings_() {
  const sheet = getSS_().getSheetByName(APP.SHEETS.SETTINGS);
  const values = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (key) result[key] = String(values[i][1] || '');
  }
  return Object.assign(defaultSettings_(), result);
}

function writeSettings_(settings) {
  const sheet = getSS_().getSheetByName(APP.SHEETS.SETTINGS);
  const rows = Object.keys(defaultSettings_()).map(key => [key, settings[key] == null ? '' : settings[key]]);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 2).setValues([['Kunci', 'Nilai']]);
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  styleHeader_(sheet.getRange(1, 1, 1, 2), '#5b5cf0');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 190);
  sheet.setColumnWidth(2, 420);
}

function getParticipants_(includeInactive) {
  const sheet = getSS_().getSheetByName(APP.SHEETS.PARTICIPANTS);
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues()
    .map(row => ({
      id: String(row[0] || ''),
      number: String(row[1] || ''),
      name: String(row[2] || ''),
      gender: String(row[3] || ''),
      active: row[4] !== false && String(row[4]).toLowerCase() !== 'false'
    }))
    .filter(item => item.id && item.name && (includeInactive || item.active))
    .sort((a, b) => a.name.localeCompare(b.name, 'id'));
}

function getAttendanceByDate_(date) {
  const sheet = getSS_().getSheetByName(APP.SHEETS.ATTENDANCE);
  const map = {};
  if (sheet.getLastRow() < 2) return map;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  values.forEach(row => {
    if (dateKeyFromCell_(row[1]) === date) {
      map[String(row[2])] = {
        id: String(row[0] || ''),
        participantId: String(row[2] || ''),
        status: String(row[5] || ''),
        note: String(row[6] || '')
      };
    }
  });
  return map;
}

function buildDashboard_(month) {
  const report = buildMonthlyReport_(month);
  const today = Utilities.formatDate(new Date(), APP.TIMEZONE, 'yyyy-MM-dd');
  const todayAttendance = getAttendanceByDate_(today);
  const todayCounts = { H: 0, S: 0, I: 0, A: 0, empty: 0 };
  const participants = getParticipants_();
  participants.forEach(participant => {
    const status = todayAttendance[participant.id] && todayAttendance[participant.id].status;
    if (APP.STATUS.includes(status)) todayCounts[status]++;
    else todayCounts.empty++;
  });
  return {
    month,
    participantCount: participants.length,
    attendanceDays: report.attendanceDays,
    todayCounts,
    totals: report.totals,
    daily: report.daily,
    attendanceRate: report.attendanceRate
  };
}

function buildMonthlyReport_(month) {
  const participants = getParticipants_(true);
  const participantMap = Object.fromEntries(participants.map(item => [item.id, item]));
  const sheet = getSS_().getSheetByName(APP.SHEETS.ATTENDANCE);
  const rows = sheet.getLastRow() < 2
    ? []
    : sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues()
      .filter(row => dateKeyFromCell_(row[1]).startsWith(month));

  rows.forEach(row => {
    const participantId = String(row[2] || '');
    if (participantId && !participantMap[participantId]) {
      participantMap[participantId] = {
        id: participantId,
        number: String(row[3] || ''),
        name: String(row[4] || 'Peserta terhapus'),
        gender: '',
        active: false
      };
    }
  });

  const recapMap = {};
  Object.values(participantMap).forEach(participant => {
    recapMap[participant.id] = {
      id: participant.id,
      number: participant.number,
      name: participant.name,
      H: 0, S: 0, I: 0, A: 0,
      recorded: 0,
      rate: 0
    };
  });

  const totals = { H: 0, S: 0, I: 0, A: 0 };
  const dailyMap = {};
  const attendanceDates = new Set();

  rows.forEach(row => {
    const date = dateKeyFromCell_(row[1]);
    const participantId = String(row[2] || '');
    const status = String(row[5] || '').toUpperCase();
    if (!date || !participantId || !APP.STATUS.includes(status) || !recapMap[participantId]) return;

    recapMap[participantId][status]++;
    recapMap[participantId].recorded++;
    totals[status]++;
    attendanceDates.add(date);

    if (!dailyMap[date]) dailyMap[date] = { date, H: 0, S: 0, I: 0, A: 0 };
    dailyMap[date][status]++;
  });

  const recap = Object.values(recapMap)
    .filter(item => item.recorded > 0 || (participantMap[item.id] && participantMap[item.id].active))
    .map(item => {
      item.rate = item.recorded ? Math.round((item.H / item.recorded) * 1000) / 10 : 0;
      return item;
    })
    .sort((a, b) => a.name.localeCompare(b.name, 'id'));

  const totalRecorded = totals.H + totals.S + totals.I + totals.A;
  const attendanceRate = totalRecorded ? Math.round((totals.H / totalRecorded) * 1000) / 10 : 0;
  const daily = Object.values(dailyMap)
    .sort((a, b) => a.date.localeCompare(b.date))
    .map(item => ({
      date: item.date,
      label: item.date.slice(8, 10),
      H: item.H,
      S: item.S,
      I: item.I,
      A: item.A
    }));

  return {
    month,
    monthLabel: formatMonthIndonesia_(month),
    settings: getSettings_(),
    attendanceDays: attendanceDates.size,
    participantCount: recap.length,
    totals,
    totalRecorded,
    attendanceRate,
    daily,
    recap
  };
}

// ==================== SETUP ====================

function setupSettingsSheet_(ss) {
  let sheet = ss.getSheetByName(APP.SHEETS.SETTINGS);
  if (!sheet) sheet = ss.insertSheet(APP.SHEETS.SETTINGS);
  if (sheet.getLastRow() === 0) {
    writeSettings_(defaultSettings_());
  } else {
    const current = getSettingsFromSheet_(sheet);
    writeSettings_(Object.assign(defaultSettings_(), current));
  }
}

function setupParticipantsSheet_(ss) {
  let sheet = ss.getSheetByName(APP.SHEETS.PARTICIPANTS);
  if (!sheet) sheet = ss.insertSheet(APP.SHEETS.PARTICIPANTS);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 6).setValues([['ID', 'Nomor Induk', 'Nama Peserta', 'JK', 'Aktif', 'Diperbarui']]);
  }
  styleHeader_(sheet.getRange(1, 1, 1, 6), '#16a085');
  sheet.setFrozenRows(1);
  [180, 130, 260, 70, 80, 160].forEach((width, i) => sheet.setColumnWidth(i + 1, width));
}

function setupAttendanceSheet_(ss) {
  let sheet = ss.getSheetByName(APP.SHEETS.ATTENDANCE);
  if (!sheet) sheet = ss.insertSheet(APP.SHEETS.ATTENDANCE);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 8).setValues([['ID', 'Tanggal', 'Peserta ID', 'Nomor Induk', 'Nama Peserta', 'Status', 'Keterangan', 'Dicatat Pada']]);
  }
  styleHeader_(sheet.getRange(1, 1, 1, 8), '#f59e0b');
  sheet.setFrozenRows(1);
  [180, 110, 180, 130, 240, 80, 260, 170].forEach((width, i) => sheet.setColumnWidth(i + 1, width));
  // Paksa kolom Tanggal menjadi teks agar nilai YYYY-MM-DD tidak diubah otomatis oleh Google Sheet.
  if (sheet.getMaxRows() > 1) sheet.getRange(2, 2, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
  if (sheet.getMaxRows() > 1) sheet.getRange(2, 8, sheet.getMaxRows() - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
}

/**
 * Memperbaiki data lama yang sudah terlanjur disimpan sebagai objek Date.
 * Fungsi dipanggil otomatis oleh setupAplikasi() dan aman dijalankan berulang kali.
 */
function repairAttendanceDates_(ss) {
  const spreadsheet = ss || getSS_();
  const sheet = spreadsheet.getSheetByName(APP.SHEETS.ATTENDANCE);
  if (!sheet || sheet.getLastRow() < 2) return 0;

  const range = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1);
  const values = range.getValues();
  let repaired = 0;
  const normalized = values.map(row => {
    const original = row[0];
    const key = dateKeyFromCell_(original);
    if (key && (original instanceof Date || String(original).trim() !== key)) repaired++;
    return [key || String(original || '').trim()];
  });

  range.setNumberFormat('@');
  range.setValues(normalized);
  SpreadsheetApp.flush();
  return repaired;
}

function setupSecurity_() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PIN_SALT')) props.setProperty('PIN_SALT', Utilities.getUuid());
  if (!props.getProperty('PIN_HASH')) setPin_(APP.DEFAULT_PIN);
}

function defaultSettings_() {
  return {
    institutionName: 'NAMA SEKOLAH / LEMBAGA',
    unitName: 'Kelas / Kelompok / Unit',
    address: 'Alamat sekolah atau lembaga',
    logoUrl: '',
    academicYear: '2026/2027',
    leaderTitle: 'Kepala Sekolah/Lembaga',
    leaderName: 'Nama Kepala',
    leaderId: '',
    coordinatorTitle: 'Wali Kelas/Guru/Koordinator',
    coordinatorName: 'Nama Koordinator',
    coordinatorId: '',
    footerText: 'Absensi Online',
    theme: 'pelangi'
  };
}

function getSettingsFromSheet_(sheet) {
  const result = {};
  if (sheet.getLastRow() < 2) return result;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(row => {
    const key = String(row[0] || '').trim();
    if (key) result[key] = String(row[1] || '');
  });
  return result;
}

// ==================== SECURITY & HELPERS ====================

function getSS_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SPREADSHEET_ID');
  if (!id) {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      id = active.getId();
      props.setProperty('SPREADSHEET_ID', id);
    }
  }
  if (!id) throw new Error('Spreadsheet belum ditautkan. Jalankan setupAplikasi() dari Google Sheet.');
  return SpreadsheetApp.openById(id);
}

function setPin_(pin) {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PIN_SALT')) props.setProperty('PIN_SALT', Utilities.getUuid());
  props.setProperty('PIN_HASH', hashPin_(pin));
}

function hashPin_(pin) {
  const salt = PropertiesService.getScriptProperties().getProperty('PIN_SALT') || '';
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    `${String(pin)}::${salt}`,
    Utilities.Charset.UTF_8
  );
  return Utilities.base64Encode(bytes);
}

function requirePin_(pin) {
  setupSecurity_();
  const input = String(pin || '');
  const stored = PropertiesService.getScriptProperties().getProperty('PIN_HASH');
  if (!input || hashPin_(input) !== stored) throw new Error('PIN pengelola salah.');
  return true;
}

function styleHeader_(range, color) {
  range.setBackground(color).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
}


/**
 * Mengubah nilai tanggal dari Google Sheet menjadi kunci YYYY-MM-DD.
 * Mendukung objek Date, teks ISO, dan beberapa format tanggal tampilan umum.
 */
function dateKeyFromCell_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, APP.TIMEZONE, 'yyyy-MM-dd');
  }

  const text = String(value == null ? '' : value).trim();
  if (!text) return '';

  const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;

  const localMatch = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (localMatch) {
    return `${localMatch[3]}-${String(localMatch[2]).padStart(2, '0')}-${String(localMatch[1]).padStart(2, '0')}`;
  }

  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, APP.TIMEZONE, 'yyyy-MM-dd');
  }
  return text;
}

function normalizeDate_(value) {
  const text = String(value || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) throw new Error('Format tanggal harus YYYY-MM-DD.');
  const parsed = new Date(`${text}T00:00:00+07:00`);
  if (isNaN(parsed.getTime())) throw new Error('Tanggal tidak valid.');
  return text;
}

function normalizeMonth_(value) {
  const text = String(value || '').trim();
  if (!/^\d{4}-\d{2}$/.test(text)) throw new Error('Format bulan harus YYYY-MM.');
  return text;
}

function sanitizeText_(value, maxLength) {
  return String(value == null ? '' : value)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .trim()
    .slice(0, maxLength || 250);
}

function formatDateIndonesia_(dateText) {
  const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  const [year, month, day] = dateText.split('-').map(Number);
  return `${day} ${months[month - 1]} ${year}`;
}

function formatMonthIndonesia_(monthText) {
  const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  const [year, month] = monthText.split('-').map(Number);
  return `${months[month - 1]} ${year}`;
}

function escapeHtml_(text) {
  return String(text).replace(/[&<>'"]/g, char => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;'
  })[char]);
}
