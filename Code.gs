/**
 * ============================================================
 *  APLIKASI LAPORAN PART BEKAS  (Teknisi -> PIC)
 *  Backend: Google Apps Script + Google Sheets + Google Drive
 * ------------------------------------------------------------
 *  Fitur:
 *   - Teknisi melapor part bekas + foto SPK + foto tiap part
 *   - OCR foto SPK -> No. SPK, No. Polisi, Type Mobil terisi otomatis
 *   - Serah terima ke PIC (cek per part, PIN, tanda tangan)
 *   - Indikator kelengkapan foto & status serah terima
 *   - Semua data tersimpan otomatis ke Spreadsheet, foto ke Drive
 * ============================================================
 */

/** ====================== KONFIGURASI ====================== */
const CONFIG = {
  APP_NAME: 'Laporan Part Bekas',
  ROOT_FOLDER_NAME: 'Laporan Part Bekas - Foto',
  TIMEZONE: 'Asia/Jakarta',
  OCR_LANGUAGE: 'id',
  // Kirim email ke PIC saat ada laporan baru (isi kolom Email di Master PIC)
  NOTIFY_PIC_EMAIL: true,
  // (Opsional) Pola nomor SPK khusus bengkel Anda, contoh: 'SPK[\\/\\-]?\\d{4,}'
  // Kosongkan jika ingin memakai deteksi umum.
  SPK_CUSTOM_REGEX: '',
  MAX_IMAGE_MB: 8
};

const SHEETS = {
  LAPORAN: 'Laporan',
  DETAIL: 'Detail Part',
  TEKNISI: 'Master Teknisi',
  PIC: 'Master PIC',
  TYPE: 'Master Type Mobil',
  ASURANSI: 'Master Asuransi'
};

const HEADERS = {
  LAPORAN: [
    'ID Laporan', 'Tanggal Lapor', 'Nama Teknisi', 'PIC Tujuan',
    'No. SPK', 'No. Polisi', 'Type Mobil',
    'Jumlah Part', 'Part Berfoto', 'Indikator Foto', 'Part Diterima PIC',
    'Status Serah Terima', 'Indikator',
    'Diserahkan Oleh', 'Diterima Oleh', 'Tanggal Diterima',
    'Catatan Teknisi', 'Catatan PIC',
    'Foto SPK', 'Folder Foto', 'TTD Teknisi', 'TTD PIC', 'Teks OCR SPK',
    // --- Pengantaran ke asuransi oleh PIC ---
    'Status Asuransi', 'Nama Asuransi', 'No. Klaim',
    'Tanggal Planning Diantar ke Asuransi', 'Tanggal Aktual Diantar ke Asuransi', 'Selisih Hari (Aktual - Planning)', 'Ketepatan Waktu',
    'Diantar Oleh (PIC)', 'Penerima di Asuransi', 'Bukti Tanda Terima Asuransi', 'Catatan Pengantaran',
    'Waktu Input Planning', 'Waktu Input Pengantaran'
  ],
  DETAIL: [
    'ID Laporan', 'No', 'Nama Part', 'No. Part', 'Qty', 'Kondisi',
    'Foto Part', 'Cek PIC', 'Catatan PIC', 'Waktu Cek'
  ],
  TEKNISI: ['Nama Teknisi', 'Aktif (Y/N)'],
  PIC: ['Nama PIC', 'Email', 'PIN', 'Aktif (Y/N)'],
  TYPE: ['Type Mobil'],
  ASURANSI: ['Nama Asuransi']
};

const STATUS = {
  MENUNGGU: 'Menunggu PIC',
  DITERIMA: 'Diterima',
  SEBAGIAN: 'Diterima Sebagian',
  DITOLAK: 'Ditolak'
};

const ASURANSI_STATUS = {
  BELUM: 'Belum Dijadwalkan',
  TERJADWAL: 'Terjadwal',
  SUDAH: 'Sudah Diantar'
};

// Nama kolom lama -> baru (agar spreadsheet versi sebelumnya otomatis menyesuaikan saat setup dijalankan ulang)
const HEADER_RENAMES = {
  'Tanggal Diantar ke Asuransi': 'Tanggal Aktual Diantar ke Asuransi'
};

/** ====================== WEB APP ====================== */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(CONFIG.APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** ====================== SETUP (jalankan 1x) ====================== */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Jalankan setup() dari Apps Script yang terikat (Extensions > Apps Script) pada Spreadsheet.');
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SS_ID', ss.getId());
  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);

  const lap = ensureSheet_(ss, SHEETS.LAPORAN, HEADERS.LAPORAN, HEADER_RENAMES);
  // Status lama "Belum Diantar" -> "Belum Dijadwalkan"
  const stCol = col_(HEADERS.LAPORAN, 'Status Asuransi');
  if (lap.getLastRow() > 1) {
    const r = lap.getRange(2, stCol, lap.getLastRow() - 1, 1);
    r.setValues(r.getValues().map(v => [v[0] === 'Belum Diantar' ? ASURANSI_STATUS.BELUM : v[0]]));
  }
  const det = ensureSheet_(ss, SHEETS.DETAIL, HEADERS.DETAIL);
  const tek = ensureSheet_(ss, SHEETS.TEKNISI, HEADERS.TEKNISI);
  const pic = ensureSheet_(ss, SHEETS.PIC, HEADERS.PIC);
  const typ = ensureSheet_(ss, SHEETS.TYPE, HEADERS.TYPE);
  const asr = ensureSheet_(ss, SHEETS.ASURANSI, HEADERS.ASURANSI);
  if (asr.getLastRow() < 2) {
    const list = ['Asuransi Astra (Garda Oto)', 'Sinar Mas', 'ACA', 'Adira Insurance (Zurich)', 'Allianz', 'BCA Insurance',
      'MAG', 'Tugu Insurance', 'Jasindo', 'Raksa Pratikara', 'Mandiri AXA', 'Sompo'];
    asr.getRange(2, 1, list.length, 1).setValues(list.map(t => [t]));
  }

  if (tek.getLastRow() < 2) tek.getRange(2, 1, 3, 2).setValues([['Budi', 'Y'], ['Andi', 'Y'], ['Rizky', 'Y']]);
  if (pic.getLastRow() < 2) pic.getRange(2, 1, 2, 4).setValues([['Pak Joko', '', '1234', 'Y'], ['Bu Sari', '', '5678', 'Y']]);
  pic.getRange('C:C').setNumberFormat('@'); // PIN sebagai teks
  if (typ.getLastRow() < 2) {
    const types = ['Avanza', 'Xenia', 'Innova', 'Fortuner', 'Rush', 'Terios', 'Agya', 'Ayla', 'Calya', 'Sigra',
      'Brio', 'Jazz', 'HR-V', 'CR-V', 'Mobilio', 'BR-V', 'Xpander', 'Pajero Sport', 'Ertiga', 'XL7',
      'Carry', 'APV', 'Grand Max', 'Hilux', 'Livina', 'Almaz', 'Confero', 'Raize', 'Rocky', 'Veloz', 'Yaris', 'Corolla', 'Camry', 'L300'];
    typ.getRange(2, 1, types.length, 1).setValues(types.map(t => [t]));
  }

  // Format kolom & indikator warna
  lap.setColumnWidths(1, HEADERS.LAPORAN.length, 140);
  lap.setColumnWidth(col_(HEADERS.LAPORAN, 'Teks OCR SPK'), 260);
  lap.getRange(2, col_(HEADERS.LAPORAN, 'Tanggal Planning Diantar ke Asuransi'), 2000, 2).setNumberFormat('dd/MM/yyyy');
  const asrFirst = col_(HEADERS.LAPORAN, 'Status Asuransi');
  lap.getRange(1, asrFirst, 1, HEADERS.LAPORAN.length - asrFirst + 1).setBackground('#1d4f8f');
  applyConditionalFormat_(lap);
  det.setColumnWidths(1, HEADERS.DETAIL.length, 130);
  applyDetailFormat_(det);

  // Hapus Sheet1 bawaan jika kosong
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Lembar1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);

  getRootFolder_();
  SpreadsheetApp.getUi && safeAlert_('Setup selesai ✅\n\nLangkah berikut: Deploy > New deployment > Web app.');
}

function safeAlert_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

function ensureSheet_(ss, name, headers, renames) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  // Migrasi aman: ganti nama kolom lama & sisipkan kolom baru di posisinya tanpa menggeser data lama
  const lc = sh.getLastColumn();
  if (lc > 0) {
    let cur = sh.getRange(1, 1, 1, lc).getValues()[0].map(String);
    if (cur.some(Boolean)) {
      if (renames) cur = cur.map(h => renames[h] || h);
      sh.getRange(1, 1, 1, cur.length).setValues([cur]);
      headers.forEach((h, i) => {
        if (cur.indexOf(h) >= 0) return;
        const pos = i === 0 ? 0 : cur.indexOf(headers[i - 1]) + 1; // indeks 0-based
        if (pos < cur.length) sh.insertColumnBefore(pos + 1);
        cur.splice(pos, 0, h);
        sh.getRange(1, pos + 1).setValue(h);
      });
    }
  }
  const rng = sh.getRange(1, 1, 1, headers.length);
  rng.setValues([headers]).setFontWeight('bold').setBackground('#0f2d52').setFontColor('#ffffff')
    .setVerticalAlignment('middle').setWrap(true);
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 36);
  return sh;
}

function applyConditionalFormat_(sh) {
  const H = HEADERS.LAPORAN;
  const n = 2000;
  const fotoRng = sh.getRange(2, col_(H, 'Indikator Foto'), n, 1);
  const indRng = sh.getRange(2, col_(H, 'Indikator'), n, 1);
  const stRng = sh.getRange(2, col_(H, 'Status Serah Terima'), n, 1);
  const asRng = sh.getRange(2, col_(H, 'Status Asuransi'), n, 1);
  const rule = (rng, txt, bg, fg) => SpreadsheetApp.newConditionalFormatRule()
    .whenTextStartsWith(txt).setBackground(bg).setFontColor(fg).setRanges([rng]).build();
  const ruleEq = (rng, txt, bg, fg) => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(txt).setBackground(bg).setFontColor(fg).setRanges([rng]).build();
  sh.setConditionalFormatRules([
    rule(fotoRng, '✅', '#d9f5e3', '#11643a'),
    rule(fotoRng, '❌', '#fde0e0', '#9b1c1c'),
    rule(indRng, '🔵', '#dbeafe', '#1e3a8a'),
    rule(indRng, '🟣', '#ede9fe', '#5b21b6'),
    rule(indRng, '🟢', '#d9f5e3', '#11643a'),
    rule(indRng, '🟡', '#fff4cc', '#7a5b00'),
    rule(indRng, '🟠', '#ffe6cc', '#8a4300'),
    rule(indRng, '🔴', '#fde0e0', '#9b1c1c'),
    ruleEq(stRng, STATUS.DITERIMA, '#d9f5e3', '#11643a'),
    ruleEq(stRng, STATUS.MENUNGGU, '#fff4cc', '#7a5b00'),
    ruleEq(stRng, STATUS.SEBAGIAN, '#ffe6cc', '#8a4300'),
    ruleEq(stRng, STATUS.DITOLAK, '#fde0e0', '#9b1c1c'),
    ruleEq(asRng, ASURANSI_STATUS.SUDAH, '#dbeafe', '#1e3a8a'),
    ruleEq(asRng, ASURANSI_STATUS.TERJADWAL, '#ede9fe', '#5b21b6'),
    ruleEq(asRng, ASURANSI_STATUS.BELUM, '#fff4cc', '#7a5b00'),
    // Planning sudah lewat tapi belum diantar -> merah
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($' + colLetter_(col_(H, 'Status Asuransi')) + '2="' + ASURANSI_STATUS.TERJADWAL + '",$' +
        colLetter_(col_(H, 'Tanggal Planning Diantar ke Asuransi')) + '2<>"",$' + colLetter_(col_(H, 'Tanggal Planning Diantar ke Asuransi')) + '2<TODAY())')
      .setBackground('#fde0e0').setFontColor('#9b1c1c')
      .setRanges([sh.getRange(2, col_(H, 'Tanggal Planning Diantar ke Asuransi'), n, 1)]).build(),
    rule(sh.getRange(2, col_(H, 'Ketepatan Waktu'), n, 1), '✅', '#d9f5e3', '#11643a'),
    rule(sh.getRange(2, col_(H, 'Ketepatan Waktu'), n, 1), '⚠️', '#fde0e0', '#9b1c1c')
  ]);
}

function applyDetailFormat_(sh) {
  const rng = sh.getRange(2, col_(HEADERS.DETAIL, 'Cek PIC'), 5000, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('✔').setBackground('#d9f5e3').setFontColor('#11643a').setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('✖').setBackground('#fde0e0').setFontColor('#9b1c1c').setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('⏳').setBackground('#fff4cc').setFontColor('#7a5b00').setRanges([rng]).build()
  ]);
}

/** ====================== API UNTUK FRONTEND ====================== */

/** Data awal: daftar teknisi, PIC, type mobil */
function getInitData() {
  const ss = ss_();
  const tek = readTable_(ss.getSheetByName(SHEETS.TEKNISI)).rows
    .filter(r => String(r['Aktif (Y/N)']).toUpperCase() !== 'N' && r['Nama Teknisi'])
    .map(r => String(r['Nama Teknisi']).trim());
  const pic = readTable_(ss.getSheetByName(SHEETS.PIC)).rows
    .filter(r => String(r['Aktif (Y/N)']).toUpperCase() !== 'N' && r['Nama PIC'])
    .map(r => String(r['Nama PIC']).trim());
  const asuransi = readTable_(ss.getSheetByName(SHEETS.ASURANSI)).rows
    .map(r => String(r['Nama Asuransi'] || '').trim()).filter(Boolean);
  return { appName: CONFIG.APP_NAME, teknisi: tek, pic: pic, types: getTypes_(), asuransi: asuransi };
}

/** Upload foto SPK lalu OCR -> kembalikan data terdeteksi */
function scanSpk(dataUrl) {
  const blob = dataUrlToBlob_(dataUrl, 'SPK_' + stamp_() + '.jpg');
  const file = getTempFolder_().createFile(blob);
  let text = '';
  let ocrError = '';
  try {
    text = ocrImage_(blob);
  } catch (e) {
    ocrError = 'OCR gagal: ' + e.message + ' (pastikan layanan Drive API sudah diaktifkan)';
  }
  const parsed = parseSpk_(text);
  parsed.fileId = file.getId();
  parsed.text = text.slice(0, 3000);
  parsed.ocrError = ocrError;
  parsed.duplicate = parsed.spk ? findDuplicateSpk_(parsed.spk) : null;
  return parsed;
}

/** Cek ulang duplikat SPK (dipanggil saat teknisi mengubah No. SPK manual) */
function checkSpk(spk) {
  return findDuplicateSpk_(spk);
}

/** Upload satu foto (part / tanda tangan) ke folder sementara */
function uploadImage(dataUrl, label) {
  const name = (label || 'FOTO').replace(/[^\w\-]+/g, '_') + '_' + stamp_() + '.jpg';
  const blob = dataUrlToBlob_(dataUrl, name);
  const file = getTempFolder_().createFile(blob);
  return { fileId: file.getId() };
}

/** Simpan laporan teknisi */
function submitReport(p) {
  // ---- Validasi ----
  const init = getInitData();
  if (!p || !p.teknisi || init.teknisi.indexOf(p.teknisi) < 0) throw new Error('Nama teknisi tidak valid.');
  if (!p.pic || init.pic.indexOf(p.pic) < 0) throw new Error('PIC tujuan tidak valid.');
  ['spk', 'nopol', 'type'].forEach(k => { if (!p[k] || !String(p[k]).trim()) throw new Error('Data SPK belum lengkap: ' + k.toUpperCase()); });
  if (!p.spkFileId) throw new Error('Foto SPK wajib diunggah.');
  if (!p.parts || !p.parts.length) throw new Error('Minimal 1 part bekas.');
  p.parts.forEach((pt, i) => {
    if (!pt.nama) throw new Error('Nama part #' + (i + 1) + ' belum diisi.');
    if (!pt.fileId) throw new Error('Foto part #' + (i + 1) + ' (' + pt.nama + ') belum ada.');
  });
  if (!p.ttdTeknisi) throw new Error('Tanda tangan teknisi wajib diisi.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = ss_();
    const lap = ss.getSheetByName(SHEETS.LAPORAN);
    const det = ss.getSheetByName(SHEETS.DETAIL);
    const id = nextId_(lap);
    const now = new Date();
    const spk = String(p.spk).trim().toUpperCase();
    const nopol = formatNopol_(p.nopol);

    // ---- Folder & pindah file ----
    const month = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM');
    const folder = getOrCreate_(getOrCreate_(getRootFolder_(), month), id + '_' + spk.replace(/[^\w\-]+/g, '_'));
    const spkFile = moveFile_(p.spkFileId, folder, 'SPK_' + spk);
    const ttdTek = folder.createFile(dataUrlToBlob_(p.ttdTeknisi, 'TTD_Teknisi_' + id + '.png'));

    const detailRows = p.parts.map((pt, i) => {
      const f = moveFile_(pt.fileId, folder, 'Part' + (i + 1) + '_' + pt.nama);
      return [id, i + 1, String(pt.nama).trim(), String(pt.nopart || '').trim(), Number(pt.qty) || 1,
        pt.kondisi || '', fileUrl_(f), '⏳ Belum dicek', '', ''];
    });

    const jml = detailRows.length;
    const berfoto = detailRows.filter(r => r[6]).length;
    const ind = indikator_(jml, berfoto, !!spkFile, STATUS.MENUNGGU, 0);

    const rowObj = {
      'ID Laporan': id,
      'Tanggal Lapor': now,
      'Nama Teknisi': p.teknisi,
      'PIC Tujuan': p.pic,
      'No. SPK': spk,
      'No. Polisi': nopol,
      'Type Mobil': String(p.type).trim(),
      'Jumlah Part': jml,
      'Part Berfoto': berfoto,
      'Indikator Foto': ind.foto,
      'Part Diterima PIC': 0,
      'Status Serah Terima': STATUS.MENUNGGU,
      'Indikator': ind.overall,
      'Diserahkan Oleh': p.teknisi,
      'Diterima Oleh': '',
      'Tanggal Diterima': '',
      'Catatan Teknisi': p.catatan || '',
      'Catatan PIC': '',
      'Foto SPK': fileUrl_(spkFile),
      'Folder Foto': folder.getUrl(),
      'TTD Teknisi': fileUrl_(ttdTek),
      'TTD PIC': '',
      'Teks OCR SPK': String(p.ocrText || '').slice(0, 2000)
    };
    lap.appendRow(HEADERS.LAPORAN.map(h => rowObj[h] === undefined ? '' : rowObj[h]));
    lap.getRange(lap.getLastRow(), col_(HEADERS.LAPORAN, 'Tanggal Lapor')).setNumberFormat('dd/MM/yyyy HH:mm');
    det.getRange(det.getLastRow() + 1, 1, detailRows.length, HEADERS.DETAIL.length).setValues(detailRows);

    notifyPic_(p.pic, rowObj);
    return { id: id, indikator: ind.overall };
  } finally {
    lock.releaseLock();
  }
}

/** PIC login (cek PIN) */
function loginPic(pic, pin) {
  verifyPic_(pic, pin);
  return { ok: true, pic: pic };
}

/** Daftar laporan yang menunggu serah terima untuk PIC */
function getPending(pic, pin) {
  verifyPic_(pic, pin);
  const rows = readTable_(ss_().getSheetByName(SHEETS.LAPORAN)).rows;
  return rows.filter(r => r['PIC Tujuan'] === pic && r['Status Serah Terima'] === STATUS.MENUNGGU)
    .reverse().map(summary_);
}

/** Detail laporan + daftar part */
function getReport(id) {
  const ss = ss_();
  const lapRow = readTable_(ss.getSheetByName(SHEETS.LAPORAN)).rows.find(r => r['ID Laporan'] === id);
  if (!lapRow) throw new Error('Laporan ' + id + ' tidak ditemukan.');
  const parts = readTable_(ss.getSheetByName(SHEETS.DETAIL)).rows
    .filter(r => r['ID Laporan'] === id)
    .map(r => ({
      no: r['No'], nama: r['Nama Part'], nopart: r['No. Part'], qty: r['Qty'], kondisi: r['Kondisi'],
      fotoId: idFromUrl_(r['Foto Part']), cek: r['Cek PIC'], catatan: r['Catatan PIC'], waktu: r['Waktu Cek']
    }));
  const s = summary_(lapRow);
  s.catatanTeknisi = lapRow['Catatan Teknisi'];
  s.catatanPic = lapRow['Catatan PIC'];
  s.spkFotoId = idFromUrl_(lapRow['Foto SPK']);
  s.ttdTeknisiId = idFromUrl_(lapRow['TTD Teknisi']);
  s.ttdPicId = idFromUrl_(lapRow['TTD PIC']);
  s.folderUrl = lapRow['Folder Foto'];
  s.buktiAsuransiId = idFromUrl_(lapRow['Bukti Tanda Terima Asuransi']);
  s.catatanAsuransi = lapRow['Catatan Pengantaran'] || '';
  s.parts = parts;
  return s;
}

/** PIC menyelesaikan serah terima */
function confirmHandover(p) {
  verifyPic_(p.pic, p.pin);
  if (!p.id) throw new Error('ID laporan kosong.');
  const decision = p.decision === 'tolak' ? 'tolak' : 'terima';
  if (decision === 'tolak' && !String(p.catatan || '').trim()) throw new Error('Alasan penolakan wajib diisi.');
  if (decision === 'terima' && !p.ttdPic) throw new Error('Tanda tangan PIC wajib diisi.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = ss_();
    const lap = ss.getSheetByName(SHEETS.LAPORAN);
    const det = ss.getSheetByName(SHEETS.DETAIL);
    const H = HEADERS.LAPORAN;
    const t = readTable_(lap);
    const row = t.rows.find(r => r['ID Laporan'] === p.id);
    if (!row) throw new Error('Laporan tidak ditemukan.');
    if (row['PIC Tujuan'] !== p.pic) throw new Error('Laporan ini bukan untuk PIC ' + p.pic + '.');
    if (row['Status Serah Terima'] !== STATUS.MENUNGGU) throw new Error('Laporan sudah diproses: ' + row['Status Serah Terima']);

    const now = new Date();
    const nowStr = Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm');
    const checks = {};
    (p.checks || []).forEach(c => { checks[String(c.no)] = c; });

    // Update detail part
    const dt = readTable_(det);
    const parts = dt.rows.filter(r => r['ID Laporan'] === p.id);
    let diterima = 0;
    parts.forEach(r => {
      const c = checks[String(r['No'])];
      if (decision === 'terima' && !c) throw new Error('Part #' + r['No'] + ' belum dicek.');
      const ok = decision === 'terima' && c && c.sesuai === true;
      if (ok) diterima++;
      const cek = decision === 'tolak' ? '✖ Ditolak' : (ok ? '✔ Sesuai & Diterima' : '✖ Tidak Sesuai');
      det.getRange(r._row, col_(HEADERS.DETAIL, 'Cek PIC'), 1, 3)
        .setValues([[cek, (c && c.catatan) || '', nowStr]]);
    });

    let status;
    if (decision === 'tolak') status = STATUS.DITOLAK;
    else status = diterima === parts.length ? STATUS.DITERIMA : STATUS.SEBAGIAN;

    let ttdUrl = '';
    if (p.ttdPic) {
      const folder = DriveApp.getFolderById(idFromUrl_(row['Folder Foto']) || getRootFolder_().getId());
      ttdUrl = fileUrl_(folder.createFile(dataUrlToBlob_(p.ttdPic, 'TTD_PIC_' + p.id + '.png')));
    }
    const jml = Number(row['Jumlah Part']) || parts.length;
    const ind = indikator_(jml, Number(row['Part Berfoto']) || 0, !!row['Foto SPK'], status, diterima);

    const upd = {
      'Part Diterima PIC': diterima,
      'Status Serah Terima': status,
      'Indikator': ind.overall,
      'Indikator Foto': ind.foto,
      'Diterima Oleh': decision === 'terima' ? p.pic : '',
      'Tanggal Diterima': now,
      'Catatan PIC': p.catatan || '',
      'TTD PIC': ttdUrl,
      'Status Asuransi': status === STATUS.DITOLAK ? '-' : ASURANSI_STATUS.BELUM
    };
    Object.keys(upd).forEach(k => lap.getRange(row._row, col_(H, k)).setValue(upd[k]));
    lap.getRange(row._row, col_(H, 'Tanggal Diterima')).setNumberFormat('dd/MM/yyyy HH:mm');
    return { id: p.id, status: status, indikator: ind.overall, diterima: diterima, jumlah: jml };
  } finally {
    lock.releaseLock();
  }
}

/** Daftar laporan yang sudah diterima PIC dan belum diantar ke asuransi
 *  (belum dijadwalkan + sudah dijadwalkan), urut tanggal planning terdekat */
function getToDeliver(pic, pin) {
  verifyPic_(pic, pin);
  const rows = readTable_(ss_().getSheetByName(SHEETS.LAPORAN)).rows;
  return rows.filter(r => r['Diterima Oleh'] === pic &&
      (r['Status Serah Terima'] === STATUS.DITERIMA || r['Status Serah Terima'] === STATUS.SEBAGIAN) &&
      r['Status Asuransi'] !== ASURANSI_STATUS.SUDAH)
    .map(summary_)
    .sort((x, y) => (x.isoPlan || '9999') < (y.isoPlan || '9999') ? -1 : 1);
}

/** LANGKAH 1 — PIC membuat / mengubah planning tanggal antar ke asuransi */
function planDelivery(p) {
  verifyPic_(p.pic, p.pin);
  const asuransi = String(p.asuransi || '').trim();
  if (!asuransi) throw new Error('Nama asuransi wajib diisi.');
  const plan = parseDateInput_(p.tanggalPlan, 'Tanggal planning');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const lap = ss_().getSheetByName(SHEETS.LAPORAN);
    const H = HEADERS.LAPORAN;
    const row = getDeliverableRow_(lap, p);
    const tglTerima = parseSheetDate_(row['Tanggal Diterima']);
    if (tglTerima && dayDiff_(tglTerima, plan) < 0)
      throw new Error('Tanggal planning tidak boleh sebelum tanggal diterima PIC (' + dateOnly_(row['Tanggal Diterima']) + ').');

    const lama = dateOnly_(row['Tanggal Planning Diantar ke Asuransi']);
    const baru = Utilities.formatDate(plan, CONFIG.TIMEZONE, 'dd/MM/yyyy');
    let catatan = String(row['Catatan Pengantaran'] || '');
    const tambahan = String(p.catatan || '').trim();
    if (lama && lama !== baru) catatan = (catatan ? catatan + ' | ' : '') + 'Reschedule ' + lama + ' → ' + baru + (tambahan ? ': ' + tambahan : '');
    else if (tambahan) catatan = (catatan ? catatan + ' | ' : '') + tambahan;

    const upd = {
      'Status Asuransi': ASURANSI_STATUS.TERJADWAL,
      'Nama Asuransi': asuransi,
      'No. Klaim': String(p.noKlaim || '').trim(),
      'Tanggal Planning Diantar ke Asuransi': plan,
      'Catatan Pengantaran': catatan,
      'Waktu Input Planning': new Date(),
      'Indikator': '🟣 Terjadwal ke Asuransi - Planning ' + baru
    };
    Object.keys(upd).forEach(k => lap.getRange(row._row, col_(H, k)).setValue(upd[k]));
    lap.getRange(row._row, col_(H, 'Tanggal Planning Diantar ke Asuransi')).setNumberFormat('dd/MM/yyyy');
    lap.getRange(row._row, col_(H, 'Waktu Input Planning')).setNumberFormat('dd/MM/yyyy HH:mm');
    return { id: row['ID Laporan'], indikator: upd['Indikator'], reschedule: !!(lama && lama !== baru) };
  } finally {
    lock.releaseLock();
  }
}

/** LANGKAH 2 — PIC mencatat tanggal AKTUAL part diantar ke asuransi */
function deliverToInsurance(p) {
  verifyPic_(p.pic, p.pin);
  const aktual = parseDateInput_(p.tanggal, 'Tanggal aktual');
  if (dayDiff_(todayDate_(), aktual) > 0) throw new Error('Tanggal aktual tidak boleh melebihi hari ini.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const lap = ss_().getSheetByName(SHEETS.LAPORAN);
    const H = HEADERS.LAPORAN;
    const row = getDeliverableRow_(lap, p);
    if (row['Status Asuransi'] !== ASURANSI_STATUS.TERJADWAL)
      throw new Error('Buat planning tanggal antar terlebih dahulu sebelum mencatat tanggal aktual.');
    const plan = parseSheetDate_(row['Tanggal Planning Diantar ke Asuransi']);
    if (!plan) throw new Error('Tanggal planning tidak ditemukan.');
    const tglTerima = parseSheetDate_(row['Tanggal Diterima']);
    if (tglTerima && dayDiff_(tglTerima, aktual) < 0)
      throw new Error('Tanggal aktual tidak boleh sebelum tanggal diterima PIC (' + dateOnly_(row['Tanggal Diterima']) + ').');

    let buktiUrl = '';
    if (p.buktiFileId) {
      const folder = DriveApp.getFolderById(idFromUrl_(row['Folder Foto']) || getRootFolder_().getId());
      buktiUrl = fileUrl_(moveFile_(p.buktiFileId, folder, 'Bukti_Asuransi_' + row['ID Laporan']));
    }
    const selisih = dayDiff_(plan, aktual); // + = terlambat
    const ketepatan = ketepatan_(selisih);
    const diterima = Number(row['Part Diterima PIC']) || 0;
    const jml = Number(row['Jumlah Part']) || 0;
    const tglStr = Utilities.formatDate(aktual, CONFIG.TIMEZONE, 'dd/MM/yyyy');
    const catatanLama = String(row['Catatan Pengantaran'] || '');
    const tambahan = String(p.catatan || '').trim();

    const upd = {
      'Status Asuransi': ASURANSI_STATUS.SUDAH,
      'Tanggal Aktual Diantar ke Asuransi': aktual,
      'Selisih Hari (Aktual - Planning)': selisih,
      'Ketepatan Waktu': ketepatan,
      'Diantar Oleh (PIC)': p.pic,
      'Penerima di Asuransi': String(p.penerima || '').trim(),
      'Bukti Tanda Terima Asuransi': buktiUrl,
      'Catatan Pengantaran': tambahan ? (catatanLama ? catatanLama + ' | ' : '') + tambahan : catatanLama,
      'Waktu Input Pengantaran': new Date(),
      'Indikator': '🔵 Sudah Diantar ke Asuransi ' + tglStr + ' (' + diterima + '/' + jml + ' part) - ' + ketepatan.replace(/^\S+\s/, '')
    };
    Object.keys(upd).forEach(k => lap.getRange(row._row, col_(H, k)).setValue(upd[k]));
    lap.getRange(row._row, col_(H, 'Tanggal Aktual Diantar ke Asuransi')).setNumberFormat('dd/MM/yyyy');
    lap.getRange(row._row, col_(H, 'Waktu Input Pengantaran')).setNumberFormat('dd/MM/yyyy HH:mm');
    return { id: row['ID Laporan'], indikator: upd['Indikator'], ketepatan: ketepatan, selisih: selisih };
  } finally {
    lock.releaseLock();
  }
}

function getDeliverableRow_(lap, p) {
  if (!p.id) throw new Error('ID laporan kosong.');
  const row = readTable_(lap).rows.find(r => r['ID Laporan'] === p.id);
  if (!row) throw new Error('Laporan tidak ditemukan.');
  if (row['Diterima Oleh'] !== p.pic) throw new Error('Hanya PIC penerima (' + row['Diterima Oleh'] + ') yang dapat mengatur pengantaran part ini.');
  if (row['Status Serah Terima'] !== STATUS.DITERIMA && row['Status Serah Terima'] !== STATUS.SEBAGIAN)
    throw new Error('Part belum diterima PIC, belum bisa diantar ke asuransi.');
  if (row['Status Asuransi'] === ASURANSI_STATUS.SUDAH) throw new Error('Part laporan ini sudah tercatat diantar ke asuransi.');
  return row;
}

function ketepatan_(selisih) {
  if (selisih < 0) return '✅ Lebih Cepat ' + (-selisih) + ' hari';
  if (selisih === 0) return '✅ Tepat Waktu';
  return '⚠️ Terlambat ' + selisih + ' hari';
}

function parseDateInput_(s, label) {
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error((label || 'Tanggal') + ' wajib diisi.');
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
}

function parseSheetDate_(s) {
  const m = String(s || '').match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  return m ? new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), 12, 0, 0) : null;
}

/** Selisih hari kalender (b - a) */
function dayDiff_(a, b) {
  const u = d => Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
  return Math.round((u(b) - u(a)) / 86400000);
}

function todayDate_() {
  return parseDateInput_(Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd'));
}

function dateOnly_(s) {
  const m = String(s || '').match(/^\d{2}\/\d{2}\/\d{4}/);
  return m ? m[0] : '';
}

function colLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/** Riwayat + ringkasan untuk dashboard */
function getDashboard(q) {
  const rows = readTable_(ss_().getSheetByName(SHEETS.LAPORAN)).rows.reverse();
  const counts = { total: rows.length, menunggu: 0, diterima: 0, sebagian: 0, ditolak: 0,
    belumJadwal: 0, terjadwal: 0, lewatJadwal: 0, diantar: 0, tepatWaktu: 0, terlambat: 0 };
  rows.forEach(r => {
    const s = r['Status Serah Terima'];
    const sm = summary_(r);
    if (sm.statusAsuransi === ASURANSI_STATUS.SUDAH) {
      counts.diantar++;
      if (String(r['Ketepatan Waktu']).indexOf('✅') === 0) counts.tepatWaktu++; else counts.terlambat++;
    } else if (sm.statusAsuransi === ASURANSI_STATUS.TERJADWAL) {
      counts.terjadwal++;
      if (sm.sisaHari < 0) counts.lewatJadwal++;
    } else if (s === STATUS.DITERIMA || s === STATUS.SEBAGIAN) counts.belumJadwal++;
    if (s === STATUS.MENUNGGU) counts.menunggu++;
    else if (s === STATUS.DITERIMA) counts.diterima++;
    else if (s === STATUS.SEBAGIAN) counts.sebagian++;
    else if (s === STATUS.DITOLAK) counts.ditolak++;
  });
  const query = String(q || '').trim().toUpperCase().replace(/\s+/g, '');
  const list = rows.filter(r => {
    if (!query) return true;
    return [r['ID Laporan'], r['No. SPK'], r['No. Polisi'], r['Nama Teknisi'], r['PIC Tujuan'], r['Type Mobil'], r['Nama Asuransi'], r['No. Klaim']]
      .some(v => String(v).toUpperCase().replace(/\s+/g, '').indexOf(query) >= 0);
  }).slice(0, 100).map(summary_);
  return { counts: counts, list: list };
}

/** Ambil foto (hanya dari folder aplikasi) sebagai data URL */
function getImage(fileId) {
  if (!fileId) return '';
  const f = DriveApp.getFileById(fileId);
  if (!isInRoot_(f)) throw new Error('Akses foto ditolak.');
  const b = f.getBlob();
  return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
}

/** ====================== OCR & PARSING SPK ====================== */

function ocrImage_(blob) {
  let docId;
  try {
    // Drive API v3 (default saat menambah layanan Drive)
    const f = Drive.Files.create({ name: 'OCR_TMP_' + stamp_(), mimeType: MimeType.GOOGLE_DOCS }, blob,
      { ocrLanguage: CONFIG.OCR_LANGUAGE, fields: 'id' });
    docId = f.id;
  } catch (e3) {
    // Fallback Drive API v2
    const f = Drive.Files.insert({ title: 'OCR_TMP_' + stamp_(), mimeType: blob.getContentType() }, blob,
      { ocr: true, ocrLanguage: CONFIG.OCR_LANGUAGE });
    docId = f.id;
  }
  try {
    return DocumentApp.openById(docId).getBody().getText();
  } finally {
    try { DriveApp.getFileById(docId).setTrashed(true); } catch (e) { /* abaikan */ }
  }
}

const PLATE_PREFIX = ['A', 'B', 'D', 'E', 'F', 'G', 'H', 'K', 'L', 'M', 'N', 'P', 'R', 'S', 'T', 'W', 'Z',
  'AA', 'AB', 'AD', 'AE', 'AG', 'BA', 'BB', 'BD', 'BE', 'BG', 'BH', 'BK', 'BL', 'BM', 'BN', 'BP',
  'DA', 'DB', 'DC', 'DD', 'DE', 'DG', 'DH', 'DK', 'DL', 'DM', 'DN', 'DP', 'DR', 'DS', 'DT', 'DW',
  'EA', 'EB', 'ED', 'KB', 'KH', 'KT', 'KU', 'PA', 'PB', 'RI', 'CD', 'CC'];

function parseSpk_(text) {
  const out = { spk: '', nopol: '', type: '' };
  if (!text) return out;
  const lines = String(text).replace(/\r/g, '').split('\n').map(s => s.trim()).filter(Boolean);
  const upLines = lines.map(s => s.toUpperCase());
  const full = upLines.join('\n');

  // ---------- No. SPK ----------
  if (CONFIG.SPK_CUSTOM_REGEX) {
    const m = full.match(new RegExp(CONFIG.SPK_CUSTOM_REGEX, 'i'));
    if (m) out.spk = (m[1] || m[0]).trim();
  }
  if (!out.spk) {
    out.spk = findLabeled_(upLines,
      /(?:NO\.?|NOMOR|NOMER|NUMBER)?\s*(?:SPK|S\.P\.K\.?|W\/O|WO|WORK\s*ORDER|PKB|R\/O|RO|REPAIR\s*ORDER)\s*(?:NO\.?|NOMOR|#)?\s*[:.\-]?/,
      /([A-Z0-9][A-Z0-9\/\-\.]{3,24})/, v => /\d/.test(v));
  }

  // ---------- No. Polisi ----------
  const plateRe = /\b([A-Z]{1,2})\s*-?\s*(\d{1,4})\s*-?\s*([A-Z]{1,3})\b/;
  const labeledPlate = findLabeled_(upLines,
    /(?:NO\.?\s*POL(?:ISI)?|NOPOL|NO\.?\s*PLAT|PLAT\s*(?:NO|NOMOR)?|NO\.?\s*KENDARAAN|NO\.?\s*REG(?:ISTRASI)?|POLICE\s*NO|LICEN[CS]E\s*(?:PLATE|NO))\s*[:.\-]?/,
    plateRe, v => true, true);
  if (labeledPlate) out.nopol = labeledPlate;
  else {
    const re = new RegExp(plateRe.source, 'g');
    let m;
    while ((m = re.exec(full)) !== null) {
      if (PLATE_PREFIX.indexOf(m[1]) >= 0) { out.nopol = m[1] + ' ' + m[2] + ' ' + m[3]; break; }
    }
  }
  if (out.nopol) out.nopol = formatNopol_(out.nopol);

  // ---------- Type Mobil ----------
  const types = getTypes_();
  const labeledType = findLabeled_(upLines,
    /(?:TYPE|TIPE|MODEL|MERK\s*\/\s*TYPE|JENIS)\s*(?:MOBIL|KENDARAAN|UNIT)?\s*[:.\-]/,
    /([A-Z0-9][A-Z0-9 .\-\/]{1,40})/, v => !/^\d+$/.test(v));
  if (labeledType) {
    // Nilai berlabel (mis. "Type : Toyota Avanza 1.3 G") dipakai apa adanya
    out.type = titleCase_(labeledType.replace(/\s{2,}/g, ' ').trim());
  } else {
    // Tanpa label: cocokkan dengan daftar di sheet "Master Type Mobil"
    const compact = s => String(s).toUpperCase().replace(/[\s\-]/g, '');
    const hay = compact(full);
    const found = types.slice().sort((a, b) => b.length - a.length).find(t => hay.indexOf(compact(t)) >= 0);
    if (found) out.type = found;
  }
  return out;
}

/** Cari nilai setelah label (di baris yang sama atau baris berikutnya) */
function findLabeled_(upLines, labelRe, valueRe, validFn, isPlate) {
  for (let i = 0; i < upLines.length; i++) {
    const m = upLines[i].match(labelRe);
    if (!m) continue;
    const rest = upLines[i].slice(m.index + m[0].length).replace(/^[\s:.\-]+/, '');
    const candidates = [rest, upLines[i + 1] || '', upLines[i + 2] || ''];
    for (let c of candidates) {
      if (!c) continue;
      const v = c.match(valueRe);
      if (v) {
        const val = isPlate ? (v[1] + ' ' + v[2] + ' ' + v[3]) : (v[1] || v[0]).trim();
        if (validFn(val)) return val.trim();
      }
    }
  }
  return '';
}

/** ====================== HELPERS ====================== */

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Spreadsheet belum di-setup. Jalankan fungsi setup() terlebih dahulu.');
  return ss;
}

function col_(headers, name) {
  const i = headers.indexOf(name);
  if (i < 0) throw new Error('Kolom tidak ditemukan: ' + name);
  return i + 1;
}

function readTable_(sh) {
  if (!sh) return { headers: [], rows: [] };
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 1 || lc < 1) return { headers: [], rows: [] };
  const values = sh.getRange(1, 1, lr, lc).getValues();
  const headers = values[0].map(String);
  const rows = values.slice(1).map((r, i) => {
    const o = { _row: i + 2 };
    headers.forEach((h, j) => {
      const v = r[j];
      o[h] = v instanceof Date ? Utilities.formatDate(v, CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm') : v;
    });
    return o;
  }).filter(o => headers.some(h => o[h] !== '' && o[h] !== null));
  return { headers: headers, rows: rows };
}

function summary_(r) {
  return {
    id: r['ID Laporan'], tanggal: r['Tanggal Lapor'], teknisi: r['Nama Teknisi'], pic: r['PIC Tujuan'],
    spk: r['No. SPK'], nopol: r['No. Polisi'], type: r['Type Mobil'],
    jumlah: r['Jumlah Part'], berfoto: r['Part Berfoto'], diterima: r['Part Diterima PIC'],
    indFoto: r['Indikator Foto'], status: r['Status Serah Terima'], indikator: r['Indikator'],
    diterimaOleh: r['Diterima Oleh'], tglDiterima: r['Tanggal Diterima'],
    statusAsuransi: statusAsuransi_(r), asuransi: r['Nama Asuransi'] || '', noKlaim: r['No. Klaim'] || '',
    tglPlan: dateOnly_(r['Tanggal Planning Diantar ke Asuransi']),
    isoPlan: toIso_(r['Tanggal Planning Diantar ke Asuransi']),
    sisaHari: sisaHari_(r),
    tglAktual: dateOnly_(r['Tanggal Aktual Diantar ke Asuransi']),
    selisihHari: r['Selisih Hari (Aktual - Planning)'] === '' ? null : r['Selisih Hari (Aktual - Planning)'],
    ketepatan: r['Ketepatan Waktu'] || '',
    diantarOleh: r['Diantar Oleh (PIC)'] || '', penerimaAsuransi: r['Penerima di Asuransi'] || ''
  };
}

function statusAsuransi_(r) {
  const st = String(r['Status Asuransi'] || '');
  if (st === 'Belum Diantar') return ASURANSI_STATUS.BELUM; // data versi lama
  if (!st && (r['Status Serah Terima'] === STATUS.DITERIMA || r['Status Serah Terima'] === STATUS.SEBAGIAN)) return ASURANSI_STATUS.BELUM;
  return st;
}

function toIso_(s) {
  const d = parseSheetDate_(s);
  return d ? Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd') : '';
}

/** Sisa hari menuju planning (negatif = sudah lewat jadwal). Hanya untuk status Terjadwal. */
function sisaHari_(r) {
  if (statusAsuransi_(r) !== ASURANSI_STATUS.TERJADWAL) return null;
  const plan = parseSheetDate_(r['Tanggal Planning Diantar ke Asuransi']);
  return plan ? dayDiff_(todayDate_(), plan) : null;
}

function indikator_(jml, berfoto, adaSpk, status, diterima) {
  const fotoOk = adaSpk && jml > 0 && berfoto === jml;
  const foto = fotoOk ? '✅ Lengkap (' + berfoto + '/' + jml + ')' : '❌ Belum Lengkap (' + berfoto + '/' + jml + ')';
  let overall;
  if (!fotoOk) overall = '🔴 Foto Belum Lengkap';
  else if (status === STATUS.MENUNGGU) overall = '🟡 Menunggu Serah Terima PIC';
  else if (status === STATUS.DITERIMA && diterima === jml) overall = '🟢 Selesai - Semua Part Sesuai & Diterima';
  else if (status === STATUS.SEBAGIAN) overall = '🟠 Diterima Sebagian (' + diterima + '/' + jml + ')';
  else if (status === STATUS.DITOLAK) overall = '🔴 Ditolak PIC';
  else overall = '🟡 ' + status;
  return { foto: foto, overall: overall };
}

function verifyPic_(pic, pin) {
  const rows = readTable_(ss_().getSheetByName(SHEETS.PIC)).rows;
  const r = rows.find(x => String(x['Nama PIC']).trim() === pic && String(x['Aktif (Y/N)']).toUpperCase() !== 'N');
  if (!r) throw new Error('PIC tidak ditemukan.');
  if (String(r['PIN']).trim() !== String(pin || '').trim()) throw new Error('PIN salah.');
  return r;
}

function getTypes_() {
  return readTable_(ss_().getSheetByName(SHEETS.TYPE)).rows
    .map(r => String(r['Type Mobil'] || '').trim()).filter(Boolean);
}

function findDuplicateSpk_(spk) {
  const key = String(spk || '').toUpperCase().replace(/\s+/g, '');
  if (!key) return null;
  const r = readTable_(ss_().getSheetByName(SHEETS.LAPORAN)).rows
    .find(x => String(x['No. SPK']).toUpperCase().replace(/\s+/g, '') === key);
  return r ? { id: r['ID Laporan'], tanggal: r['Tanggal Lapor'], status: r['Status Serah Terima'] } : null;
}

function nextId_(sh) {
  const prefix = 'LP-' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyMMdd') + '-';
  let n = 0;
  const lr = sh.getLastRow();
  if (lr > 1) {
    sh.getRange(2, 1, lr - 1, 1).getValues().forEach(r => {
      const s = String(r[0]);
      if (s.indexOf(prefix) === 0) n = Math.max(n, parseInt(s.slice(prefix.length), 10) || 0);
    });
  }
  return prefix + ('00' + (n + 1)).slice(-3);
}

function formatNopol_(v) {
  const s = String(v || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  const m = s.match(/^([A-Z]{1,2})(\d{1,4})([A-Z]{0,3})$/);
  return m ? [m[1], m[2], m[3]].filter(Boolean).join(' ') : String(v || '').toUpperCase().trim();
}

function titleCase_(s) {
  return s.toLowerCase().replace(/\b([a-z])/g, c => c.toUpperCase());
}

function dataUrlToBlob_(dataUrl, name) {
  const m = String(dataUrl || '').match(/^data:([\w\/+.\-]+);base64,(.+)$/);
  if (!m) throw new Error('Format gambar tidak valid.');
  if (!/^image\//.test(m[1])) throw new Error('File harus berupa gambar.');
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > CONFIG.MAX_IMAGE_MB * 1024 * 1024) throw new Error('Ukuran gambar terlalu besar.');
  return Utilities.newBlob(bytes, m[1], name);
}

function getRootFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('ROOT_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* buat ulang */ } }
  const it = DriveApp.getFoldersByName(CONFIG.ROOT_FOLDER_NAME);
  const f = it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.ROOT_FOLDER_NAME);
  props.setProperty('ROOT_FOLDER_ID', f.getId());
  return f;
}

function getTempFolder_() { return getOrCreate_(getRootFolder_(), '_upload_sementara'); }

function getOrCreate_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function moveFile_(fileId, folder, newName) {
  const f = DriveApp.getFileById(fileId);
  if (!isInRoot_(f)) throw new Error('File tidak valid.');
  f.moveTo(folder);
  if (newName) f.setName(String(newName).replace(/[^\w\-]+/g, '_') + '.jpg');
  return f;
}

function isInRoot_(file) {
  const rootId = getRootFolder_().getId();
  let parents = file.getParents();
  for (let depth = 0; depth < 5 && parents.hasNext(); depth++) {
    const p = parents.next();
    if (p.getId() === rootId) return true;
    parents = p.getParents();
  }
  return false;
}

function fileUrl_(f) { return 'https://drive.google.com/file/d/' + f.getId() + '/view'; }

function idFromUrl_(url) {
  const m = String(url || '').match(/[-\w]{25,}/);
  return m ? m[0] : '';
}

function stamp_() { return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd_HHmmss_SSS'); }

function notifyPic_(pic, r) {
  if (!CONFIG.NOTIFY_PIC_EMAIL) return;
  try {
    const row = readTable_(ss_().getSheetByName(SHEETS.PIC)).rows.find(x => String(x['Nama PIC']).trim() === pic);
    const email = row && String(row['Email'] || '').trim();
    if (!email) return;
    const url = ScriptApp.getService().getUrl();
    MailApp.sendEmail({
      to: email,
      subject: '[Part Bekas] Serah terima baru ' + r['ID Laporan'] + ' - ' + r['No. Polisi'],
      htmlBody: 'Halo ' + pic + ',<br><br>Ada laporan part bekas yang menunggu serah terima:<br>' +
        '<b>ID:</b> ' + r['ID Laporan'] + '<br><b>No. SPK:</b> ' + r['No. SPK'] + '<br><b>No. Polisi:</b> ' + r['No. Polisi'] +
        '<br><b>Type:</b> ' + r['Type Mobil'] + '<br><b>Teknisi:</b> ' + r['Nama Teknisi'] + '<br><b>Jumlah Part:</b> ' + r['Jumlah Part'] +
        '<br><br><a href="' + url + '">Buka aplikasi serah terima</a>'
    });
  } catch (e) { Logger.log('Email gagal: ' + e); }
}

/** ====================== UJI OCR (opsional) ====================== */
function testParse() {
  const sample = 'BENGKEL MAJU JAYA\nSURAT PERINTAH KERJA\nNo. SPK : SPK/2026/11/00123\nNo. Polisi : B 1234 XYZ\nType : Toyota Avanza 1.3 G\nKeluhan: ganti kampas rem';
  Logger.log(JSON.stringify(parseSpk_(sample)));
}
