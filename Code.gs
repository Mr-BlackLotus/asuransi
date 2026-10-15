/**
 * Report Asuransi - backend Google Apps Script
 * Kolom: Foto (opsional), Nominal (wajib), No Pesanan MP (wajib), No Invoice EK (opsional).
 * Data disimpan ke Google Sheet, foto disimpan ke Google Drive.
 */

const SHEET_ID = '14FIH5zkBPqYbOO_vi2Hgx-R1HBgmnmQPeAjnm6S8Mtw';
const SHEET_NAME = 'Sheet1';
const FOLDER_NAME = 'Foto Report Asuransi';
const HEADERS = ['Tanggal Input', 'Foto', 'Nominal', 'No Pesanan MP', 'No Invoice EK', 'Link Foto'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Report Asuransi')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setBackground('#e8f0fe');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 180);
    sheet.setColumnWidth(5, 180);
    sheet.setColumnWidth(6, 220);
  }
  return sheet;
}

function getFolder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function norm_(s) {
  return String(s == null ? '' : s).trim().toLowerCase();
}

/** Cek apakah No MP / No EK sudah pernah diinput (EK kosong tidak dicek). */
function checkDuplicate(noMP, noEK) {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return { mp: false, ek: false };
  const vals = sheet.getRange(2, 4, last - 1, 2).getValues();
  const mp = norm_(noMP);
  const ek = norm_(noEK);
  let dupMP = false, dupEK = false;
  vals.forEach(function (r) {
    if (mp && norm_(r[0]) === mp) dupMP = true;
    if (ek && norm_(r[1]) === ek) dupEK = true;
  });
  return { mp: dupMP, ek: dupEK };
}

/** Simpan satu data baru ke baris berikutnya. Foto dan EK boleh kosong. */
function saveEntry(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const nominal = Number(data.nominal);
    if (!nominal || nominal <= 0) throw new Error('Nominal tidak valid');
    const noMP = String(data.noMP || '').trim();
    if (!noMP) throw new Error('No Pesanan MP wajib diisi');
    const noEK = String(data.noEK || '').trim();

    const sheet = getSheet_();

    // 1) Foto (opsional) -> simpan ke Drive
    let viewUrl = '', thumbUrl = '';
    if (data.photoBase64) {
      const bytes = Utilities.base64Decode(data.photoBase64);
      const stamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss');
      const safeMP = noMP.replace(/[^\w\-]/g, '_');
      const blob = Utilities.newBlob(bytes, data.mime || 'image/jpeg', stamp + '_' + safeMP + '.jpg');
      const file = getFolder_().createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const id = file.getId();
      viewUrl = 'https://drive.google.com/file/d/' + id + '/view';
      thumbUrl = 'https://drive.google.com/thumbnail?id=' + id + '&sz=w400';
    }

    // 2) Tambah baris ke Sheet
    const row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1).setValue(new Date()).setNumberFormat('dd/MM/yyyy HH:mm');
    if (thumbUrl) sheet.getRange(row, 2).setFormula('=IMAGE("' + thumbUrl + '")');
    sheet.getRange(row, 3).setValue(nominal).setNumberFormat('"Rp"#,##0');
    sheet.getRange(row, 4).setNumberFormat('@').setValue(noMP);
    sheet.getRange(row, 5).setNumberFormat('@').setValue(noEK);
    if (viewUrl) sheet.getRange(row, 6).setValue(viewUrl);
    if (thumbUrl) sheet.setRowHeight(row, 110);
    sheet.getRange(row, 1, 1, HEADERS.length).setVerticalAlignment('middle');

    return { ok: true, row: row };
  } finally {
    lock.releaseLock();
  }
}

/** Ubah isi sel tanggal menjadi 'yyyy-MM-dd' (kosong jika tidak terbaca). */
function isoDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Jakarta', 'yyyy-MM-dd');
  const s = String(v || '');
  let m = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); // dd/MM/yyyy
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return '';
}

/** Ambil semua data (terbaru di atas) untuk ditampilkan di aplikasi. */
function getAll() {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return { items: [] };
  const vals = sheet.getRange(2, 1, last - 1, 6).getValues();
  const items = [];
  vals.forEach(function (r, i) {
    if (!r[2] && !r[3] && !r[4]) return; // lewati baris kosong
    const link = String(r[5] || '');
    const m = link.match(/\/d\/([^\/?]+)/);
    items.push({
      row: i + 2,
      tgl: isoDate_(r[0]),
      tanggal: r[0] instanceof Date
        ? Utilities.formatDate(r[0], 'Asia/Jakarta', 'dd/MM/yyyy HH:mm') : String(r[0] || ''),
      nominal: Number(r[2]) || 0,
      noMP: String(r[3] || ''),
      noEK: String(r[4] || ''),
      link: link,
      fileId: m ? m[1] : ''
    });
  });
  items.reverse();
  return { items: items };
}

/**
 * Ubah No Invoice EK pada satu baris.
 * noMP dipakai sebagai pengaman: kalau baris di Sheet sudah bergeser/berubah, perubahan ditolak.
 * Jika EK baru sudah dipakai baris lain dan force != true, kembalikan {ok:false, dup:true}.
 */
function updateEK(row, noMP, newEK, force) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getSheet_();
    row = Number(row);
    if (!row || row < 2 || row > sheet.getLastRow()) {
      throw new Error('Baris tidak ditemukan, muat ulang halaman.');
    }
    const currentMP = String(sheet.getRange(row, 4).getValue()).trim();
    if (currentMP !== String(noMP).trim()) {
      throw new Error('Data di Sheet sudah berubah, muat ulang halaman.');
    }
    const ek = String(newEK == null ? '' : newEK).trim();

    if (ek && !force) {
      const last = sheet.getLastRow();
      const eks = sheet.getRange(2, 5, last - 1, 1).getValues();
      for (let i = 0; i < eks.length; i++) {
        if (i + 2 !== row && norm_(eks[i][0]) === norm_(ek)) {
          return { ok: false, dup: true };
        }
      }
    }

    sheet.getRange(row, 5).setNumberFormat('@').setValue(ek);
    return { ok: true, noEK: ek };
  } finally {
    lock.releaseLock();
  }
}
