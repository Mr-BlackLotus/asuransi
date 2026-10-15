/**
 * Report Asuransi - backend Google Apps Script
 * Menyimpan: Foto, Nominal, No Pesanan MP, No Invoice EK ke Google Sheet + Google Drive.
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

/** Cek apakah No MP / No EK sudah pernah diinput. */
function checkDuplicate(noMP, noEK) {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return { mp: false, ek: false };
  const vals = sheet.getRange(2, 4, last - 1, 2).getValues();
  const mp = String(noMP).trim().toLowerCase();
  const ek = String(noEK).trim().toLowerCase();
  let dupMP = false, dupEK = false;
  vals.forEach(function (r) {
    if (mp && String(r[0]).trim().toLowerCase() === mp) dupMP = true;
    if (ek && String(r[1]).trim().toLowerCase() === ek) dupEK = true;
  });
  return { mp: dupMP, ek: dupEK };
}

/** Simpan satu data baru ke baris berikutnya. */
function saveEntry(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const nominal = Number(data.nominal);
    if (!nominal || nominal <= 0) throw new Error('Nominal tidak valid');
    if (!data.noMP || !data.noEK) throw new Error('No Pesanan MP dan No Invoice EK wajib diisi');
    if (!data.photoBase64) throw new Error('Foto wajib diisi');

    const sheet = getSheet_();

    // 1) Simpan foto ke Drive
    const bytes = Utilities.base64Decode(data.photoBase64);
    const stamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss');
    const safeMP = String(data.noMP).replace(/[^\w\-]/g, '_');
    const blob = Utilities.newBlob(bytes, data.mime || 'image/jpeg', stamp + '_' + safeMP + '.jpg');
    const file = getFolder_().createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const id = file.getId();
    const viewUrl = 'https://drive.google.com/file/d/' + id + '/view';
    const thumbUrl = 'https://drive.google.com/thumbnail?id=' + id + '&sz=w400';

    // 2) Tambah baris ke Sheet
    const row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1).setValue(new Date()).setNumberFormat('dd/MM/yyyy HH:mm');
    sheet.getRange(row, 2).setFormula('=IMAGE("' + thumbUrl + '")');
    sheet.getRange(row, 3).setValue(nominal).setNumberFormat('"Rp"#,##0');
    sheet.getRange(row, 4).setNumberFormat('@').setValue(String(data.noMP).trim());
    sheet.getRange(row, 5).setNumberFormat('@').setValue(String(data.noEK).trim());
    sheet.getRange(row, 6).setValue(viewUrl);
    sheet.setRowHeight(row, 110);
    sheet.getRange(row, 1, 1, HEADERS.length).setVerticalAlignment('middle');

    return { ok: true, row: row, url: viewUrl };
  } finally {
    lock.releaseLock();
  }
}

/** Ambil data terakhir untuk ditampilkan di aplikasi. */
function getRecent(limit) {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return { items: [], total: 0 };
  const n = Math.min(limit || 10, last - 1);
  const start = last - n + 1;
  const vals = sheet.getRange(start, 1, n, 6).getValues();
  const items = vals.map(function (r) {
    return {
      tanggal: r[0] instanceof Date
        ? Utilities.formatDate(r[0], 'Asia/Jakarta', 'dd/MM/yyyy HH:mm') : String(r[0]),
      nominal: Number(r[2]) || 0,
      noMP: String(r[3]),
      noEK: String(r[4]),
      link: String(r[5])
    };
  }).reverse();
  return { items: items, total: last - 1 };
}
