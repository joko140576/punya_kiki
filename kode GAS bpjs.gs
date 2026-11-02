/**
 * Google Apps Script — penerima data Form BPJS
 *  - Data teks   -> Google Spreadsheet (sheet "Respon")
 *  - File upload -> Google Drive (langsung di dalam 1 folder, TANPA subfolder per pengisi),
 *                   link tiap file ditulis di kolom pertanyaannya masing-masing.
 *
 * Cara pakai:
 * 1. Buka Spreadsheet > Extensions > Apps Script > tempel kode ini > simpan.
 * 2. Pilih fungsi "authorize" di toolbar > Jalankan > izinkan akses (Spreadsheet + Drive).
 * 3. Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * 4. Salin URL "/exec" ke GAS_URL di index.html.
 * Setelah mengubah kode: Deploy > Manage deployments > pensil > New version > Deploy.
 */

const SHEET_NAME       = 'Respon';              // nama sheet tujuan (dibuat otomatis)
const SPREADSHEET_ID   = '';                    // kosongkan bila script dibuat dari dalam Spreadsheet
const DRIVE_FOLDER_ID  = '';                    // ID folder tujuan upload (opsional)
const ROOT_FOLDER_NAME = 'Upload Form BPJS';    // dipakai bila DRIVE_FOLDER_ID kosong; kosongkan "" = simpan di My Drive
const TZ               = 'Asia/Jakarta';

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const body = parseBody_(e);
    if (body._hp) return out_({ ok: true });                 // honeypot anti-bot

    const data = Object.assign({}, body.fields || {});

    // ---- simpan file ke Drive: langsung di folder induk, hanya link file di kolom pertanyaannya ----
    // (bila gagal, data teks tetap masuk Spreadsheet dan sebabnya ditulis di kolom "Catatan Error")
    const files = body.files || [];
    if (files.length) {
      try {
        const folder = targetFolder_();
        const stamp  = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss');
        files.forEach(f => {
          const ext  = (String(f.filename).match(/\.[A-Za-z0-9]+$/) || [''])[0];
          const name = safe_((data['NPP pegawai'] || 'tanpa-npp') + ' - ' + (data['Nama pegawai'] || 'tanpa-nama') + ' - ' +
                             String(f.field).replace(/^Upload\s+/i, '') + ' - ' + stamp) + ext;
          const blob = Utilities.newBlob(Utilities.base64Decode(f.data), f.mime || 'application/octet-stream', name);
          data[f.field] = folder.createFile(blob).getUrl();
        });
      } catch (fileErr) {
        data['Catatan Error'] = 'Upload file gagal: ' + String(fileErr);
      }
    }

    // ---- tulis ke Spreadsheet ----
    const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    const keys = Object.keys(data);

    let headers = [];
    if (sheet.getLastRow() === 0) {
      headers = ['Timestamp'].concat(keys);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e2f4ea');
      sheet.setFrozenRows(1);
    } else {
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      while (headers.length && headers[headers.length - 1] === '') headers.pop();   // buang kolom kosong di ujung
      keys.forEach(k => {
        if (headers.indexOf(k) === -1) {
          headers.push(k);
          sheet.getRange(1, headers.length).setValue(k).setFontWeight('bold').setBackground('#e2f4ea');
        }
      });
    }
    const row = headers.map(h => h === 'Timestamp'
      ? Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss')
      : (data[h] !== undefined ? data[h] : ''));

    // Tulis ke baris kosong pertama (tahan terhadap "Tabel" Google Sheets / baris kosong yang sudah disiapkan)
    sheet.getRange(nextRow_(sheet), 1, 1, headers.length).setValues([row]);
    SpreadsheetApp.flush();

    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

// Tes cepat: buka URL /exec di browser -> {"ok":true,...}
function doGet() {
  let info = '';
  try {
    const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    info = ss ? ('Terhubung ke Spreadsheet: ' + ss.getName()) : 'TIDAK terhubung ke Spreadsheet (isi SPREADSHEET_ID)';
  } catch (err) { info = 'Error Spreadsheet: ' + err; }
  return out_({ ok: true, message: 'Web app aktif', spreadsheet: info });
}

// Jalankan dari editor untuk tes tulis ke Spreadsheet (tanpa form)
function tesTulis() {
  doPost({ postData: { contents: JSON.stringify({ fields: { 'Nama pegawai': 'TES', 'NPP pegawai': '000' }, files: [] }) } });
  Logger.log('Cek sheet "' + SHEET_NAME + '" - harus ada baris TES.');
}

// Jalankan SEKALI dari editor untuk memberi izin Spreadsheet + Drive.
// Membuat lalu langsung membuang folder uji, supaya izin Drive penuh benar-benar diminta.
function authorize() {
  const f = DriveApp.createFolder('tes-izin-hapus-saja');
  f.setTrashed(true);
  if (SPREADSHEET_ID) SpreadsheetApp.openById(SPREADSHEET_ID); else SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Izin Drive + Spreadsheet sudah aktif.');
}

/* ---------- helper ---------- */

// Cari baris kosong pertama di kolom A (Timestamp), tanpa bergantung pada getLastRow()
function nextRow_(sheet) {
  const max = sheet.getMaxRows();
  if (max < 2) { sheet.insertRowsAfter(max, 1); return 2; }
  const colA = sheet.getRange(2, 1, max - 1, 1).getValues();
  for (let i = 0; i < colA.length; i++) if (colA[i][0] === '') return i + 2;
  sheet.insertRowsAfter(max, 1);
  return max + 1;
}

function parseBody_(e) {
  const txt = e && e.postData && e.postData.contents;
  if (txt) { try { return JSON.parse(txt); } catch (x) {} }
  const fields = {}; const p = (e && e.parameter) || {};
  Object.keys(p).forEach(k => { if (k !== '_hp') fields[k] = p[k]; });
  return { fields: fields, files: [], _hp: p._hp };
}

function targetFolder_() {
  if (DRIVE_FOLDER_ID) return DriveApp.getFolderById(DRIVE_FOLDER_ID);
  if (!ROOT_FOLDER_NAME) return DriveApp.getRootFolder();
  const it = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(ROOT_FOLDER_NAME);
}

function safe_(s) { return String(s).replace(/[\\\/:*?"<>|]+/g, '_').trim().slice(0, 100); }

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
