/**
 * Sistem Harian Kepala Sekolah — SDN Labuhan Bajo
 * Backend Google Apps Script. Data dicadangkan di file JSON privat pada Google Drive.
 * Lampiran disimpan sebagai file privat di folder Drive aplikasi.
 */
const APP_FOLDER_NAME = 'Sistem Harian Kepala Sekolah - SDN Labuhan Bajo';
const APP_STATE_FILENAME = 'data-sistem-harian-kepala-sekolah.json';
const ATTACHMENT_FOLDER_NAME = 'Lampiran';
const MAX_ATTACHMENT_BYTES = 15 * 1024 * 1024;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Sistem Harian Kepala Sekolah')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getAppFolder_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('APP_FOLDER_ID');
  if (savedId) {
    try { return DriveApp.getFolderById(savedId); } catch (e) {}
  }
  const folder = DriveApp.createFolder(APP_FOLDER_NAME);
  props.setProperty('APP_FOLDER_ID', folder.getId());
  return folder;
}

function getAttachmentFolder_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('APP_ATTACHMENT_FOLDER_ID');
  if (savedId) {
    try { return DriveApp.getFolderById(savedId); } catch (e) {}
  }
  const parent = getAppFolder_();
  const found = parent.getFoldersByName(ATTACHMENT_FOLDER_NAME);
  const folder = found.hasNext() ? found.next() : parent.createFolder(ATTACHMENT_FOLDER_NAME);
  props.setProperty('APP_ATTACHMENT_FOLDER_ID', folder.getId());
  return folder;
}

function loadAppState() {
  const props = PropertiesService.getScriptProperties();
  const fileId = props.getProperty('APP_STATE_FILE_ID');
  if (!fileId) return null;
  try {
    const text = DriveApp.getFileById(fileId).getBlob().getDataAsString('UTF-8');
    return text ? JSON.parse(text) : null;
  } catch (err) {
    throw new Error('Data aplikasi di Google Drive tidak dapat dibaca: ' + err.message);
  }
}

function saveAppState(jsonText) {
  if (typeof jsonText !== 'string' || !jsonText) throw new Error('Data kosong atau tidak valid.');
  if (jsonText.length > 45000000) throw new Error('Data aplikasi terlalu besar untuk disimpan sekaligus. Buat cadangan dan kurangi lampiran/logo yang terlalu besar.');
  let parsed;
  try { parsed = JSON.parse(jsonText); } catch (err) { throw new Error('Format data tidak valid.'); }
  if (!parsed || typeof parsed !== 'object' || !parsed.settings) throw new Error('Struktur data aplikasi tidak dikenali.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const props = PropertiesService.getScriptProperties();
    const fileId = props.getProperty('APP_STATE_FILE_ID');
    if (fileId) {
      try {
        DriveApp.getFileById(fileId).setContent(jsonText);
        return { ok: true, updated: true, savedAt: new Date().toISOString() };
      } catch (err) {
        props.deleteProperty('APP_STATE_FILE_ID');
      }
    }
    const file = getAppFolder_().createFile(APP_STATE_FILENAME, jsonText, MimeType.PLAIN_TEXT);
    props.setProperty('APP_STATE_FILE_ID', file.getId());
    return { ok: true, created: true, savedAt: new Date().toISOString() };
  } finally {
    lock.releaseLock();
  }
}

function putAttachment(payload) {
  if (!payload || !payload.id || !payload.name || !payload.dataUrl) throw new Error('Data lampiran tidak lengkap.');
  const dataUrl = String(payload.dataUrl);
  const match = dataUrl.match(/^data:([^;,]+)?;base64,([A-Za-z0-9+/=\r\n]+)$/);
  if (!match) throw new Error('Format file lampiran tidak dikenali.');
  const bytes = Utilities.base64Decode(match[2]);
  if (bytes.length > MAX_ATTACHMENT_BYTES) throw new Error('Ukuran lampiran melebihi 15 MB.');
  const mime = String(payload.type || match[1] || 'application/octet-stream');
  if (!(mime === 'application/pdf' || mime.indexOf('image/') === 0 || /\.pdf$/i.test(payload.name))) {
    throw new Error('Hanya dokumen PDF atau gambar yang diperbolehkan.');
  }
  const props = PropertiesService.getScriptProperties();
  const key = 'ATTACHMENT_FILE_' + String(payload.id).replace(/[^A-Za-z0-9_-]/g, '_');
  const priorId = props.getProperty(key);
  if (priorId) {
    try { DriveApp.getFileById(priorId).setTrashed(true); } catch (e) {}
  }
  const safeName = String(payload.name).replace(/[\\/:*?"<>|#%{}~&]/g, '_').slice(0, 180);
  const blob = Utilities.newBlob(bytes, mime, safeName);
  const file = getAttachmentFolder_().createFile(blob);
  file.setName(String(payload.id) + '__' + safeName);
  props.setProperty(key, file.getId());
  return { id: String(payload.id), name: safeName, type: mime, size: bytes.length, createdAt: payload.createdAt || new Date().toISOString() };
}

function getAttachmentData(id) {
  if (!id) return null;
  const key = 'ATTACHMENT_FILE_' + String(id).replace(/[^A-Za-z0-9_-]/g, '_');
  const fileId = PropertiesService.getScriptProperties().getProperty(key);
  if (!fileId) return null;
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const bytes = blob.getBytes();
    if (bytes.length > MAX_ATTACHMENT_BYTES) throw new Error('Lampiran melebihi ukuran maksimum yang didukung.');
    const mime = blob.getContentType() || 'application/octet-stream';
    const dataUrl = 'data:' + mime + ';base64,' + Utilities.base64Encode(bytes);
    return { id: String(id), name: file.getName().replace(String(id) + '__', ''), type: mime, size: bytes.length, createdAt: file.getDateCreated().toISOString(), dataUrl: dataUrl };
  } catch (err) {
    throw new Error('Lampiran tidak dapat dibaca dari Google Drive: ' + err.message);
  }
}

function deleteAttachment(id) {
  if (!id) return { ok: true };
  const key = 'ATTACHMENT_FILE_' + String(id).replace(/[^A-Za-z0-9_-]/g, '_');
  const props = PropertiesService.getScriptProperties();
  const fileId = props.getProperty(key);
  if (fileId) {
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
    props.deleteProperty(key);
  }
  return { ok: true };
}

/** Meminta otorisasi Drive lebih awal saat pemilik menjalankan fungsi ini dari editor. */
function authorizeApp() {
  getAppFolder_();
  getAttachmentFolder_();
  return 'Akses Google Drive sudah diotorisasi. Sekarang deploy sebagai Web App.';
}
