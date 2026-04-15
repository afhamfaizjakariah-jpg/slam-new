/******************************
 * 09_PdfExtraction.gs
 ******************************/

/** ====== DRIVE TEMP EXTRACTOR ====== **/
function _getOrCreateTempFolder_() {
  const name = String(CONFIG.TEMP_PDF_FOLDER_NAME || "SLAM_TEMP_PDF").trim();
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

function _extractPdfTextViaDrive_(bytes, filename) {
  const notes = [];
  let pdfFileId = null;
  let docFileId = null;

  try {
    const folder = _getOrCreateTempFolder_();
    const blob = Utilities.newBlob(bytes, "application/pdf", filename || "aduan.pdf");
    const pdfFile = folder.createFile(blob);
    pdfFileId = pdfFile.getId();
    notes.push("Temp PDF created: " + pdfFileId);

    const baseName = (filename || "aduan").replace(/\.pdf$/i, "");

    docFileId = _driveConvertPdfToDoc_(blob, folder.getId(), baseName + "_DOC", false);
    notes.push("Converted to Google Doc (no OCR): " + docFileId);

    let text = _readDocTextWithRetry_(docFileId);
    if (_looksLikeSispaa_(text)) return { text: (text || "").trim(), notes };

    if (CONFIG.DRIVE_OCR_ENABLED) {
      try { DriveApp.getFileById(docFileId).setTrashed(true); } catch (e) {}
      docFileId = _driveConvertPdfToDoc_(blob, folder.getId(), baseName + "_OCR", true);
      notes.push("Converted to Google Doc (OCR): " + docFileId);
      text = _readDocTextWithRetry_(docFileId);
    } else {
      notes.push("OCR disabled.");
    }

    return { text: (text || "").trim(), notes };
  } catch (e) {
    notes.push("Drive conversion failed: " + (e && e.message ? e.message : String(e)));
    return { text: "", notes };
  } finally {
    try { if (docFileId) DriveApp.getFileById(docFileId).setTrashed(true); } catch (e) {}
    try { if (pdfFileId) DriveApp.getFileById(pdfFileId).setTrashed(true); } catch (e) {}
  }
}

function _driveConvertPdfToDoc_(blob, folderId, name, useOcr) {
  if (typeof Drive !== "object" || !Drive.Files) {
    throw new Error("Advanced Drive Service tidak tersedia. Enable di Apps Script: Services > Drive API.");
  }

  if (typeof Drive.Files.insert === "function") {
    const resourceV2 = { title: name, mimeType: blob.getContentType(), parents: [{ id: folderId }] };
    const optsV2 = { convert: true };
    if (useOcr) {
      optsV2.ocr = true;
      optsV2.ocrLanguage = CONFIG.DRIVE_OCR_LANG || "ms";
    }
    const file = Drive.Files.insert(resourceV2, blob, optsV2);
    return file && file.id;
  }

  if (typeof Drive.Files.create === "function") {
    const resourceV3 = { name: name, mimeType: "application/vnd.google-apps.document", parents: [folderId] };
    const optsV3 = {};
    if (useOcr) {
      optsV3.ocr = true;
      optsV3.ocrLanguage = CONFIG.DRIVE_OCR_LANG || "ms";
    }
    const file = Drive.Files.create(resourceV3, blob, optsV3);
    return file && file.id;
  }

  throw new Error("Drive.Files.insert/create tiada. Semak versi Drive API service.");
}

function _readDocTextWithRetry_(docId) {
  let lastErr = null;
  for (let i = 0; i < 6; i++) {
    try { return DocumentApp.openById(docId).getBody().getText() || ""; }
    catch (e) { lastErr = e; Utilities.sleep(1500); }
  }
  throw new Error("Gagal baca teks dokumen selepas convert: " + String(lastErr && lastErr.message || lastErr));
}

function _looksLikeSispaa_(text) {
  const t = String(text || "");
  if (/ID Maklum Balas\s*:/i.test(t)) return true;
  if (/[A-Z]{2,5}\.\d{3,}/.test(t)) return true;
  return false;
}

function _bestEffortExtractPdfText_(bytes) {
  const notes = [];
  let str = "";
  try { str = Utilities.newBlob(bytes).getDataAsString("ISO-8859-1"); }
  catch (e) {
    notes.push("Gagal baca bytes sebagai string.");
    return { text: "", notes };
  }

  const out = [];
  try {
    const re = /\(([^()]|\\\(|\\\)|\\n|\\r|\\t|\\\d{1,3}){2,}\)\s*(Tj|TJ)/g;
    let m;
    while ((m = re.exec(str)) !== null) {
      const raw = m[0];
      const inner = raw.replace(/\)\s*(Tj|TJ)\s*$/i, "").slice(1);
      out.push(_pdfUnescapeLiteral_(inner));
      if (out.length > 4000) break;
    }
  } catch (e) {
    notes.push("Regex literal string gagal.");
  }

  const text = out.join("\n").replace(/\s+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
  if (!text) notes.push("Teks kosong (mungkin PDF scan / stream compressed).");
  return { text, notes };
}

function _pdfUnescapeLiteral_(s) {
  let out = s;
  out = out.replace(/\\n/g, "\n").replace(/\\r/g, "\r").replace(/\\t/g, "\t");
  out = out.replace(/\\\(/g, "(").replace(/\\\)/g, ")").replace(/\\\\/g, "\\");
  out = out.replace(/\\(\d{1,3})/g, function(_, d) {
    const code = parseInt(d, 8);
    if (isNaN(code)) return _;
    return String.fromCharCode(code);
  });
  return out;
}

function _getOrCreateSourcePdfFolder_() {
  const name = "SLAM_SOURCE_PDF";
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

function _storeSourcePdf_(bytes, filename) {
  const folder = _getOrCreateSourcePdfFolder_();
  const safeName = String(filename || "aduan.pdf").trim() || "aduan.pdf";
  const blob = Utilities.newBlob(bytes, "application/pdf", safeName);
  const file = folder.createFile(blob);

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}

  return {
    file_id: file.getId(),
    url: "https://drive.google.com/file/d/" + file.getId() + "/view"
  };
}