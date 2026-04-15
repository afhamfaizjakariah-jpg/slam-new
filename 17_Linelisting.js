/******************************
 * 17_Linelisting.gs
 ******************************/

const LINELISTING_HEADERS = [
  "BIL",
  "NO TIKET",
  "TARIKH TERIMA",
  "TARIKH AGIH",
  "TARIKH SIASATAN",
  "TARIKH SIAP",
  "BULAN",
  "SUMBER",
  "STATUS",
  "TAHAP KESUKARAN",
  "PENGADU",
  "PREMIS",
  "LOKASI",
  "PEGAWAI PENYIASAT",
  "PARLIMEN",
  "KATEGORI",
  "JENIS PREMIS",
  "RUMUSAN",
  "MARKAH PEMERIKSAAN",
  "TINDAKAN"
];

function _linelistingText_(v) {
  return String(v || "").trim();
}

function _linelistingMonthName_(val) {
  const d = _safeDate_(val) || _safeDate_(_normalizeDate_(val));
  if (!d) return "";
  const bulan = [
    "Januari", "Februari", "Mac", "April", "Mei", "Jun",
    "Julai", "Ogos", "September", "Oktober", "November", "Disember"
  ];
  return bulan[d.getMonth()];
}

function _linelistingSourceLabel_(rec) {
  const raw = _linelistingText_(rec.source || rec.report_jenis_aduan || rec.jenis_maklumbalas).toUpperCase();
  if (!raw) return "";
  if (raw === "MOH" || raw === "PCB" || raw === "JPA" || raw === "SISPAA") return "SiSPAA";
  if (raw === "MEDIA SOSIAL" || raw === "MS") return "Media Sosial";
  if (raw === "EMEL" || raw === "EMAIL") return "Emel";
  if (raw === "WHATSAPP ADUAN" || raw === "WHATSAPP" || raw === "WA") return "Whatsapp Aduan";
  if (raw === "HADIR SENDIRI" || raw === "HS") return "Hadir Sendiri";
  if (raw === "TELEFON" || raw === "TL") return "Telefon";
  if (raw === "SURAT" || raw === "SR") return "Surat";
  if (raw === "LAIN-LAIN" || raw === "LAIN LAIN" || raw === "LL") return "Lain-Lain";
  if (raw === "PLATFORM NEGERI" || raw === "PN") return "Platform Negeri";
  return _linelistingText_(rec.source || rec.report_jenis_aduan || rec.jenis_maklumbalas);
}

function _linelistingCompletionDate_(rec) {
  return (
    _linelistingText_(rec.report_generated_at) ||
    _linelistingText_(rec.report_submitted_at) ||
    _linelistingText_(rec.report_updated_at) ||
    _linelistingText_(rec.status_updated_at) ||
    ""
  );
}

function _linelistingDiffDays_(startVal, endVal) {
  const start = _safeDate_(startVal) || _safeDate_(_normalizeDate_(startVal));
  const end = _safeDate_(endVal) || _safeDate_(_normalizeDate_(endVal));
  if (!start || !end) return null;
  const ms = end.getTime() - start.getTime();
  if (ms < 0) return null;
  return Math.floor(ms / 86400000);
}

function _linelistingSlaLabel_(diffDays, status) {
  const st = _linelistingText_(status).toUpperCase();
  if (diffDays === null || diffDays === "") {
    return st === "SELESAI" ? "Tiada Data Tempoh" : "Belum Selesai";
  }
  return diffDays < 15 ? "Dalam Tempoh" : "Melebihi Tempoh";
}

function _linelistingBuildRow_(rec, bil) {
  const status = _linelistingText_(_computeEffectiveCardStatus_(rec) || rec.status || "");
  const siapAt = _linelistingCompletionDate_(rec);
  const diffDays = _linelistingDiffDays_(rec.tarikh_terima, siapAt);

  return {
    "BIL": bil,
    "NO TIKET": _linelistingText_(rec.id_maklumbalas || rec.complaint_id),
    "TARIKH TERIMA": _formatDateMalay_(rec.tarikh_terima || ""),
    "TARIKH AGIH": _formatDateMalay_(rec.assigned_at || ""),
    "TARIKH SIASATAN": _formatDateMalay_(rec.report_tarikh_siasatan || ""),
    "TARIKH SIAP": _formatDateMalay_(siapAt || ""),
    "BULAN": _linelistingMonthName_(rec.tarikh_terima || rec.created_at || ""),
    "SUMBER": _linelistingSourceLabel_(rec),
    "STATUS": status,
    "TAHAP KESUKARAN": _linelistingText_(rec.tahap_kesukaran),
    "PENGADU": _linelistingText_(rec.nama_pengadu),
    "PREMIS": _linelistingText_(rec.premis_nama || _derivePremisName_(rec.lokasi, rec.tajuk)),
    "LOKASI": _linelistingText_(rec.lokasi),
    "PEGAWAI PENYIASAT": _linelistingText_(rec.assigned_to),
    "PARLIMEN": _linelistingText_(rec.report_parlimen),
    "KATEGORI": _linelistingText_(rec.report_kategori),
    "JENIS PREMIS": _linelistingText_(rec.report_jenis_premis),
    "RUMUSAN": _linelistingText_(rec.report_rumusan),
    "MARKAH PEMERIKSAAN": _linelistingText_(rec.report_markah_pemeriksaan_semasa),
    "TINDAKAN": _linelistingText_(rec.report_tindakan_penguatkuasaan)
  };
}

function _getLinelistingRows_() {
  const complaints = _getComplaintRows_();
  return complaints.map(function(rec, idx) {
    return _linelistingBuildRow_(rec, idx + 1);
  });
}

function getLinelistingData(token) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };

  return {
    ok: true,
    headers: LINELISTING_HEADERS.slice(),
    rows: _getLinelistingRows_()
  };
}

function _getOrCreateLinelistingExportFolder_() {
  const folderName = "SLAM_LINELISTING_EXPORT";
  const it = DriveApp.getFoldersByName(folderName);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(folderName);
}

function _formatLinelistingSheet_(sh) {
  const widths = [
    60, 130, 115, 115, 120, 120, 110, 120, 110, 130,
    180, 220, 260, 180, 120, 150, 160, 130, 140, 260
  ];

  widths.forEach(function(w, i) {
    sh.setColumnWidth(i + 1, w);
  });

  sh.getRange(1, 1, 1, LINELISTING_HEADERS.length)
    .setBackground("#dbeafe")
    .setFontWeight("bold")
    .setFontFamily("Calibri")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, "#93c5fd", SpreadsheetApp.BorderStyle.SOLID_THIN);

  sh.setRowHeight(1, 42);

  const lastRow = Math.max(sh.getLastRow(), 1);
  const bodyRange = sh.getRange(2, 1, Math.max(lastRow - 1, 1), LINELISTING_HEADERS.length);
  bodyRange
    .setFontFamily("Calibri")
    .setFontSize(10)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("left")
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(2, 1, Math.max(lastRow - 1, 1), 1).setHorizontalAlignment("center");
  sh.getRange(2, 3, Math.max(lastRow - 1, 1), 8).setHorizontalAlignment("center");

  sh.setFrozenRows(1);
  sh.getRange(1, 1, lastRow, LINELISTING_HEADERS.length).createFilter();
}

function downloadLinelistingExcel(token) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };

  const rows = _getLinelistingRows_();
  const temp = SpreadsheetApp.create("LINELISTING_EXPORT_" + Date.now());
  const ssId = temp.getId();
  const sh = temp.getSheets()[0];
  sh.setName("Linelisting");

  sh.getRange(1, 1, 1, LINELISTING_HEADERS.length).setValues([LINELISTING_HEADERS]);

  if (rows.length) {
    const values = rows.map(function(r) {
      return LINELISTING_HEADERS.map(function(h) { return r[h] || ""; });
    });
    sh.getRange(2, 1, values.length, LINELISTING_HEADERS.length).setValues(values);
  }

  _formatLinelistingSheet_(sh);
  SpreadsheetApp.flush();

  const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?format=xlsx";
  const params = {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
  const resp = UrlFetchApp.fetch(url, params);

  if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300) {
    try { DriveApp.getFileById(ssId).setTrashed(true); } catch (e) {}
    return { ok: false, message: "Fail Excel linelisting tidak berjaya dijana." };
  }

  const folder = _getOrCreateLinelistingExportFolder_();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  const blob = resp.getBlob().setName("LINELISTING_SLAM_" + stamp + ".xlsx");
  const file = folder.createFile(blob);

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}

  try { DriveApp.getFileById(ssId).setTrashed(true); } catch (e) {}

  return {
    ok: true,
    url: "https://drive.google.com/file/d/" + file.getId() + "/view",
    file_id: file.getId()
  };
}
