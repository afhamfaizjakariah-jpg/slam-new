/******************************
 * 18_Reten.gs
 ******************************/

const RETEN_ALLOWED_ROLES = ["ADMIN", "PENTADBIR SISTEM", "PEGAWAI PENYEMAK"];
const RETEN_HEADERS = [
  "BIL",
  "BULAN",
  "TARIKH ADUAN / PENUGASAN",
  "TARIKH TERIMA PENUGASAN",
  "NO TIKET",
  "TAJUK ADUAN",
  "PENUGASAN (JKN/PKD/PKK/PKB)",
  "ADUAN DI BAWAH KAWASAN PARLIMEN",
  "SUMBER",
  "TAHAP KESUKARAN",
  "<15 HARI",
  ">15 HARI",
  "STATUS",
  "BERASAS",
  "TIDAK BERASAS",
  "TIDAK BERKAITAN",
  "JENIS MAKLUMBALAS AWAM",
  "KATEGORI",
  "JENIS PREMIS",
  "SUB KATEGORI PREMIS MAKANAN",
  "SUB KATEGORI PRODUK MAKANAN",
  "STATUS PENSIJILAN",
  "STATUS PEMERIKSAAN PREMIS TERDAHULU",
  "TARIKH SIASATAN",
  "MARKAH PEMERIKSAAN SEMASA",
  "TINDAKAN YANG TELAH DIBUAT (TUTUP / NOTIS 32B DLL  NYATAKAN) ",
  "CATATAN"
];

function _retenRequireAccess_(session) {
  _requireRole_(session, RETEN_ALLOWED_ROLES);
}

function _retenText_(v) {
  return String(v || "").trim();
}

function _retenYesNo_(cond) {
  return cond ? "Ya" : "Tidak";
}

function _retenMonthName_(val) {
  const d = _safeDate_(val) || _safeDate_(_normalizeDate_(val));
  if (!d) return "";
  const bulan = [
    "Januari", "Februari", "Mac", "April", "Mei", "Jun",
    "Julai", "Ogos", "September", "Oktober", "November", "Disember"
  ];
  return bulan[d.getMonth()];
}

function _retenSourceLabel_(rec) {
  const raw = _retenText_(rec.source || rec.report_jenis_aduan || rec.jenis_maklumbalas || rec.sumber_aduan).toUpperCase();
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
  return _retenText_(rec.source || rec.report_jenis_aduan || rec.jenis_maklumbalas || rec.sumber_aduan);
}

function _retenStatus_(rec) {
  return _retenText_(_computeEffectiveCardStatus_(rec) || rec.status || "");
}

function _retenCompletionDate_(rec) {
  return (
    _retenText_(rec.report_generated_at) ||
    _retenText_(rec.report_submitted_at) ||
    _retenText_(rec.report_updated_at) ||
    _retenText_(rec.status_updated_at) ||
    ""
  );
}

function _retenDiffDays_(startVal, endVal) {
  const start = _safeDate_(startVal) || _safeDate_(_normalizeDate_(startVal));
  const end = _safeDate_(endVal) || _safeDate_(_normalizeDate_(endVal));
  if (!start || !end) return null;
  const ms = end.getTime() - start.getTime();
  if (ms < 0) return null;
  return Math.floor(ms / 86400000);
}

function _retenBuildRow_(rec, bil) {
  const tarikhAduan = _formatDateMalay_(rec.tarikh_terima || "");
  const tarikhPenugasan = _formatDateMalay_(rec.assigned_at || "");
  const tarikhSiasatan = _formatDateMalay_(rec.report_tarikh_siasatan || "");
  const markah = _retenText_(rec.report_markah_pemeriksaan_semasa);
  const rumusan = _retenText_(rec.report_rumusan).toUpperCase();
  const selesaiAt = _retenCompletionDate_(rec);
  const diffDays = _retenDiffDays_(rec.tarikh_terima, selesaiAt);

  return {
    "BIL": bil,
    "BULAN": _retenMonthName_(rec.tarikh_terima || rec.created_at || ""),
    "TARIKH ADUAN / PENUGASAN": tarikhAduan,
    "TARIKH TERIMA PENUGASAN": tarikhPenugasan,
    "NO TIKET": _retenText_(rec.id_maklumbalas || rec.complaint_id),
    "TAJUK ADUAN": _retenText_(rec.tajuk),
    "PENUGASAN (JKN/PKD/PKK/PKB)": "PKD Kepong",
    "ADUAN DI BAWAH KAWASAN PARLIMEN": _retenText_(rec.report_parlimen),
    "SUMBER": _retenSourceLabel_(rec),
    "TAHAP KESUKARAN": _retenText_(rec.tahap_kesukaran),
    "<15 HARI": _retenYesNo_(diffDays !== null && diffDays < 15),
    ">15 HARI": _retenYesNo_(diffDays !== null && diffDays > 15),
    "STATUS": _retenStatus_(rec),
    "BERASAS": _retenYesNo_(rumusan === "BERASAS"),
    "TIDAK BERASAS": _retenYesNo_(rumusan === "TIDAK BERASAS"),
    "TIDAK BERKAITAN": _retenYesNo_(rumusan === "TIDAK BERKAITAN"),
    "JENIS MAKLUMBALAS AWAM": _retenText_(rec.report_jenis_maklumbalas_awam),
    "KATEGORI": _retenText_(rec.report_kategori),
    "JENIS PREMIS": _retenText_(rec.report_jenis_premis),
    "SUB KATEGORI PREMIS MAKANAN": _retenText_(rec.report_subkategori_premis_makanan),
    "SUB KATEGORI PRODUK MAKANAN": _retenText_(rec.report_subkategori_produk_makanan),
    "STATUS PENSIJILAN": _retenText_(rec.report_status_pensijilan),
    "STATUS PEMERIKSAAN PREMIS TERDAHULU": _retenText_(rec.report_status_pemeriksaan_terdahulu),
    "TARIKH SIASATAN": tarikhSiasatan,
    "MARKAH PEMERIKSAAN SEMASA": markah,
    "TINDAKAN YANG TELAH DIBUAT (TUTUP / NOTIS 32B DLL  NYATAKAN) ": _retenText_(rec.report_tindakan_penguatkuasaan),
    "CATATAN": ""
  };
}

function _getRetenRows_() {
  const complaints = _getComplaintRows_();
  return complaints.map(function(rec, idx) {
    return _retenBuildRow_(rec, idx + 1);
  });
}

function getRetenData(token) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };
  _retenRequireAccess_(s);

  return {
    ok: true,
    headers: RETEN_HEADERS.slice(),
    rows: _getRetenRows_()
  };
}

function _getOrCreateRetenExportFolder_() {
  const folderName = "SLAM_RETEN_EXPORT";
  const it = DriveApp.getFoldersByName(folderName);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(folderName);
}

function _formatRetenSheet_(sh) {
  const widths = [
    65, 120, 150, 145, 125, 260, 240, 170, 120, 135, 90, 90, 120,
    110, 120, 130, 170, 130, 145, 180, 180, 145, 190, 130, 145, 260, 120
  ];

  widths.forEach(function(w, i) {
    sh.setColumnWidth(i + 1, w);
  });

  sh.getRange(1, 1, 1, RETEN_HEADERS.length)
    .setBackground("#ffff00")
    .setFontWeight("bold")
    .setFontFamily("Calibri")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_THIN);

  sh.setRowHeight(1, 60);

  const lastRow = Math.max(sh.getLastRow(), 1);
  const bodyRange = sh.getRange(2, 1, Math.max(lastRow - 1, 1), RETEN_HEADERS.length);
  bodyRange
    .setFontFamily("Calibri")
    .setFontSize(10)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("left")
    .setWrap(false)
    .setBorder(true, true, true, true, true, true, "#d9d9d9", SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(2, 1, Math.max(lastRow - 1, 1), 1).setHorizontalAlignment("center");
  sh.getRange(2, 3, Math.max(lastRow - 1, 1), 2).setHorizontalAlignment("center");
  sh.getRange(2, 10, Math.max(lastRow - 1, 1), 17).setHorizontalAlignment("center");

  sh.setFrozenRows(1);
  sh.getRange(1, 1, lastRow, RETEN_HEADERS.length).createFilter();
}

function downloadRetenExcel(token) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };
  _retenRequireAccess_(s);

  const rows = _getRetenRows_();
  const temp = SpreadsheetApp.create("RETEN_EXPORT_" + Date.now());
  const ssId = temp.getId();
  const sh = temp.getSheets()[0];
  sh.setName("Reten");

  sh.getRange(1, 1, 1, RETEN_HEADERS.length).setValues([RETEN_HEADERS]);

  if (rows.length) {
    const values = rows.map(function(r) {
      return RETEN_HEADERS.map(function(h) { return r[h] || ""; });
    });
    sh.getRange(2, 1, values.length, RETEN_HEADERS.length).setValues(values);
  }

  _formatRetenSheet_(sh);
  SpreadsheetApp.flush();

  const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?format=xlsx";
  const params = {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
  const resp = UrlFetchApp.fetch(url, params);

  if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300) {
    try { DriveApp.getFileById(ssId).setTrashed(true); } catch (e) {}
    return { ok: false, message: "Fail Excel reten tidak berjaya dijana." };
  }

  const folder = _getOrCreateRetenExportFolder_();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  const blob = resp.getBlob().setName("RETEN_UKKM_KEPONG_" + stamp + ".xlsx");
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
