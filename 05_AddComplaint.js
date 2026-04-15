/******************************
 * 05_AddComplaint.gs
 ******************************/

/** ========= TAMBAH ADUAN ========= **/
function extractComplaintFromPdf(token, base64Pdf, filename) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };
  _requireRole_(s, CONFIG.ROLE_ALLOW_TAMBAH);

  filename = String(filename || "").trim() || "aduan.pdf";
  base64Pdf = String(base64Pdf || "").trim();
  if (!base64Pdf) return { ok: false, message: "Fail PDF tidak diterima." };

  let bytes;
  try {
    bytes = Utilities.base64Decode(base64Pdf);
  } catch (e) {
    return { ok: false, message: "Format base64 tidak sah." };
  }

  let sourcePdf = { file_id: "", url: "" };
  try {
    sourcePdf = _storeSourcePdf_(bytes, filename);
  } catch (e) {
    sourcePdf = { file_id: "", url: "" };
  }

  let extracted;
  if (CONFIG.ALLOW_TEMP_DRIVE_EXTRACTION) {
    extracted = _extractPdfTextViaDrive_(bytes, filename);
    if (!extracted.text) {
      const fallback = _bestEffortExtractPdfText_(bytes);
      extracted.notes = (extracted.notes || []).concat(fallback.notes || []);
      extracted.text = fallback.text || "";
    }
  } else {
    extracted = _bestEffortExtractPdfText_(bytes);
  }

  const rawText = String(extracted.text || "").trim();
  const parseNotes = (extracted.notes || []).filter(Boolean);

  parseNotes.push("SPREADSHEET_ID=" + CONFIG.SPREADSHEET_ID);
  parseNotes.push("COMPLAINTS_SHEET=" + CONFIG.COMPLAINTS_SHEET);

  if (!rawText) {
    return {
      ok: true,
      data: _emptyExtractData_(filename),
      duplicate: false,
      extract_ok: false,
      message: "PDF tidak boleh diekstrak. Sila lengkapkan maklumat secara manual (mungkin scan/encoding/permission).",
      raw_text: "",
      parse_notes: parseNotes.join(" | ")
    };
  }

  const mapped = _parseComplaintText_(rawText, filename);
  if (mapped.parse_notes) parseNotes.push(mapped.parse_notes);

  let inferredSumber = _inferSumberFromId_(mapped.id_maklumbalas) || _inferSumberFromId_(rawText);
  if (!inferredSumber) {
    const j = String(mapped.jenis_maklumbalas || "").trim().toUpperCase();
    if (j === "MOH" || j === "PCB" || j === "JPA" || j === "EMEL") inferredSumber = j;
  }
  mapped.sumber_sistem = String(inferredSumber || "").trim();
  parseNotes.push("Sumber inferred: " + (mapped.sumber_sistem || "(blank)"));

  let butiranForSummary = String(mapped.butiran || "").trim();
  if (!butiranForSummary) {
    butiranForSummary = _extractButiranSmart_(rawText) || "";
    if (butiranForSummary) parseNotes.push("Butiran fallback smart extracted.");
  }
  const summaryInput = (butiranForSummary || "").trim() || rawText;

  let ringkasan = "";
  if (CONFIG.GROQ_SUMMARY_ENABLED && summaryInput) {
    const g1 = _summarizeButiranWithGroq_(summaryInput, filename, { strict: false, notes: parseNotes });
    if (g1 && g1.ok && g1.text) {
      ringkasan = g1.text;
      parseNotes.push("Groq summary ok.");
    } else {
      parseNotes.push("Groq summary failed.");
      if (g1 && g1.note) parseNotes.push(g1.note);

      const g2 = _summarizeButiranWithGroq_(summaryInput, filename, { strict: true, notes: parseNotes });
      if (g2 && g2.ok && g2.text) {
        ringkasan = g2.text;
        parseNotes.push("Groq retry ok.");
      } else {
        parseNotes.push("Groq retry failed.");
        if (g2 && g2.note) parseNotes.push(g2.note);
      }
    }
  }

  if (!ringkasan) {
    ringkasan = _makeNeutralSummary_(summaryInput);
    parseNotes.push("Heuristic PRO summary used.");
  }

  mapped.ringkasan_butiran = ringkasan;
  mapped.tahap_kesukaran = mapped.tahap_kesukaran || "Biasa";
  mapped.jenis_maklumbalas = mapped.jenis_maklumbalas || _guessJenisFromFilename_(filename) || "";
  delete mapped.butiran;

  const idKey = String(mapped.id_maklumbalas || "").trim();
  const dup = idKey ? checkDuplicate(token, idKey) : { ok: true, duplicate: false };
  if (dup && dup.debug) parseNotes.push("dup_debug=" + JSON.stringify(dup.debug).slice(0, 800));

  const sumberVal = mapped.sumber_sistem || mapped.jenis_maklumbalas || "";

  return {
    ok: true,
    data: {
      sumber_sistem: sumberVal,
      source: sumberVal,
      sumberSistem: sumberVal,
      sumber: sumberVal,
      sumber_system: sumberVal,
      sistem_sumber: sumberVal,
      sumber_sistem_value: sumberVal,

      id_maklumbalas: mapped.id_maklumbalas || "",
      tarikh_terima: mapped.tarikh_terima || "",
      jenis_maklumbalas: mapped.jenis_maklumbalas || "",
      tajuk: mapped.tajuk || "",
      lokasi: mapped.lokasi || "",
      tahap_kesukaran: mapped.tahap_kesukaran || "Biasa",
      nama_pengadu: mapped.nama_pengadu || "",
      ringkasan_butiran: mapped.ringkasan_butiran || "",

      source_pdf_file_id: sourcePdf.file_id || "",
      source_pdf_url: sourcePdf.url || "",
    },

    duplicate: !!(dup && dup.duplicate),
    duplicate_row: (dup && dup.duplicate) ? (dup.row || null) : null,
    extract_ok: true,
    message: (dup && dup.duplicate) ? "Duplikasi dikesan: ID Maklum Balas/No Aduan telah wujud." : "Extraction berjaya.",
    raw_text: rawText,
    parse_notes: parseNotes.join(" | ")
  };
}

function _emptyExtractData_(filename) {
  return {
    sumber_sistem: "",
    source: "",
    sumberSistem: "",
    sumber: "",
    sumber_system: "",
    sistem_sumber: "",
    sumber_sistem_value: "",
    id_maklumbalas: "",
    tarikh_terima: "",
    jenis_maklumbalas: _guessJenisFromFilename_(filename) || "",
    tajuk: "",
    lokasi: "",
    tahap_kesukaran: "Biasa",
    nama_pengadu: "",
    ringkasan_butiran: ""
  };
}

function checkDuplicate(token, idMaklumBalasOrNoAduan) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };
  _requireRole_(s, CONFIG.ROLE_ALLOW_TAMBAH);

  const key = _normalizeIdMaklumBalas_(idMaklumBalasOrNoAduan);
  if (!key) return { ok: true, duplicate: false };

  const sh = _getComplaintsSheet_(true);
  const headers = _ensureComplaintsHeader_(sh);
  const map = _headerMap_(headers);
  const lastRow = sh.getLastRow();
  const lastCol = Math.max(sh.getLastColumn(), 1);

  const debug = { sheet: sh.getName(), lastRow, lastCol, idx: map.id_maklumbalas || -1 };

  if (map.id_maklumbalas >= 0 && lastRow >= 2) {
    const col = map.id_maklumbalas + 1;
    const values = sh.getRange(2, col, lastRow - 1, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      const raw = String(values[i][0] || "");
      const norm = _normalizeIdMaklumBalas_(raw);
      if (norm && norm === key) {
        return { ok: true, duplicate: true, row: i + 2, debug };
      }
    }
    return { ok: true, duplicate: false, debug };
  }

  try {
    const finder = sh.createTextFinder(String(idMaklumBalasOrNoAduan || "").trim()).matchCase(false);
    const ranges = finder.findAll() || [];
    for (const rg of ranges) {
      const v = String(rg.getValue() || "");
      if (_normalizeIdMaklumBalas_(v) === key) {
        debug.fallback = "TextFinder";
        debug.hitA1 = rg.getA1Notation();
        return { ok: true, duplicate: true, row: rg.getRow(), debug };
      }
    }
  } catch (e) {
    debug.fallback_err = String(e && e.message ? e.message : e);
  }

  debug.fallback = "none";
  return { ok: true, duplicate: false, debug };
}

function getSourcePdfUrl(token, complaintId) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };

  const found = _findComplaintById_(complaintId);
  if (!found || !found.record) {
    return { ok: false, message: "Rekod aduan tidak dijumpai." };
  }

  const rec = found.record || {};

  const directUrl = String(rec.source_pdf_url || "").trim();
  if (directUrl) {
    return { ok: true, url: directUrl };
  }

  const fileId = String(rec.source_pdf_file_id || "").trim();
  if (!fileId) {
    return { ok: false, message: "PDF aduan asal belum disimpan." };
  }

  return {
    ok: true,
    url: "https://drive.google.com/file/d/" + fileId + "/view"
  };
}

function registerComplaint(token, record) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };
  _requireRole_(s, CONFIG.ROLE_ALLOW_TAMBAH);

  record = record || {};
  const cleaned = _sanitizeComplaintRecord_(record);

  const required = [
    ["id_maklumbalas", "ID Maklum Balas / No Aduan"],
    ["tarikh_terima", "Tarikh Terima"],
    ["jenis_maklumbalas", "Jenis Maklum Balas"],
    ["tajuk", "Tajuk"],
    ["lokasi", "Lokasi"],
    ["tahap_kesukaran", "Tahap Kesukaran"],
    ["nama_pengadu", "Nama Pengadu"],
    ["ringkasan_butiran", "Ringkasan Butiran"]
  ];

  const errors = {};
  required.forEach(([k, label]) => {
    if (!String(cleaned[k] || "").trim()) errors[k] = label + " wajib diisi.";
  });

  const isoDate = _normalizeDate_(cleaned.tarikh_terima);
  if (!isoDate) errors.tarikh_terima = "Tarikh Terima tidak sah.";
  else cleaned.tarikh_terima = isoDate;

  if (Object.keys(errors).length) {
    return { ok: false, message: "Sila semak ralat pada borang.", field_errors: errors };
  }

  const dup = checkDuplicate(token, cleaned.id_maklumbalas);
  if (dup && dup.duplicate) {
    return {
      ok: false,
      message: "Duplikasi dikesan: rekod dengan ID Maklum Balas/No Aduan ini telah wujud.",
      field_errors: { id_maklumbalas: "Duplikasi dikesan." }
    };
  }

  const sh = _getComplaintsSheet_(true);
  const headers = _ensureComplaintsHeader_(sh);

  const now = new Date();
  const nowIso = now.toISOString();
  const complaintId = _newComplaintId_();
  const kadId = _createComplaintCardId_(complaintId);
  const generatedCardAt = nowIso;
  const dueDate = _addDaysIso_(now, CONFIG.CARD_DEFAULT_DUE_DAYS);
  const premisNama = _derivePremisName_(cleaned.lokasi, cleaned.tajuk);
  const sumber = String(cleaned.jenis_maklumbalas || "").trim() || "PDF";

  const rowObj = {
    complaint_id: complaintId,
    id_maklumbalas: cleaned.id_maklumbalas,
    tarikh_terima: cleaned.tarikh_terima,
    jenis_maklumbalas: cleaned.jenis_maklumbalas,
    tajuk: cleaned.tajuk,
    ringkasan_butiran: cleaned.ringkasan_butiran,
    lokasi: cleaned.lokasi,
    premis_nama: premisNama,
    tahap_kesukaran: cleaned.tahap_kesukaran,
    nama_pengadu: cleaned.nama_pengadu,
    source: sumber,
    status: "Baharu",
    card_status: "Baharu",
    assigned_to: "",
    assigned_at: "",
    status_updated_at: nowIso,
    generated_card_at: generatedCardAt,
    due_date: dueDate,
    created_at: nowIso,
    created_by: s.userId,
    kad_id: kadId,

    appointment_letter_ref_no: "",
    appointment_letter_generated_at: "",
    appointment_letter_pdf_file_id: "",
    appointment_letter_pdf_url: "",
    appointment_letter_doc_id: "",

    raw_text: cleaned.raw_text || "",
    parse_notes: cleaned.parse_notes || "",

    source_pdf_file_id: cleaned.source_pdf_file_id || "",
    source_pdf_url: cleaned.source_pdf_url || "",
    extraction_temp_pdf_file_id: "",
    extraction_temp_doc_id: "",

    assigned_user_id: "",
    assigned_role: ""
  };

  const outRow = headers.map(h => (rowObj[h] !== undefined ? rowObj[h] : ""));
  sh.appendRow(outRow);
  SpreadsheetApp.flush();

  return {
    ok: true,
    message: "Aduan berjaya didaftarkan dan Kad Aduan telah dijana.",
    complaint_id: complaintId,
    kad_id: kadId,
    id_maklumbalas: cleaned.id_maklumbalas
  };
}

function _normalizeComplaintSource_(source) {
  const s = String(source || "").trim();
  return s;
}

function _complaintSourcePrefix_(source) {
  const map = {
    "Media Sosial": "MS",
    "Emel": "EM",
    "Whatsapp Aduan": "WA",
    "Hadir Sendiri": "HS",
    "Telefon": "TL",
    "Surat": "SR",
    "Lain-Lain": "LL",
    "Platform Negeri": "PN"
  };
  return map[String(source || "").trim()] || "";
}

function _nextComplaintSourceRunningNo_(prefix) {
  const props = PropertiesService.getScriptProperties();
  const key = "SLAM_SRC_SEQ_" + prefix;
  const current = Number(props.getProperty(key) || "0");
  const next = current + 1;
  props.setProperty(key, String(next));
  return prefix + String(next).padStart(4, "0");
}

function _peekComplaintSourceRunningNo_(prefix) {
  const props = PropertiesService.getScriptProperties();
  const key = "SLAM_SRC_SEQ_" + prefix;
  const current = Number(props.getProperty(key) || "0");
  const next = current + 1;
  return prefix + String(next).padStart(4, "0");
}

function previewComplaintReference(token, source) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };
  _requireRole_(s, CONFIG.ROLE_ALLOW_TAMBAH);

  const normalized = _normalizeComplaintSource_(source);
  if (!normalized) return { ok: false, message: "Sumber Aduan belum dipilih." };
  if (normalized === "SiSPAA") return { ok: true, reference: "" };

  const prefix = _complaintSourcePrefix_(normalized);
  if (!prefix) return { ok: false, message: "Sumber Aduan tidak disokong." };

  return { ok: true, reference: _peekComplaintSourceRunningNo_(prefix) };
}

function registerComplaintV2(token, record) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };
  _requireRole_(s, CONFIG.ROLE_ALLOW_TAMBAH);

  record = record || {};
  const cleaned = _sanitizeComplaintRecord_(record);
  cleaned.jenis_maklumbalas = _normalizeComplaintSource_(cleaned.jenis_maklumbalas);

  if (!cleaned.id_maklumbalas) {
    const prefix = _complaintSourcePrefix_(cleaned.jenis_maklumbalas);
    if (prefix) cleaned.id_maklumbalas = _nextComplaintSourceRunningNo_(prefix);
  }

  const required = [
    ["jenis_maklumbalas", "Sumber Aduan"],
    ["tarikh_terima", "Tarikh Terima"],
    ["tajuk", "Tajuk"],
    ["lokasi", "Lokasi"],
    ["tahap_kesukaran", "Tahap Kesukaran"],
    ["nama_pengadu", "Nama Pengadu"],
    ["ringkasan_butiran", "Ringkasan Butiran"]
  ];

  const errors = {};
  required.forEach(function(pair) {
    const k = pair[0], label = pair[1];
    if (!String(cleaned[k] || "").trim()) errors[k] = label + " wajib diisi.";
  });

  if (cleaned.jenis_maklumbalas === "SiSPAA" && !String(cleaned.id_maklumbalas || "").trim()) {
    errors.id_maklumbalas = "ID Rujukan Aduan wajib diisi bagi sumber SiSPAA.";
  }

  const isoDate = _normalizeDate_(cleaned.tarikh_terima);
  if (!isoDate) errors.tarikh_terima = "Tarikh Terima tidak sah.";
  else cleaned.tarikh_terima = isoDate;

  if (Object.keys(errors).length) {
    return { ok: false, message: "Sila semak ralat pada borang.", field_errors: errors };
  }

  if (cleaned.id_maklumbalas) {
    const dup = checkDuplicate(token, cleaned.id_maklumbalas);
    if (dup && dup.duplicate) {
      return {
        ok: false,
        message: "Duplikasi dikesan: rekod dengan ID Maklum Balas/No Aduan ini telah wujud.",
        field_errors: { id_maklumbalas: "Duplikasi dikesan." }
      };
    }
  }

  const sh = _getComplaintsSheet_(true);
  const headers = _ensureComplaintsHeader_(sh);

  const now = new Date();
  const nowIso = now.toISOString();
  const complaintId = _newComplaintId_();
  const kadId = _createComplaintCardId_(complaintId);
  const dueDate = _addDaysIso_(now, CONFIG.CARD_DEFAULT_DUE_DAYS);
  const premisNama = _derivePremisName_(cleaned.lokasi, cleaned.tajuk);
  const sumber = String(cleaned.jenis_maklumbalas || "").trim();

  const rowObj = {
    complaint_id: complaintId,
    id_maklumbalas: cleaned.id_maklumbalas,
    tarikh_terima: cleaned.tarikh_terima,
    jenis_maklumbalas: cleaned.jenis_maklumbalas,
    tajuk: cleaned.tajuk,
    ringkasan_butiran: cleaned.ringkasan_butiran,
    lokasi: cleaned.lokasi,
    premis_nama: premisNama,
    tahap_kesukaran: cleaned.tahap_kesukaran,
    nama_pengadu: cleaned.nama_pengadu,
    source: sumber,
    status: "Baharu",
    card_status: "Baharu",
    assigned_to: "",
    assigned_at: "",
    status_updated_at: nowIso,
    generated_card_at: nowIso,
    due_date: dueDate,
    created_at: nowIso,
    created_by: s.userId,
    kad_id: kadId,
    appointment_letter_ref_no: "",
    appointment_letter_generated_at: "",
    appointment_letter_pdf_file_id: "",
    appointment_letter_pdf_url: "",
    appointment_letter_doc_id: "",
    raw_text: cleaned.raw_text || "",
    parse_notes: cleaned.parse_notes || "",
    source_pdf_file_id: cleaned.source_pdf_file_id || "",
    source_pdf_url: cleaned.source_pdf_url || "",
    extraction_temp_pdf_file_id: "",
    extraction_temp_doc_id: "",
    assigned_user_id: "",
    assigned_role: ""
  };

  const outRow = headers.map(function(h) { return rowObj[h] !== undefined ? rowObj[h] : ""; });
  sh.appendRow(outRow);
  SpreadsheetApp.flush();

  return {
    ok: true,
    message: "Aduan berjaya didaftarkan dan Kad Aduan telah dijana.",
    complaint_id: complaintId,
    kad_id: kadId,
    id_maklumbalas: cleaned.id_maklumbalas
  };
}
