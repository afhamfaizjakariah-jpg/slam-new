/******************************
 * 04_Complaints.gs
 ******************************/

function _requireRole_(session, allowedRoles) {
  const role = _normalizeRole_(session && session.role);
  const ok = (allowedRoles || []).some(r => _normalizeRole_(r) === role);
  if (!ok) throw new Error("Akses tidak dibenarkan untuk fungsi ini.");
}

function _normalizeRole_(role) {
  const r = String(role || "").trim().toUpperCase();
  if (r === "PENTADBIR SISTEM") return "ADMIN";
  return r;
}

function _getComplaintsSheet_(autoCreate) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  const candidates = []
    .concat(CONFIG.COMPLAINTS_SHEET_CANDIDATES || [])
    .concat([CONFIG.COMPLAINTS_SHEET || "COMPLAINTS"])
    .filter((v, i, a) => v && a.indexOf(v) === i);

  for (const name of candidates) {
    const sh = ss.getSheetByName(name);
    if (sh) return sh;
  }

  if (autoCreate) {
    const sh = ss.insertSheet(CONFIG.COMPLAINTS_SHEET);
    sh.appendRow(CONFIG.COMPLAINTS_HEADERS);
    return sh;
  }

  return null;
}

function _ensureComplaintsHeader_(sh) {
  const need = CONFIG.COMPLAINTS_HEADERS.slice();
  const lastCol = Math.max(need.length, sh.getLastColumn(), 1);
  const existing = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || "").trim());

  let changed = false;
  need.forEach((h, i) => {
    if (String(existing[i] || "").trim() !== h) {
      existing[i] = h;
      changed = true;
    }
  });

  if (changed) sh.getRange(1, 1, 1, need.length).setValues([existing.slice(0, need.length)]);
  return need;
}

function _headerMap_(headers) {
  const map = {};
  (headers || []).forEach((h, i) => map[String(h || "").trim()] = i);
  return map;
}

function _getComplaintRows_() {
  const sh = _getComplaintsSheet_(true);
  const headers = _ensureComplaintsHeader_(sh);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function _normKey_(v) {
  return String(v || "")
    .replace(/\u00A0/g, " ")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function _sameComplaintKey_(inputKey, value) {
  const a = _normKey_(inputKey);
  const b = _normKey_(value);
  if (!a || !b) return false;
  if (a === b) return true;

  const na = _normalizeIdMaklumBalas_(a);
  const nb = _normalizeIdMaklumBalas_(b);
  return !!(na && nb && na === nb);
}

function _findComplaintById_(complaintId) {
  const key = String(complaintId || "").trim();
  if (!key) return null;

  const sh = _getComplaintsSheet_(true);
  const headers = _ensureComplaintsHeader_(sh);
  const map = _headerMap_(headers);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();

  for (let i = 0; i < values.length; i++) {
    const rowVals = values[i];
    const record = {};
    headers.forEach((h, idx) => record[h] = rowVals[idx]);

    if (
      _sameComplaintKey_(key, record.complaint_id) ||
      _sameComplaintKey_(key, record.kad_id) ||
      _sameComplaintKey_(key, record.id_maklumbalas)
    ) {
      return {
        sheet: sh,
        row: i + 2,
        headers: headers,
        record: record
      };
    }
  }

  const tryCols = ["complaint_id", "kad_id", "id_maklumbalas"];
  for (const colName of tryCols) {
    const idx = map[colName];
    if (idx < 0) continue;

    try {
      const finder = sh.createTextFinder(String(key).trim()).matchCase(false);
      const ranges = finder.findAll() || [];
      for (const rg of ranges) {
        if (rg.getColumn() !== idx + 1) continue;

        const cellVal = String(rg.getValue() || "");
        if (_sameComplaintKey_(key, cellVal)) {
          const rowVals = sh.getRange(rg.getRow(), 1, 1, headers.length).getValues()[0];
          const record = {};
          headers.forEach((h, j) => record[h] = rowVals[j]);

          return {
            sheet: sh,
            row: rg.getRow(),
            headers: headers,
            record: record
          };
        }
      }
    } catch (e) {}
  }

  return null;
}

function _setCellByHeader_(sh, row, map, headerName, value) {
  const idx = map[headerName];
  if (idx >= 0) sh.getRange(row, idx + 1).setValue(value);
}

function _newComplaintId_() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return "C-" +
    d.getFullYear() +
    pad(d.getMonth() + 1) +
    pad(d.getDate()) + "-" +
    pad(d.getHours()) +
    pad(d.getMinutes()) +
    pad(d.getSeconds()) + "-" +
    Utilities.getUuid().slice(0, 4).toUpperCase();
}

function _createComplaintCardId_(complaintId) {
  return "KAD-" + String(complaintId || "").replace(/[^A-Z0-9\-]/gi, "");
}

function _sanitizeComplaintRecord_(record) {
  const pick = (k) => String(record[k] || "").trim();
  return {
    id_maklumbalas: pick("id_maklumbalas"),
    tarikh_terima: pick("tarikh_terima"),
    jenis_maklumbalas: pick("jenis_maklumbalas"),
    tajuk: pick("tajuk"),
    lokasi: pick("lokasi"),
    tahap_kesukaran: pick("tahap_kesukaran"),
    nama_pengadu: pick("nama_pengadu"),
    ringkasan_butiran: pick("ringkasan_butiran"),
    raw_text: pick("raw_text"),
    parse_notes: pick("parse_notes"),
    source_pdf_file_id: pick("source_pdf_file_id"),
    source_pdf_url: pick("source_pdf_url")
  };
}

function _normalizeDate_(input) {
  const s = String(input || "").trim();
  if (!s) return "";

  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    const d = new Date(m[1] + "-" + m[2] + "-" + m[3] + "T00:00:00Z");
    if (!isNaN(d.getTime())) return m[1] + "-" + m[2] + "-" + m[3];
  }

  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    let dd = parseInt(m[1], 10);
    let mm = parseInt(m[2], 10);
    let yy = parseInt(m[3], 10);
    if (yy < 100) yy = 2000 + yy;
    if (dd >= 1 && dd <= 31 && mm >= 1 && mm <= 12 && yy >= 2000 && yy <= 2100) {
      const pad = (n) => String(n).padStart(2, "0");
      return yy + "-" + pad(mm) + "-" + pad(dd);
    }
  }

  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    const pad = (n) => String(n).padStart(2, "0");
    return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
  }
  return "";
}

function _addDaysIso_(dateObj, days) {
  const d = new Date(dateObj.getTime());
  d.setDate(d.getDate() + Number(days || 0));
  const pad = (n) => String(n).padStart(2, "0");
  return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
}

function _safeDate_(val) {
  if (!val) return null;
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

function _formatDateMalay_(val) {
  const d = _safeDate_(val);
  if (!d) {
    const iso = _normalizeDate_(val);
    if (!iso) return "";
    const dd = new Date(iso + "T00:00:00");
    if (isNaN(dd.getTime())) return "";
    return Utilities.formatDate(dd, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

function _derivePremisName_(lokasi, tajuk) {
  const l = String(lokasi || "").trim();
  if (l) {
    const first = l.split(/,|\n|;/)[0].trim();
    if (first) return first;
  }
  const t = String(tajuk || "").trim();
  if (!t) return "";
  return t.length > 80 ? t.slice(0, 80).trim() : t;
}

function _normalizeCardStatus_(status) {
  const s = String(status || "").trim().toUpperCase();
  if (!s) return "";
  if (s === "BAHARU" || s === "BARU") return "Baharu";
  if (s === "TELAH DIAGIHKAN" || s === "DIAGIHKAN") return "Telah Diagihkan";
  if (s === "PENDING") return "Pending";
  if (s === "LEWAT") return "Lewat";
  if (s === "SELESAI") return "Selesai";
  if (s === "PINDAH") return "Pindah";
  return "";
}

function _mapCardStatusToMainStatus_(cardStatus) {
  cardStatus = _normalizeCardStatus_(cardStatus);
  if (cardStatus === "Baharu") return "Baharu";
  if (cardStatus === "Telah Diagihkan") return "Dalam Siasatan";
  if (cardStatus === "Pending") return "Pending";
  if (cardStatus === "Lewat") return "Lewat";
  if (cardStatus === "Selesai") return "Selesai";
  if (cardStatus === "Pindah") return "Pindah";
  return "Baharu";
}

function _computeEffectiveCardStatus_(r) {
  let st = _normalizeCardStatus_(r.card_status || r.status || "Baharu") || "Baharu";
  const due = _normalizeDate_(r.due_date || "");
  const done = st === "Selesai" || st === "Pindah";
  if (!done && due) {
    const today = _normalizeDate_(new Date());
    if (today && today > due && (st === "Pending" || st === "Telah Diagihkan" || st === "Baharu")) {
      st = "Lewat";
    }
  }
  return st;
}

function _formatCardRecord_(r) {
  const statusEff = _computeEffectiveCardStatus_(r);
  return {
    complaint_id: String(r.complaint_id || ""),
    id_maklumbalas: String(r.id_maklumbalas || ""),
    tajuk: String(r.tajuk || ""),
    premis_nama: String(r.premis_nama || _derivePremisName_(r.lokasi, r.tajuk) || ""),
    tarikh_jana: _formatDateMalay_(r.generated_card_at || r.created_at || ""),
    generated_card_at: String(r.generated_card_at || r.created_at || ""),
    tarikh_terima: _formatDateMalay_(r.tarikh_terima || ""),
    status_card: statusEff,
    assigned_to: String(r.assigned_to || ""),
    assigned_user_id: String(r.assigned_user_id || ""),
    assigned_role: String(r.assigned_role || ""),
    lokasi: String(r.lokasi || ""),
    nama_pengadu: String(r.nama_pengadu || ""),
    source: String(r.source || r.jenis_maklumbalas || ""),
    kesukaran: String(r.tahap_kesukaran || ""),
    ringkasan: String(r.ringkasan_butiran || ""),
    due_date: _formatDateMalay_(r.due_date || ""),
    kad_id: String(r.kad_id || ""),
    appointment_letter_pdf_url: String(r.appointment_letter_pdf_url || "")
  };
}

/** ========= KAD ADUAN ========= **/
function listComplaintCards(token) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };

  const rows = _getComplaintRows_();
  const cards = rows
    .map(r => _formatCardRecord_(r))
    .sort((a, b) => {
      const da = _safeDate_(a.generated_card_at || a.created_at);
      const db = _safeDate_(b.generated_card_at || b.created_at);
      return (db ? db.getTime() : 0) - (da ? da.getTime() : 0);
    });

  return { ok: true, cards: cards };
}

function getComplaintDetail(token, complaintId) {
  try {
    Logger.log("=== getComplaintDetail START ===");
    Logger.log("complaintId = " + complaintId);

    const s = _getSession_(token);
    if (!s) {
      return { ok: false, message: "Sesi tamat. Sila log masuk semula." };
    }

    const found = _findComplaintById_(complaintId);
    Logger.log("found = " + JSON.stringify(found ? {
      row: found.row,
      sheet: found.sheet ? found.sheet.getName() : "",
      hasRecord: !!found.record
    } : null));

    if (!found || !found.record) {
      return {
        ok: false,
        message: "Rekod aduan tidak dijumpai. Key dicari: " + String(complaintId || "")
      };
    }

    const rec = found.record || {};

    const detail = {
      complaint_id: String(rec.complaint_id || ""),
      id_maklumbalas: String(rec.id_maklumbalas || ""),
      tajuk: String(rec.tajuk || ""),
      premis_nama: String(rec.premis_nama || _derivePremisName_(rec.lokasi, rec.tajuk) || ""),
      tarikh_jana: _formatDateMalay_(rec.generated_card_at || rec.created_at || ""),
      generated_card_at: String(rec.generated_card_at || rec.created_at || ""),
      tarikh_terima: _formatDateMalay_(rec.tarikh_terima || ""),
      status_card: String(_computeEffectiveCardStatus_(rec) || "Baharu"),
      assigned_to: String(rec.assigned_to || ""),
      assigned_user_id: String(rec.assigned_user_id || ""),
      assigned_role: String(rec.assigned_role || ""),
      lokasi: String(rec.lokasi || ""),
      nama_pengadu: String(rec.nama_pengadu || ""),
      source: String(rec.source || rec.jenis_maklumbalas || ""),
      kesukaran: String(rec.tahap_kesukaran || ""),
      ringkasan: String(rec.ringkasan_butiran || ""),
      ringkasan_butiran: String(rec.ringkasan_butiran || ""),
      due_date: _formatDateMalay_(rec.due_date || ""),
      kad_id: String(rec.kad_id || ""),
      raw_text: String(rec.raw_text || ""),
      parse_notes: String(rec.parse_notes || ""),
      source_pdf_file_id: String(rec.source_pdf_file_id || ""),
      source_pdf_url: String(rec.source_pdf_url || ""),
      tahap_kesukaran: String(rec.tahap_kesukaran || ""),
      status: String(rec.status || ""),
      status_updated_at: String(rec.status_updated_at || ""),
      assigned_at: String(rec.assigned_at || ""),
      created_by: String(rec.created_by || ""),
      created_at: String(rec.created_at || ""),

      appointment_letter_ref_no: String(rec.appointment_letter_ref_no || ""),
      appointment_letter_generated_at: String(rec.appointment_letter_generated_at || ""),
      appointment_letter_pdf_file_id: String(rec.appointment_letter_pdf_file_id || ""),
      appointment_letter_pdf_url: String(rec.appointment_letter_pdf_url || ""),
      appointment_letter_doc_id: String(rec.appointment_letter_doc_id || ""),

      report_status: String(rec.report_status || ""),
      report_pdf_file_id: String(rec.report_pdf_file_id || ""),
      report_pdf_url: String(rec.report_pdf_url || ""),
      report_doc_id: String(rec.report_doc_id || ""),
      report_generated_at: String(rec.report_generated_at || "")
    };

    Logger.log("detail = " + JSON.stringify(detail));

    return {
      ok: true,
      detail: detail
    };

  } catch (e) {
    Logger.log("getComplaintDetail ERROR = " + (e && e.stack ? e.stack : e));
    return {
      ok: false,
      message: "getComplaintDetail error: " + (e && e.message ? e.message : String(e))
    };
  }
}

function updateComplaintCardStatus(token, complaintId, newStatus) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };
  _requireRole_(s, CONFIG.ROLE_ALLOW_ASSIGN);

  newStatus = _normalizeCardStatus_(newStatus);
  if (!newStatus) return { ok: false, message: "Status kad tidak sah." };

  const found = _findComplaintById_(complaintId);
  if (!found) return { ok: false, message: "Rekod aduan tidak dijumpai." };

  const sh = found.sheet;
  const row = found.row;
  const headers = found.headers;
  const map = _headerMap_(headers);
  const nowIso = new Date().toISOString();

  _setCellByHeader_(sh, row, map, "card_status", newStatus);
  _setCellByHeader_(sh, row, map, "status_updated_at", nowIso);
  _setCellByHeader_(sh, row, map, "status", _mapCardStatusToMainStatus_(newStatus));

  return { ok: true, message: "Status kad berjaya dikemas kini." };
}

function getComplaintReportSummary(token, complaintId) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };

  const found = _findComplaintById_(complaintId);
  if (!found || !found.record) {
    return { ok: false, message: "Rekod aduan tidak dijumpai." };
  }

  const rec = found.record || {};
  return {
    ok: true,
    summary: {
      markah_pemeriksaan_semasa: String(rec.report_markah_pemeriksaan_semasa || "").trim(),
      rumusan: String(rec.report_rumusan || "").trim()
    }
  };
}
