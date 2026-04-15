/******************************
 * 99_Debug.gs
 ******************************/

/** ====== DEBUG ====== **/
function debugDuplicateNoAuth(idMaklumBalasOrNoAduan) {
  const key = _normalizeIdMaklumBalas_(idMaklumBalasOrNoAduan);
  if (!key) return { ok: false, message: "ID kosong." };

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = _getComplaintsSheet_(false);
  if (!sh) return { ok: false, message: "Sheet COMPLAINTS tidak wujud.", spreadsheetName: ss.getName() };

  const headers = _ensureComplaintsHeader_(sh);
  const map = _headerMap_(headers);
  const lastRow = sh.getLastRow();
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const idx = map.id_maklumbalas;

  const debug = { spreadsheetName: ss.getName(), sheetName: sh.getName(), lastRow, lastCol, idx, key };

  const matches = [];
  if (idx >= 0 && lastRow >= 2) {
    const col = idx + 1;
    const values = sh.getRange(2, col, lastRow - 1, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      const raw = String(values[i][0] || "");
      const norm = _normalizeIdMaklumBalas_(raw);
      if (norm === key) matches.push({ row: i + 2, a1: sh.getRange(i + 2, col).getA1Notation(), raw });
    }
    return { ok: true, found: matches.length > 0, debug, matches };
  }

  try {
    const finder = sh.createTextFinder(String(idMaklumBalasOrNoAduan || "").trim()).matchCase(false);
    const ranges = finder.findAll() || [];
    for (const rg of ranges) {
      const v = String(rg.getValue() || "");
      if (_normalizeIdMaklumBalas_(v) === key) matches.push({ row: rg.getRow(), a1: rg.getA1Notation(), raw: v });
    }
    debug.fallback = "TextFinder";
    return { ok: true, found: matches.length > 0, debug, matches };
  } catch (e) {
    return { ok: false, message: "TextFinder error: " + String(e && e.message ? e.message : e), debug };
  }
}

function debugFindComplaintNoAuth(key) {
  const found = _findComplaintById_(key);
  if (!found) {
    return { ok: false, key: key, message: "Tidak jumpa rekod." };
  }
  return {
    ok: true,
    key: key,
    row: found.row,
    sheet: found.sheet ? found.sheet.getName() : "",
    record: found.record
  };
}

function testGroqAuth() {
  const notes = [];
  const r = _summarizeButiranWithGroq_(
    "Pengadu memaklumkan kebanyakan gerai makanan di restoran ini dikendalikan warga asing. Pengadu mendakwa kebersihan diri pekerja dan persekitaran gerai tidak memuaskan serta terdapat kebimbangan kesihatan seperti batuk berterusan. Pengadu juga menyatakan identiti pemilik pada sistem bayaran tidak selari dengan operator sebenar.",
    "test.pdf",
    { strict: true, notes: notes }
  );
  Logger.log(JSON.stringify({ r: r, notes: notes }));
  return { r: r, notes: notes };
}

function testFindComplaint() {
  const res = debugFindComplaintNoAuth("MOH.312502");
  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

function pingServer(token) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat." };
  return { ok: true, message: "pong", user: s.userId };
}

function debugGetComplaintDetailNoAuth(complaintId) {
  try {
    const found = _findComplaintById_(complaintId);
    if (!found) {
      return {
        ok: false,
        message: "Tidak jumpa rekod.",
        complaintId: String(complaintId || "")
      };
    }

    const rec = found.record || {};
    const detail = _formatCardRecord_(rec);

    detail.raw_text = rec.raw_text || "";
    detail.parse_notes = rec.parse_notes || "";
    detail.ringkasan_butiran = rec.ringkasan_butiran || "";
    detail.lokasi = rec.lokasi || "";
    detail.nama_pengadu = rec.nama_pengadu || "";
    detail.tahap_kesukaran = rec.tahap_kesukaran || "";
    detail.source = rec.source || rec.jenis_maklumbalas || "";
    detail.status = rec.status || "";
    detail.kad_id = rec.kad_id || "";
    detail.status_updated_at = rec.status_updated_at || "";
    detail.assigned_at = rec.assigned_at || "";
    detail.created_by = rec.created_by || "";
    detail.created_at = rec.created_at || "";
    detail.due_date = _formatDateMalay_(rec.due_date || "");
    detail.assigned_user_id = rec.assigned_user_id || "";
    detail.assigned_role = rec.assigned_role || "";
    detail.cleanliness_score = rec.cleanliness_score || "";
    detail.complaint_classification = rec.complaint_classification || "";

    detail.appointment_letter_ref_no = rec.appointment_letter_ref_no || "";
    detail.appointment_letter_generated_at = rec.appointment_letter_generated_at || "";
    detail.appointment_letter_pdf_file_id = rec.appointment_letter_pdf_file_id || "";
    detail.appointment_letter_pdf_url = rec.appointment_letter_pdf_url || "";
    detail.appointment_letter_doc_id = rec.appointment_letter_doc_id || "";

    return {
      ok: true,
      row: found.row,
      sheet: found.sheet ? found.sheet.getName() : "",
      detail: detail
    };
  } catch (e) {
    return {
      ok: false,
      message: e && e.message ? e.message : String(e),
      stack: e && e.stack ? String(e.stack) : ""
    };
  }
}

function testDriveAccess() {
  const it = DriveApp.getFoldersByName("SLAM_APPOINTMENT_LETTERS");
  return it.hasNext();
}