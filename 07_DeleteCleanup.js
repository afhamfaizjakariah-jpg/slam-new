/******************************
 * 07_DeleteCleanup.gs
 ******************************/

function _trashDriveFileByIdSafe_(fileId) {
  const id = String(fileId || "").trim();
  if (!id) return { ok: false, skipped: true, reason: "empty_id" };

  try {
    const file = DriveApp.getFileById(id);
    file.setTrashed(true);
    return { ok: true, id: id };
  } catch (e) {
    Logger.log("Gagal trash file [" + id + "]: " + (e && e.message ? e.message : e));
    return {
      ok: false,
      id: id,
      error: e && e.message ? e.message : String(e)
    };
  }
}

function _extractDriveFileIdFromUrl_(url) {
  const s = String(url || "").trim();
  if (!s) return "";

  let m = s.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m && m[1]) return m[1];

  m = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m && m[1]) return m[1];

  return "";
}

function _deleteComplaintArtifacts_(rec) {
  rec = rec || {};

  const fileIds = [];

  const pushId = (v) => {
    const id = String(v || "").trim();
    if (id && fileIds.indexOf(id) === -1) fileIds.push(id);
  };

  const pushUrlId = (url) => {
    const id = _extractDriveFileIdFromUrl_(url);
    if (id && fileIds.indexOf(id) === -1) fileIds.push(id);
  };

  // Surat lantikan
  pushId(rec.appointment_letter_pdf_file_id);
  pushId(rec.appointment_letter_doc_id);
  pushUrlId(rec.appointment_letter_pdf_url);

  // PDF sumber asal (future-safe, jika nanti disimpan)
  pushId(rec.source_pdf_file_id);
  pushUrlId(rec.source_pdf_url);

  // Fail extraction sementara (future-safe, jika nanti disimpan)
  pushId(rec.extraction_temp_pdf_file_id);
  pushId(rec.extraction_temp_doc_id);

  const results = fileIds.map(id => _trashDriveFileByIdSafe_(id));

  return {
    total: fileIds.length,
    deleted: results.filter(r => r && r.ok).map(r => r.id),
    failed: results.filter(r => r && !r.ok && !r.skipped)
  };
}

function deleteComplaintCard(token, complaintId) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };
  _requireRole_(s, CONFIG.ROLE_ALLOW_DELETE);

  const found = _findComplaintById_(complaintId);
  if (!found || !found.record) {
    return { ok: false, message: "Rekod aduan tidak dijumpai." };
  }

  const rec = found.record || {};
  const cleanup = _deleteComplaintArtifacts_(rec);

  const failedCount = (cleanup.failed || []).length;
  if (failedCount > 0) {
    return {
      ok: false,
      message: "Rekod tidak dipadam kerana terdapat " + failedCount + " fail berkaitan yang gagal dipadam. Sila semak dahulu fail di Google Drive.",
      deleted_files: cleanup.deleted || [],
      failed_files: cleanup.failed || []
    };
  }

  found.sheet.deleteRow(found.row);

  const deletedCount = (cleanup.deleted || []).length;
  let message = "Rekod aduan berjaya dipadam.";
  if (deletedCount > 0) {
    message += " " + deletedCount + " fail berkaitan turut dipadam.";
  }

  return {
    ok: true,
    message: message,
    deleted_files: cleanup.deleted || [],
    failed_files: []
  };
}