/******************************
 * 08_AppointmentLetter.gs
 ******************************/

/** ====== SURAT LANTIKAN PEGAWAI PENYIASAT ====== **/
function _getOrCreateAppointmentFolder_() {
  const name = String(CONFIG.APPOINTMENT_OUTPUT_FOLDER_NAME || "SLAM_APPOINTMENT_LETTERS").trim();
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

function _nextAppointmentRunningNo_() {
  const props = PropertiesService.getScriptProperties();
  const current = Number(props.getProperty("SLAM_APPOINTMENT_RUNNING_NO") || "0");
  const next = current + 1;
  props.setProperty("SLAM_APPOINTMENT_RUNNING_NO", String(next));
  return next;
}

function _addBusinessDaysIso_(startDate, businessDays) {
  let d = new Date(startDate.getTime());
  let added = 0;

  while (added < Number(businessDays || 0)) {
    d.setDate(d.getDate() + 1);
    const day = d.getDay();
    if (day !== 0 && day !== 6) added++;
  }

  const pad = (n) => String(n).padStart(2, "0");
  return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
}

function _computeAppointmentDueDate_(record, assignedAt) {
  const idRef = String(record.id_maklumbalas || "").toUpperCase().trim();
  let days = 3;

  if (idRef.indexOf("MOH.") === 0) {
    days = 7;
  } else if (idRef.indexOf("PCB.") === 0 || idRef.indexOf("JPA.") === 0) {
    days = 3;
  }

  return _addBusinessDaysIso_(assignedAt, days);
}

function _replaceDocPlaceholder_(body, placeholder, value) {
  const safePattern = String(placeholder || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  body.replaceText(safePattern, _escapeForDocReplace_(value));
}

function _generateAppointmentLetterPdf_(record, chosenUser) {
  const rec = record || {};
  const chosen = chosenUser || {};

  const templateId = String(CONFIG.APPOINTMENT_DOC_TEMPLATE_ID || "").trim();
  if (!templateId) {
    throw new Error("CONFIG.APPOINTMENT_DOC_TEMPLATE_ID belum ditetapkan.");
  }

  const targetFolder = _getOrCreateAppointmentFolder_();

  const oldPdfId = String(rec.appointment_letter_pdf_file_id || "").trim();
  const oldDocId = String(rec.appointment_letter_doc_id || "").trim();

  if (oldPdfId) {
    try {
      DriveApp.getFileById(oldPdfId).setTrashed(true);
    } catch (e) {
      Logger.log("Gagal padam PDF lama: " + (e && e.message ? e.message : e));
    }
  }

  if (oldDocId) {
    try {
      DriveApp.getFileById(oldDocId).setTrashed(true);
    } catch (e) {
      Logger.log("Gagal padam Doc lama: " + (e && e.message ? e.message : e));
    }
  }

  const complaintId = String(rec.complaint_id || "").trim();
  const idMaklumBalas = String(rec.id_maklumbalas || "").trim();
  const tajukAduan = String(rec.tajuk || "").trim();

  // Ikut arahan terbaru:
  // tarikh_lantikan = 05 MAC 2026
  // tarikh_terima & due_date = dd/MM/yyyy
  const tarikhTerima = _formatDateMalay_(rec.tarikh_terima || "");
  const dueDate = _formatDateMalay_(rec.due_date || "");

  const assignedUserId = String(chosen.userId || rec.assigned_user_id || "").trim();
  const assignedProfile = assignedUserId ? _findUserFullProfile_(assignedUserId) : null;

  const pegawaiNama = String(
    (assignedProfile && (assignedProfile.full_name || assignedProfile.name)) ||
    chosen.name ||
    rec.assigned_to ||
    ""
  ).trim().toUpperCase();

  const pegawaiJawatan = String(
    (assignedProfile && assignedProfile.jawatan) ||
    rec.assigned_role ||
    chosen.role ||
    ""
  ).trim().toUpperCase();

  const assignedAtIso = String(rec.assigned_at || new Date().toISOString()).trim();
  const tarikhLantikan = _formatMalayLongDateUpper_(assignedAtIso);
  const refNo = _buildAppointmentLetterRefNo_(rec, chosen);

  const docName = "SURAT_LANTIKAN_" + (idMaklumBalas || complaintId || ("ADUAN_" + Date.now()));
  const copiedFile = DriveApp.getFileById(templateId).makeCopy(docName, targetFolder);
  const docId = copiedFile.getId();

  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();

  // Padankan ikut placeholder template surat lantikan terbaru
  _replaceDocPlaceholder_(body, "<<no_rujukan_fail>>", refNo);
  _replaceDocPlaceholder_(body, "<<tarikh_lantikan>>", tarikhLantikan);
  _replaceDocPlaceholder_(body, "<<Nama>>", pegawaiNama);
  _replaceDocPlaceholder_(body, "<<Jawatan>>", pegawaiJawatan);
  _replaceDocPlaceholder_(body, "<<Id_maklumbalas>>", idMaklumBalas);
  _replaceDocPlaceholder_(body, "<<tajuk>>", tajukAduan);
  _replaceDocPlaceholder_(body, "<<tarikh_terima>>", tarikhTerima);
  _replaceDocPlaceholder_(body, "<<due_date>>", dueDate);

  doc.saveAndClose();

  const pdfBlob = DriveApp.getFileById(docId).getAs(MimeType.PDF).setName(docName + ".pdf");
  const pdfFile = targetFolder.createFile(pdfBlob);
  const pdfFileId = pdfFile.getId();

  try {
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    Logger.log("PDF sharing failed: " + (e && e.message ? e.message : e));
  }

  const pdfUrl = "https://drive.google.com/uc?export=view&id=" + pdfFileId;

  return {
    ref_no: refNo,
    assigned_at_iso: assignedAtIso,
    due_date_iso: String(rec.due_date || "").trim(),
    pdf_file_id: pdfFileId,
    pdf_url: pdfUrl,
    doc_id: docId
  };
}

function getAppointmentLetterUrl(token, complaintId) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };

  const found = _findComplaintById_(complaintId);
  if (!found || !found.record) {
    return { ok: false, message: "Rekod aduan tidak dijumpai." };
  }

  const rec = found.record || {};
  const fileId = String(rec.appointment_letter_pdf_file_id || "").trim();

  if (!fileId) {
    return { ok: false, message: "Fail surat lantikan belum dijana." };
  }

  return {
    ok: true,
    url: "https://drive.google.com/file/d/" + fileId + "/view"
  };
}