/******************************
 * 14_ReportGenerate.gs
 ******************************/

function saveAndGenerateReport(token, complaintId, payload) {
  try {
    payload = payload || {};

    const s = _getSession_(token);
    if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };

    const found = _findComplaintById_(complaintId);
    if (!found || !found.record) {
      return { ok: false, message: "Rekod aduan tidak dijumpai." };
    }

    const rec = found.record || {};
    const investigatorUserId = String(rec.assigned_user_id || s.userId || "").trim();
    const prof = investigatorUserId ? _findUserFullProfile_(investigatorUserId) : null;

    payload.pegawai_penyiasat_user_id = String(payload.pegawai_penyiasat_user_id || investigatorUserId || "").trim();
    payload.nama = String(payload.nama || (prof && (prof.full_name || prof.name)) || rec.assigned_to || s.name || "").trim();
    payload.jawatan = String(payload.jawatan || (prof && prof.jawatan) || rec.assigned_role || "").trim();

    const saved = saveReportDraft(token, complaintId, payload);
    if (!saved || !saved.ok) return saved;

    const generated = _generateReportPdfFromComplaint_(complaintId, s);
    if (!generated || !generated.ok) return generated;

    return {
      ok: true,
      message: "Laporan berjaya disimpan dan PDF telah dijana mengikut template.",
      pdf_url: generated.pdf_url,
      pdf_file_id: generated.pdf_file_id,
      doc_id: generated.doc_id
    };
  } catch (e) {
    return { ok: false, message: "saveAndGenerateReport error: " + (e && e.message ? e.message : e) };
  }
}

function _getOrCreateReportOutputFolder_() {
  const name = String((CONFIG && CONFIG.REPORT_OUTPUT_FOLDER_NAME) || "SLAM_REPORT_OUTPUT").trim();
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

function _reportValue_(v) {
  return String(v || "").trim();
}

function _reportSplitLines_(v) {
  return _reportValue_(v)
    .split(/\r?\n|;/)
    .map(function(x) { return String(x || "").trim(); })
    .filter(Boolean);
}

function _reportEscape_(text) {
  return String(text || "").replace(/\$/g, "$$$$");
}

function _replaceAllDocText_(body, placeholder, value) {
  const safePattern = String(placeholder || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  body.replaceText(safePattern, _reportEscape_(value));
}

function _buildCheckboxLines_(selectedRaw) {
  const all = [
    "Faktor Individu (Anggota KKM)",
    "Faktor Pengurusan dan Organisasi",
    "Faktor Sistem dan SOP",
    "Faktor Kemudahan Fizikal",
    "Faktor Luar",
    "Faktor Lain-lain"
  ];
  const selected = _reportSplitLines_(selectedRaw);
  return all.map(function(item) {
    return (selected.indexOf(item) >= 0 ? "☑ " : "☐ ") + item;
  });
}

function _replaceParagraphTextExact_(body, exactText, replacement) {
  const found = body.findText(String(exactText || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
  if (!found) return false;
  const element = found.getElement();
  if (!element) return false;
  const para = element.getParent().getType() === DocumentApp.ElementType.PARAGRAPH
    ? element.getParent().asParagraph()
    : element.getParent().getParent().asParagraph();
  const container = para.getParent();
  const idx = container.getChildIndex(para);
  para.removeFromParent();
  const newPara = container.insertParagraph(idx, String(replacement || ""));
  newPara.editAsText().setBold(false);
  return true;
}



function _replacePlaceholderOccurrences_(body, placeholder, values) {
  const arr = Array.isArray(values) ? values.slice() : [values];
  const pattern = String(placeholder || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  let i = 0;
  while (true) {
    const found = body.findText(pattern);
    if (!found) break;
    const el = found.getElement();
    const txt = el.asText();
    const replacement = String(arr[Math.min(i, arr.length - 1)] || "-");
    txt.setText(txt.getText().replace(String(placeholder), replacement));
    i++;
  }
}

function _insertPageBreakBeforeText_(body, exactText) {
  const found = body.findText(String(exactText || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
  if (!found) return false;
  const el = found.getElement();
  const para = el.getParent().getType() === DocumentApp.ElementType.PARAGRAPH ? el.getParent() : el.getParent().getParent();
  const parent = para.getParent();
  if (parent.getType && parent.getType() !== DocumentApp.ElementType.BODY_SECTION) return false;
  const idx = parent.getChildIndex(para);
  if (idx > 0) parent.insertPageBreak(idx);
  return true;
}

function _insertImageBlockAtPlaceholder_(body, placeholder, items, maxWidth) {
  const found = body.findText(String(placeholder || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
  if (!found) return false;

  const el = found.getElement();
  const para = el.getParent().getType() === DocumentApp.ElementType.PARAGRAPH
    ? el.getParent().asParagraph()
    : el.getParent().getParent().asParagraph();

  try { el.asText().setText(el.asText().getText().replace(String(placeholder), "")); } catch (e) {}

  const parent = para.getParent();
  const index = parent.getChildIndex(para);
  let insertAt = index;

  (items || []).forEach(function(item) {
    const fileId = _reportValue_(item.file_id);
    if (!fileId) return;

    try {
      const imgPara = parent.insertParagraph(insertAt, "");
      imgPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      const img = imgPara.appendInlineImage(DriveApp.getFileById(fileId).getBlob());
      const w = img.getWidth();
      const h = img.getHeight();
      if (w > maxWidth) {
        const ratio = maxWidth / w;
        img.setWidth(Math.round(w * ratio));
        img.setHeight(Math.round(h * ratio));
      }
      insertAt += 1;
    } catch (e) {
      parent.insertParagraph(insertAt++, "[Imej gagal dimuatkan]");
    }
  });

  return true;
}

function _insertSignatureAtPlaceholder_(body, placeholder, signatureFileId) {
  const fileId = _reportValue_(signatureFileId);
  if (!fileId) {
    _replaceAllDocText_(body, placeholder, "");
    return false;
  }

  const found = body.findText(String(placeholder || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
  if (!found) return false;

  const el = found.getElement();
  const para = el.getParent().getType() === DocumentApp.ElementType.PARAGRAPH
    ? el.getParent().asParagraph()
    : el.getParent().getParent().asParagraph();

  try { el.asText().setText(el.asText().getText().replace(String(placeholder), "")); } catch (e) {}

  try {
    const img = para.appendInlineImage(DriveApp.getFileById(fileId).getBlob());
    const w = img.getWidth();
    const h = img.getHeight();
    const maxWidth = 120;
    if (w > maxWidth) {
      const ratio = maxWidth / w;
      img.setWidth(Math.round(w * ratio));
      img.setHeight(Math.round(h * ratio));
    }
    return true;
  } catch (e) {
    return false;
  }
}


function _findingsTableTarget_(body) {
  const found = body.findText("<<penemuan>>") || body.findText("<<isu>>");
  if (!found) return null;

  var node = found.getElement();
  var row = null;
  var table = null;
  while (node) {
    if (!row && node.getType && node.getType() === DocumentApp.ElementType.TABLE_ROW) row = node.asTableRow();
    if (!table && node.getType && node.getType() === DocumentApp.ElementType.TABLE) {
      table = node.asTable();
      break;
    }
    node = node.getParent ? node.getParent() : null;
  }
  if (!row || !table) return null;

  var rowIndex = table.getChildIndex(row);
  if (rowIndex < 0) return null;

  return { table: table, row: row, rowIndex: rowIndex };
}

function _fillFindingsTable_(body, findings) {
  const target = _findingsTableTarget_(body);
  if (!target) return false;

  const rows = (findings || []).length ? findings : [{ isu: "-", penemuan: "-" }];
  const table = target.table;
  const baseRow = target.row;
  const rowIndex = target.rowIndex;

  baseRow.getCell(0).clear();
  baseRow.getCell(0).appendParagraph(String(rows[0].isu || "-"));
  baseRow.getCell(1).clear();
  baseRow.getCell(1).appendParagraph(String(rows[0].penemuan || "-"));

  for (var j = 1; j < rows.length; j++) {
    var newRow = table.insertTableRow(rowIndex + j);
    newRow.appendTableCell(String(rows[j].isu || "-"));
    newRow.appendTableCell(String(rows[j].penemuan || "-"));
  }
  return true;
}

function _generateReportPdfFromComplaint_(complaintId, session) {
  try {
  const found = _findComplaintById_(complaintId);
  if (!found || !found.record) {
    return { ok: false, message: "Rekod aduan tidak dijumpai untuk penjanaan laporan." };
  }

  const rec = found.record || {};
  if (!_canEditReport_(session, rec) && !_isPentadbirSistemLike_(session)) {
    return { ok: false, message: "Anda tidak mempunyai akses untuk menjana laporan ini." };
  }

  const templateId = String((CONFIG && CONFIG.REPORT_DOC_TEMPLATE_ID) || "").trim();
  if (!templateId) {
    return { ok: false, message: "CONFIG.REPORT_DOC_TEMPLATE_ID belum ditetapkan." };
  }

  const investigatorUserId = _reportValue_(rec.report_pegawai_penyiasat_user_id || rec.assigned_user_id || session.userId);
  const investigatorProfile = investigatorUserId ? _findUserFullProfile_(investigatorUserId) : null;
  const draft = _buildReportDraft_(rec, investigatorProfile);

  const folder = _getOrCreateReportOutputFolder_();
  const reportId = _reportValue_(draft.id_maklumbalas || rec.id_maklumbalas || rec.complaint_id || ("REPORT_" + Date.now()));
  const safeId = reportId.replace(/[^A-Za-z0-9._-]+/g, "_");
  const docName = "LAPORAN_ADUAN_" + safeId;

  try {
    const oldPdfId = _reportValue_(rec.report_pdf_file_id);
    const oldDocId = _reportValue_(rec.report_doc_id);
    if (oldPdfId) { try { DriveApp.getFileById(oldPdfId).setTrashed(true); } catch (e) {} }
    if (oldDocId) { try { DriveApp.getFileById(oldDocId).setTrashed(true); } catch (e) {} }
  } catch (e) {}

  const copiedFile = DriveApp.getFileById(templateId).makeCopy(docName, folder);
  const docId = copiedFile.getId();
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();

  const findings = Array.isArray(draft.findings) ? draft.findings : [];
  const isuRowsText = findings.map(function(item) {
    return _reportValue_(item.isu || "-");
  }).join("\n");
  const isuCommaText = findings.map(function(item) {
    return _reportValue_(item.isu || "-");
  }).filter(Boolean).join(", ");
  const penemuanText = findings.map(function(item) {
    return _reportValue_(item.penemuan || "-");
  }).join("\n");
  const kelemahanText = (draft.kelemahan_list || []).map(function(item) {
    return _reportValue_(item || "-");
  }).join("\n");

  _replaceAllDocText_(body, "<<id_maklumbalas>>", draft.id_maklumbalas || "-");
    _replaceAllDocText_(body, "<<tajuk>>", draft.tajuk || "-");
  _replaceAllDocText_(body, "<<tarikh_terima>>", draft.tarikh_terima || "-");
  _replaceAllDocText_(body, "<<ringkasan_butiran>>", draft.ringkasan_butiran || "-");
  _replaceAllDocText_(body, "<<lokasi>>", draft.lokasi || "-");
  _replaceAllDocText_(body, "<<nama_pengadu>>", draft.nama_pengadu || "-");
  _replaceAllDocText_(body, "<<tarikh_siasatan>>", draft.tarikh_siasatan || "-");
  _replaceAllDocText_(body, "<<nama>>", draft.nama || "-");
  _replaceAllDocText_(body, "<<jawatan>>", draft.jawatan || "-");
  _fillFindingsTable_(body, findings);
  _replaceAllDocText_(body, "<<isu>>", isuCommaText || isuRowsText || "-");
  _replaceAllDocText_(body, "<isu>>", isuCommaText || isuRowsText || "-");
  _replaceAllDocText_(body, "<<penemuan>>", findings.map(function(item) { return item.penemuan || "-"; }).join("\n") || "-");
  _replaceAllDocText_(body, "<<rumusan>>", draft.rumusan || "-");
  _replaceAllDocText_(body, "<<kelemahan>>", kelemahanText || "-");
  _replaceAllDocText_(body, "<<tindakan_penguatkuasaan>>", draft.tindakan_penguatkuasaan || "-");

  _buildCheckboxLines_(draft.kategori_penyelesaian_aduan).forEach(function(line) {
    const label = line.replace(/^☑\s|^☐\s/, "");
    _replaceParagraphTextExact_(body, label, line);
  });

  _insertSignatureAtPlaceholder_(body, "<<tandatangan>>", draft.tandatangan_file_id || "");

  _insertPageBreakBeforeText_(body, "ULASAN KETUA JABATAN / PENGERUSI J/K SIASATAN :");
  _insertPageBreakBeforeText_(body, "Lampiran");

  const frontImageItems = [];
  if (_reportValue_(draft.gambar_hadapan_premis_file_id)) {
    frontImageItems.push({ file_id: draft.gambar_hadapan_premis_file_id });
  }
  _insertImageBlockAtPlaceholder_(body, "<<gambar_hadapan_premis>>", frontImageItems, 360);

  const weaknessItems = (draft.weakness_images || []).map(function(item) {
    return {
      file_id: _reportValue_(item.file_id),
      keterangan: _reportValue_(item.keterangan || "Lampiran")
    };
  }).filter(function(item) { return item.file_id; });

  _insertImageBlockAtPlaceholder_(body, "<<gambar_kelemahan>>", weaknessItems, 360);
  _replaceAllDocText_(body, "<<keterangan_gambar>>", weaknessItems.map(function(x) { return x.keterangan; }).join("\n"));

  doc.saveAndClose();

  const pdfBlob = DriveApp.getFileById(docId).getAs(MimeType.PDF).setName(docName + ".pdf");
  const pdfFile = folder.createFile(pdfBlob);
  try { pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}

  const pdfFileId = pdfFile.getId();
  const pdfUrl = "https://drive.google.com/file/d/" + pdfFileId + "/view";
  const nowIso = new Date().toISOString();
  const sh = found.sheet;
  const row = found.row;
  const map = _headerMap_(found.headers);

  _setCellByHeader_(sh, row, map, "report_generated_at", nowIso);
  _setCellByHeader_(sh, row, map, "report_doc_id", docId);
  _setCellByHeader_(sh, row, map, "report_pdf_file_id", pdfFileId);
  _setCellByHeader_(sh, row, map, "report_pdf_url", pdfUrl);
  _setCellByHeader_(sh, row, map, "report_status", "DIJANA");
  _setCellByHeader_(sh, row, map, "report_submitted_at", nowIso);
  _setCellByHeader_(sh, row, map, "card_status", "Selesai");
  _setCellByHeader_(sh, row, map, "status", "Selesai");
  _setCellByHeader_(sh, row, map, "status_updated_at", nowIso);
  SpreadsheetApp.flush();

  return {
    ok: true,
    doc_id: docId,
    pdf_file_id: pdfFileId,
    pdf_url: pdfUrl
  };
  } catch (e) {
    return { ok: false, message: "generateReport error: " + (e && e.message ? e.message : e) };
  }
}

function getGeneratedReportUrl(token, complaintId) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };

  const found = _findComplaintById_(complaintId);
  if (!found || !found.record) return { ok: false, message: "Rekod aduan tidak dijumpai." };

  const rec = found.record || {};
  const url = _reportValue_(rec.report_pdf_url);
  if (!url) return { ok: false, message: "Laporan belum dijana." };
  return { ok: true, url: url };
}
