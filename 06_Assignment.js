/******************************
 * 06_Assignment.gs
 ******************************/

function listAssignableUsers(token) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };
  _requireRole_(s, CONFIG.ROLE_ALLOW_ASSIGN);

  const sh = _getUserSheet_();
  const values = sh.getDataRange().getValues();
  if (!values || !values.length) return { ok: true, users: [] };

  const header = values[0].map(v => String(v || "").trim().toLowerCase());
  const hasHeader = header.some(h => ["id", "id pengguna", "id_pengguna", "user id", "userid"].includes(h));

  let idxId = 0, idxName = 2, idxRole = 3;
  if (hasHeader) {
    const findAny = (needles) => {
      for (const n of needles) {
        const i = header.indexOf(String(n).toLowerCase());
        if (i >= 0) return i;
      }
      return -1;
    };
    idxId = findAny(["id", "id pengguna", "id_pengguna", "userid", "user id"]); if (idxId < 0) idxId = 0;
    idxName = findAny(["nama", "name"]); if (idxName < 0) idxName = 2;
    idxRole = findAny(["tugas", "role", "jenis tugas", "jenis", "akses"]); if (idxRole < 0) idxRole = 3;
  }

  const startRow = hasHeader ? 1 : 0;
  const allowed = ["PENTADBIR SISTEM", "PEGAWAI PENYIASAT"];
  const users = [];

  for (let r = startRow; r < values.length; r++) {
    const row = values[r];
    const userId = String(row[idxId] || "").trim();
    const name = String(row[idxName] || "").trim();
    const role = String(row[idxRole] || "").trim().toUpperCase();
    if (!userId || !role) continue;
    if (!allowed.includes(role)) continue;

    users.push({
      userId: userId,
      name: name || userId,
      role: role
    });
  }

  users.sort((a, b) => {
    if (a.role !== b.role) return a.role.localeCompare(b.role);
    return a.name.localeCompare(b.name);
  });

  return { ok: true, users: users };
}

function assignComplaintInvestigator(token, complaintId, selectedUserId) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };
  _requireRole_(s, CONFIG.ROLE_ALLOW_ASSIGN);

  selectedUserId = String(selectedUserId || "").trim();
  if (!selectedUserId) return { ok: false, message: "Sila pilih pegawai penyiasat." };

  const list = listAssignableUsers(token);
  if (!list || !list.ok) return { ok: false, message: "Gagal mendapatkan senarai pegawai." };

  const chosen = (list.users || []).find(u => String(u.userId || "").trim() === selectedUserId);
  if (!chosen) return { ok: false, message: "Pegawai yang dipilih tidak sah." };

  const found = _findComplaintById_(complaintId);
  if (!found) return { ok: false, message: "Rekod aduan tidak dijumpai." };

  const rec = found.record || {};
  if (String(rec.assigned_to || "").trim()) {
    return { ok: false, message: "Aduan ini telah diagihkan sebelum ini." };
  }

  const sh = found.sheet;
  const row = found.row;
  const headers = found.headers;
  const map = _headerMap_(headers);

  const assignedAt = new Date();
  const assignedAtIso = assignedAt.toISOString();
  const dueDateIso = _computeAppointmentDueDate_(rec, assignedAt);

  const recForLetter = Object.assign({}, rec, {
    assigned_to: chosen.name,
    assigned_user_id: chosen.userId,
    assigned_role: chosen.role,
    assigned_at: assignedAtIso,
    due_date: dueDateIso
  });

  let letter;
  try {
    letter = _generateAppointmentLetterPdf_(recForLetter, chosen);
  } catch (e) {
    return {
      ok: false,
      message: "Pegawai tidak berjaya diagihkan kerana penjanaan surat lantikan gagal: " + (e && e.message ? e.message : e)
    };
  }

  _setCellByHeader_(sh, row, map, "assigned_to", chosen.name);
  _setCellByHeader_(sh, row, map, "assigned_user_id", chosen.userId);
  _setCellByHeader_(sh, row, map, "assigned_role", chosen.role);
  _setCellByHeader_(sh, row, map, "assigned_at", assignedAtIso);
  _setCellByHeader_(sh, row, map, "card_status", "Telah Diagihkan");
  _setCellByHeader_(sh, row, map, "status", "Dalam Siasatan");
  _setCellByHeader_(sh, row, map, "status_updated_at", assignedAtIso);
  _setCellByHeader_(sh, row, map, "due_date", dueDateIso);

  _setCellByHeader_(sh, row, map, "appointment_letter_ref_no", letter.ref_no);
  _setCellByHeader_(sh, row, map, "appointment_letter_generated_at", assignedAtIso);
  _setCellByHeader_(sh, row, map, "appointment_letter_pdf_file_id", letter.pdf_file_id);
  _setCellByHeader_(sh, row, map, "appointment_letter_pdf_url", letter.pdf_url);
  _setCellByHeader_(sh, row, map, "appointment_letter_doc_id", letter.doc_id);

  return {
    ok: true,
    message: "Pegawai penyiasat berjaya diagihkan dan surat lantikan PDF telah dijana.",
    assigned_to: chosen.name,
    assigned_user_id: chosen.userId,
    assigned_role: chosen.role,
    appointment_letter_pdf_url: letter.pdf_url
  };
}