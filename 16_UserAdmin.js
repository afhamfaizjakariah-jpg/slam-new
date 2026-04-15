/******************************
 * 16_UserAdmin.gs
 ******************************/

const USER_ADMIN_ALLOWED_ROLES = ["ADMIN", "PENTADBIR SISTEM"];

function _userAdminRequire_(session) {
  _requireRole_(session, USER_ADMIN_ALLOWED_ROLES);
}

function _userAdminValue_(v) {
  return String(v || "").trim();
}

function _getUserSheetContext_() {
  const meta = _getUserSheetMeta_();
  const sh = meta.sheet;
  const headers = meta.headers || [];
  const values = meta.values || [];
  const map = _getUserProfileColumnMap_(headers);
  return { sheet: sh, headers: headers, values: values, map: map };
}

function _findUserRowById_(ctx, userId) {
  const key = _userAdminValue_(userId);
  if (!ctx || !key || !ctx.values || ctx.map.userId < 0) return null;

  for (let r = 1; r < ctx.values.length; r++) {
    const row = ctx.values[r];
    if (_userAdminValue_(row[ctx.map.userId]) === key) {
      return {
        rowNumber: r + 1,
        rowValues: row
      };
    }
  }
  return null;
}

function _serializeSystemUser_(ctx, row) {
  const map = ctx.map;
  const signatureFileId = map.tandatangan >= 0 ? _userAdminValue_(row[map.tandatangan]) : "";
  return {
    userId: map.userId >= 0 ? _userAdminValue_(row[map.userId]) : "",
    full_name: map.name >= 0 ? _userAdminValue_(row[map.name]) : "",
    name: map.name >= 0 ? _userAdminValue_(row[map.name]) : "",
    role: map.role >= 0 ? _userAdminValue_(row[map.role]) : "",
    jawatan: map.jawatan >= 0 ? _userAdminValue_(row[map.jawatan]) : "",
    phone: map.phone >= 0 ? _userAdminValue_(row[map.phone]) : "",
    signature_file_id: signatureFileId,
    has_signature: !!signatureFileId
  };
}

function listSystemUsers(token) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };
  _userAdminRequire_(s);

  const ctx = _getUserSheetContext_();
  const users = [];

  for (let r = 1; r < ctx.values.length; r++) {
    const row = ctx.values[r];
    const userId = ctx.map.userId >= 0 ? _userAdminValue_(row[ctx.map.userId]) : "";
    if (!userId) continue;
    users.push(_serializeSystemUser_(ctx, row));
  }

  users.sort(function(a, b) {
    return String(a.full_name || a.userId).localeCompare(String(b.full_name || b.userId));
  });

  return { ok: true, users: users };
}

function saveSystemUser(token, payload) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };
  _userAdminRequire_(s);

  payload = payload || {};
  const userId = _userAdminValue_(payload.userId);
  const fullName = _userAdminValue_(payload.full_name || payload.name);
  const role = _userAdminValue_(payload.role).toUpperCase();
  const jawatan = _userAdminValue_(payload.jawatan);
  const phone = _userAdminValue_(payload.phone);
  const password = _userAdminValue_(payload.password);

  if (!userId) return { ok: false, message: "ID Pengguna wajib diisi." };
  if (!fullName) return { ok: false, message: "Nama Penuh wajib diisi." };
  if (!role) return { ok: false, message: "Peranan wajib dipilih." };

  const ctx = _getUserSheetContext_();
  const sh = ctx.sheet;
  const map = ctx.map;

  if (map.userId < 0 || map.password < 0 || map.name < 0 || map.role < 0) {
    return { ok: false, message: "Struktur sheet Pengguna tidak lengkap untuk modul pengurusan pengguna." };
  }

  const existing = _findUserRowById_(ctx, userId);

  if (!existing) {
    if (!password) return { ok: false, message: "Kata laluan wajib diisi untuk pengguna baru." };

    const row = new Array(ctx.headers.length).fill("");
    row[map.userId] = userId;
    row[map.password] = password;
    row[map.name] = fullName;
    row[map.role] = role;
    if (map.jawatan >= 0) row[map.jawatan] = jawatan;
    if (map.phone >= 0) row[map.phone] = phone;

    sh.appendRow(row);
    SpreadsheetApp.flush();

    return { ok: true, message: "Pengguna baru berjaya didaftarkan." };
  }

  const rowNo = existing.rowNumber;
  sh.getRange(rowNo, map.name + 1).setValue(fullName);
  sh.getRange(rowNo, map.role + 1).setValue(role);
  if (map.jawatan >= 0) sh.getRange(rowNo, map.jawatan + 1).setValue(jawatan);
  if (map.phone >= 0) sh.getRange(rowNo, map.phone + 1).setValue(phone);
  if (password) sh.getRange(rowNo, map.password + 1).setValue(password);

  SpreadsheetApp.flush();
  return { ok: true, message: "Rekod pengguna berjaya dikemas kini." };
}

function resetSystemUserPassword(token, userId, newPassword) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };
  _userAdminRequire_(s);

  userId = _userAdminValue_(userId);
  newPassword = _userAdminValue_(newPassword);

  if (!userId) return { ok: false, message: "ID pengguna tidak sah." };
  if (!newPassword) return { ok: false, message: "Kata laluan baharu wajib diisi." };
  if (newPassword.length < 6) return { ok: false, message: "Kata laluan baharu mesti sekurang-kurangnya 6 aksara." };

  const ctx = _getUserSheetContext_();
  if (ctx.map.password < 0) return { ok: false, message: "Kolum Kata Laluan tidak ditemui." };

  const existing = _findUserRowById_(ctx, userId);
  if (!existing) return { ok: false, message: "Pengguna tidak dijumpai." };

  ctx.sheet.getRange(existing.rowNumber, ctx.map.password + 1).setValue(newPassword);
  SpreadsheetApp.flush();

  return { ok: true, message: "Kata laluan pengguna berjaya dikemas kini." };
}

function deleteSystemUser(token, userId) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };
  _userAdminRequire_(s);

  userId = _userAdminValue_(userId);
  if (!userId) return { ok: false, message: "ID pengguna tidak sah." };
  if (userId === _userAdminValue_(s.userId)) {
    return { ok: false, message: "Anda tidak boleh memadam akaun anda sendiri." };
  }

  const ctx = _getUserSheetContext_();
  const existing = _findUserRowById_(ctx, userId);
  if (!existing) return { ok: false, message: "Pengguna tidak dijumpai." };

  const row = existing.rowValues || [];
  if (ctx.map.tandatangan >= 0) {
    const sigFileId = _userAdminValue_(row[ctx.map.tandatangan]);
    if (sigFileId) {
      try { DriveApp.getFileById(sigFileId).setTrashed(true); } catch (e) {}
    }
  }

  ctx.sheet.deleteRow(existing.rowNumber);
  SpreadsheetApp.flush();

  return { ok: true, message: "Pengguna berjaya dipadam." };
}
