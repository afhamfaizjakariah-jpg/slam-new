/******************************
 * 02_Profile.gs
 ******************************/

/** ========= TETAPAN / PROFIL ========= **/
function getMyProfile(token) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };

  const prof = _findUserFullProfile_(s.userId);
  if (!prof) return { ok: false, message: "Profil pengguna tidak dijumpai." };

  const sigRaw = String(prof.signature_file_id || "").trim();
  const sigUrl = sigRaw
    ? (/^https?:\/\//i.test(sigRaw) ? sigRaw : ("https://drive.google.com/thumbnail?id=" + sigRaw + "&sz=w1000"))
    : "";

  return {
    ok: true,
    profile: {
      userId: prof.userId || "",
      full_name: prof.full_name || "",
      role: prof.role || "",
      jawatan: prof.jawatan || "",
      pejabat: "",
      phone: prof.phone || "",
      signature_file_id: sigRaw,
      signature_url: sigUrl
    }
  };
}

function saveMyProfile(token, payload) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };

  payload = payload || {};
  const fullName = String(payload.full_name || "").trim();
  const jawatan = String(payload.jawatan || "").trim();

  if (!fullName) return { ok: false, message: "Nama penuh wajib diisi." };
  if (!jawatan) return { ok: false, message: "Jawatan wajib diisi." };

  const updated = _updateUserProfile_(s.userId, {
    full_name: fullName,
    jawatan: jawatan
  });
  if (!updated.ok) return updated;

  const session = {
    token: s.token,
    userId: s.userId,
    name: fullName || s.name,
    role: s.role || "PENGGUNA",
    ts: Date.now()
  };
  CacheService.getScriptCache().put(_sessKey_(s.token), JSON.stringify(session), CONFIG.SESSION_TTL_SECONDS);

  return {
    ok: true,
    message: "Profil berjaya dikemas kini."
  };
}

function changeMyPassword(token, payload) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };

  payload = payload || {};
  const currentPassword = String(payload.current_password || "").trim();
  const newPassword = String(payload.new_password || "").trim();
  const confirmPassword = String(payload.confirm_password || "").trim();

  if (!currentPassword) return { ok: false, message: "Kata laluan semasa wajib diisi." };
  if (!newPassword) return { ok: false, message: "Kata laluan baharu wajib diisi." };
  if (!confirmPassword) return { ok: false, message: "Pengesahan kata laluan baharu wajib diisi." };
  if (newPassword.length < 6) return { ok: false, message: "Kata laluan baharu mesti sekurang-kurangnya 6 aksara." };
  if (newPassword !== confirmPassword) return { ok: false, message: "Pengesahan kata laluan baharu tidak sepadan." };

  const prof = _findUserFullProfile_(s.userId);
  if (!prof) return { ok: false, message: "Profil pengguna tidak dijumpai." };

  const storedPw = String(prof.password || "").trim();
  if (storedPw !== currentPassword) {
    return { ok: false, message: "Kata laluan semasa tidak tepat." };
  }
  if (storedPw === newPassword) {
    return { ok: false, message: "Kata laluan baharu mesti berbeza daripada kata laluan semasa." };
  }

  const updated = _updateUserProfile_(s.userId, { password: newPassword });
  if (!updated.ok) return updated;

  return {
    ok: true,
    message: "Kata laluan berjaya dikemas kini."
  };
}

function uploadMySignature(token, base64Image, filename, mimeType) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };

  base64Image = String(base64Image || "").trim();
  filename = String(filename || "").trim() || ("signature_" + s.userId + ".png");
  mimeType = String(mimeType || "").trim() || "image/png";

  if (!base64Image) return { ok: false, message: "Fail tandatangan tidak diterima." };
  if (!/^image\/(png|jpg|jpeg)$/i.test(mimeType)) {
    return { ok: false, message: "Format tandatangan tidak sah. Sila gunakan PNG/JPG/JPEG." };
  }

  let bytes;
  try {
    bytes = Utilities.base64Decode(base64Image);
  } catch (e) {
    return { ok: false, message: "Format fail tandatangan tidak sah." };
  }

  const folder = _getOrCreateSignatureFolder_();
  const oldProf = _findUserFullProfile_(s.userId);

  try {
    if (oldProf && oldProf.signature_file_id) {
      try { DriveApp.getFileById(oldProf.signature_file_id).setTrashed(true); } catch (e) {}
    }

    const ext = mimeType.toLowerCase().indexOf("jpeg") >= 0 ? "jpg"
      : mimeType.toLowerCase().indexOf("jpg") >= 0 ? "jpg"
      : "png";

    const cleanName = "SLAM_SIGNATURE_" + s.userId + "_" + Date.now() + "." + ext;
    const blob = Utilities.newBlob(bytes, mimeType, cleanName);
    const file = folder.createFile(blob);

    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {}

    const fileId = file.getId();
    const url = "https://drive.google.com/uc?export=view&id=" + fileId;

    const updated = _updateUserProfile_(s.userId, {
      signature_file_id: fileId
    });
    if (!updated.ok) return updated;

    return {
      ok: true,
      message: "Tandatangan berjaya dimuat naik.",
      signature_file_id: fileId,
      signature_url: url
    };
  } catch (e) {
    return {
      ok: false,
      message: "Gagal memuat naik tandatangan: " + (e && e.message ? e.message : e)
    };
  }
}

function _getUserSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  for (const name of CONFIG.USER_SHEET_CANDIDATES) {
    const sh = ss.getSheetByName(name);
    if (sh) return sh;
  }
  return ss.getSheets()[0];
}

function _findUser(userId) {
  const sh = _getUserSheet_();
  const values = sh.getDataRange().getValues();
  if (!values || values.length < 1) return null;

  const header = values[0].map(v => String(v || "").trim().toLowerCase());
  const hasHeader = header.some(h => ["id", "id pengguna", "id_pengguna", "user id", "userid"].includes(h));

  let idxId = 0, idxPw = 1, idxName = 2, idxRole = 3;
  if (hasHeader) {
    const findAny = (needles) => {
      for (const n of needles) {
        const i = header.indexOf(String(n).toLowerCase());
        if (i >= 0) return i;
      }
      return -1;
    };
    idxId = findAny(["id", "id pengguna", "id_pengguna", "userid", "user id"]); if (idxId < 0) idxId = 0;
    idxPw = findAny(["kata laluan", "katalaluan", "password", "pass"]); if (idxPw < 0) idxPw = 1;
    idxName = findAny(["nama", "name"]); if (idxName < 0) idxName = 2;
    idxRole = findAny(["tugas", "role", "jenis tugas", "jenis", "akses"]); if (idxRole < 0) idxRole = 3;
  }

  const startRow = hasHeader ? 1 : 0;
  for (let r = startRow; r < values.length; r++) {
    const row = values[r];
    const rid = String(row[idxId] || "").trim();
    if (!rid) continue;
    if (rid === userId) {
      return {
        userId: rid,
        password: String(row[idxPw] || "").trim(),
        name: String(row[idxName] || "").trim(),
        role: String(row[idxRole] || "").trim()
      };
    }
  }
  return null;
}

function _normalizeHeaderName_(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/\u00A0/g, " ")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/[_\-]+/g, " ")
    .replace(/\./g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function _findHeaderIndex_(headers, aliases) {
  const normalizedHeaders = (headers || []).map(h => _normalizeHeaderName_(h));
  for (const alias of aliases || []) {
    const idx = normalizedHeaders.indexOf(_normalizeHeaderName_(alias));
    if (idx >= 0) return idx;
  }
  return -1;
}

function _getUserSheetMeta_() {
  const sh = _getUserSheet_();
  const values = sh.getDataRange().getValues();
  const headers = values && values.length ? values[0].map(v => String(v || "").trim()) : [];
  return { sheet: sh, headers: headers, values: values || [] };
}

function _getUserProfileColumnMap_(headers) {
  return {
    userId: _findHeaderIndex_(headers, ["ID", "id", "id pengguna", "userid", "user id"]),
    password: _findHeaderIndex_(headers, ["Kata Laluan", "kata laluan", "password", "pass"]),
    name: _findHeaderIndex_(headers, ["Nama", "nama", "name"]),
    jawatan: _findHeaderIndex_(headers, ["Jawatan", "jawatan", "position"]),
    role: _findHeaderIndex_(headers, ["Tugas", "tugas", "role", "akses"]),
    tandatangan: _findHeaderIndex_(headers, ["Tandatangan", "tandatangan", "signature", "signature url", "signature file id"]),
    phone: _findHeaderIndex_(headers, ["No. Telefon", "No Telefon", "Telefon", "no telefon", "no. telefon", "phone"])
  };
}

function _findUserFullProfile_(userId) {
  const meta = _getUserSheetMeta_();
  const sh = meta.sheet;
  const headers = meta.headers;
  const values = meta.values;
  const map = _getUserProfileColumnMap_(headers);

  if (!values || values.length < 2) return null;
  if (map.userId < 0) return null;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const rid = String(row[map.userId] || "").trim();
    if (!rid) continue;

    if (rid === String(userId || "").trim()) {
      return {
        sheet: sh,
        row: r + 1,
        userId: rid,
        password: map.password >= 0 ? String(row[map.password] || "").trim() : "",
        name: map.name >= 0 ? String(row[map.name] || "").trim() : "",
        full_name: map.name >= 0 ? String(row[map.name] || "").trim() : "",
        role: map.role >= 0 ? String(row[map.role] || "").trim() : "",
        jawatan: map.jawatan >= 0 ? String(row[map.jawatan] || "").trim() : "",
        pejabat: "",
        signature_file_id: map.tandatangan >= 0 ? String(row[map.tandatangan] || "").trim() : "",
        signature_url: map.tandatangan >= 0 ? String(row[map.tandatangan] || "").trim() : "",
        phone: map.phone >= 0 ? String(row[map.phone] || "").trim() : "",
        map: map,
        headers: headers
      };
    }
  }
  return null;
}

function _updateUserProfile_(userId, payload) {
  const prof = _findUserFullProfile_(userId);
  if (!prof) return { ok: false, message: "Pengguna tidak dijumpai." };

  const sh = prof.sheet;
  const row = prof.row;
  const map = prof.map;

  if (payload.full_name !== undefined) {
    const v = String(payload.full_name || "").trim();
    if (map.name < 0) return { ok: false, message: "Kolum Nama tidak ditemui dalam sheet Pengguna." };
    sh.getRange(row, map.name + 1).setValue(v);
  }

  if (payload.jawatan !== undefined) {
    const v = String(payload.jawatan || "").trim();
    if (map.jawatan < 0) return { ok: false, message: "Kolum Jawatan tidak ditemui dalam sheet Pengguna." };
    sh.getRange(row, map.jawatan + 1).setValue(v);
  }

  if (payload.password !== undefined) {
    const v = String(payload.password || "").trim();
    if (map.password < 0) return { ok: false, message: "Kolum Kata Laluan tidak ditemui dalam sheet Pengguna." };
    sh.getRange(row, map.password + 1).setValue(v);
  }

  if (payload.signature_file_id !== undefined) {
    const v = String(payload.signature_file_id || "").trim();
    if (map.tandatangan < 0) return { ok: false, message: "Kolum Tandatangan tidak ditemui dalam sheet Pengguna." };
    sh.getRange(row, map.tandatangan + 1).setValue(v);
  }

  if (payload.signature_url !== undefined) {
    const v = String(payload.signature_url || "").trim();
    if (map.tandatangan < 0) return { ok: false, message: "Kolum Tandatangan tidak ditemui dalam sheet Pengguna." };
    sh.getRange(row, map.tandatangan + 1).setValue(v);
  }

  return { ok: true };
}

function _getOrCreateSignatureFolder_() {
  const name = String(CONFIG.PROFILE_SIGNATURE_FOLDER_NAME || "SLAM_SIGNATURES").trim();
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}