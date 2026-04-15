/******************************
 * 13_ReportModule.gs
 ******************************/

function _isPentadbirSistemLike_(session) {
  const raw = String(session && session.role || "").trim().toUpperCase();
  const normalized = _normalizeRole_(raw);
  return raw === "PENTADBIR SISTEM" || normalized === "ADMIN";
}

function _canEditReport_(session, rec) {
  if (!session || !rec) return false;
  if (_isPentadbirSistemLike_(session)) return true;

  const role = _normalizeRole_(session.role);
  const myUserId = String(session.userId || "").trim();
  const assignedUserId = String(rec.assigned_user_id || "").trim();
  const assignedTo = String(rec.assigned_to || "").trim();

  if (role !== "PEGAWAI PENYIASAT") return false;
  if (!assignedTo) return false;
  if (!myUserId || !assignedUserId) return false;

  return myUserId === assignedUserId;
}

function _getOrCreateReportImageFolder_() {
  const name = String(CONFIG.REPORT_IMAGE_FOLDER_NAME || "SLAM_REPORT_IMAGES").trim();
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

function _reportText_(v) {
  return String(v || "").trim();
}

function _reportDateForUi_(val) {
  return _formatDateMalay_(val || "");
}

function _reportNormalizeDateOrBlank_(val) {
  const raw = _reportText_(val);
  if (!raw) return "";
  return _normalizeDate_(raw);
}

function _normalizeReportFindings_(findings) {
  const arr = Array.isArray(findings) ? findings : [];
  return arr
    .map(item => ({
      isu: _reportText_(item && item.isu),
      penemuan: _reportText_(item && item.penemuan)
    }))
    .filter(item => item.isu || item.penemuan);
}

function _readReportFindingsFromRecord_(rec) {
  const raw = _reportText_(rec.report_findings_json);
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      const cleaned = _normalizeReportFindings_(parsed);
      if (cleaned.length) return cleaned;
    } catch (e) {}
  }

  const legacyIsu = _reportText_(rec.report_isu);
  const legacyPenemuan = _reportText_(rec.report_penemuan);
  if (legacyIsu || legacyPenemuan) {
    return [{ isu: legacyIsu, penemuan: legacyPenemuan }];
  }

  return [{ isu: "", penemuan: "" }];
}

function _normalizeKelemahanList_(items) {
  const arr = Array.isArray(items) ? items : [];
  return arr.map(item => _reportText_(item)).filter(Boolean);
}

function _readKelemahanListFromRecord_(rec) {
  const raw = _reportText_(rec.report_kelemahan_list_json);
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      const cleaned = _normalizeKelemahanList_(parsed);
      if (cleaned.length) return cleaned;
    } catch (e) {}
  }

  const legacy = _reportText_(rec.report_kelemahan);
  if (!legacy) return [""];

  const split = legacy
    .split(/\r?\n/)
    .map(s => s.replace(/^\s*\d+\.\s*/, "").trim())
    .filter(Boolean);

  return split.length ? split : [legacy];
}

function _joinKelemahanForTemplate_(items) {
  const cleaned = _normalizeKelemahanList_(items);
  return cleaned.map((t, i) => (i + 1) + ". " + t).join("\n");
}

function _normalizeWeaknessImages_(items) {
  const arr = Array.isArray(items) ? items : [];
  return arr.map(item => ({
    url: _reportText_(item && item.url),
    file_id: _reportText_(item && item.file_id),
    keterangan: _reportText_(item && item.keterangan)
  })).filter(item => item.url || item.keterangan || item.file_id);
}

function _readWeaknessImagesFromRecord_(rec) {
  const raw = _reportText_(rec.report_weakness_images_json);

  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      const cleaned = _normalizeWeaknessImages_(parsed);
      if (cleaned.length) return cleaned;
    } catch (e) {}
  }

  const legacyUrl = _reportText_(rec.report_gambar_kelemahan);
  const legacyFileId = _reportText_(rec.report_gambar_kelemahan_file_id);
  const legacyKet = _reportText_(rec.report_keterangan_gambar);

  if (legacyUrl || legacyFileId || legacyKet) {
    return [{
      url: legacyUrl,
      file_id: legacyFileId,
      keterangan: legacyKet
    }];
  }

  return [{
    url: "",
    file_id: "",
    keterangan: ""
  }];
}

function _syncWeaknessLegacyColumns_(sh, row, map, items) {
  const first = Array.isArray(items) && items.length ? items[0] : { url: "", file_id: "", keterangan: "" };
  _setCellByHeader_(sh, row, map, "report_gambar_kelemahan", first.url || "");
  _setCellByHeader_(sh, row, map, "report_gambar_kelemahan_file_id", first.file_id || "");
  _setCellByHeader_(sh, row, map, "report_keterangan_gambar", first.keterangan || "");
}

function _buildReportDraft_(rec, investigatorProfile) {
  const p = investigatorProfile || {};

  const sigRaw = _reportText_(rec.report_tandatangan_file_id || p.signature_file_id || "");
  const sigUrl = _reportText_(
    rec.report_tandatangan_url ||
    (sigRaw ? ("https://drive.google.com/thumbnail?id=" + sigRaw + "&sz=w1200") : "") ||
    ""
  );

  return {
    complaint_id: _reportText_(rec.complaint_id),

    // Bahagian 1
    id_maklumbalas: _reportText_(rec.report_id_maklumbalas || rec.id_maklumbalas),
    jenis_aduan: _reportText_(rec.report_jenis_aduan || rec.jenis_maklumbalas || rec.source),
    jenis_maklumbalas_awam: _reportText_(rec.report_jenis_maklumbalas_awam),
    tajuk: _reportText_(rec.report_tajuk || rec.tajuk),
    tarikh_terima: _reportDateForUi_(rec.report_tarikh_terima || rec.tarikh_terima),
    ringkasan_butiran: _reportText_(rec.report_ringkasan_butiran || rec.ringkasan_butiran),
    lokasi: _reportText_(rec.report_lokasi || rec.lokasi),
    nama_pengadu: _reportText_(rec.report_nama_pengadu || rec.nama_pengadu),
    tarikh_siasatan: _reportDateForUi_(rec.report_tarikh_siasatan),
    pegawai_penyiasat_user_id: _reportText_(rec.report_pegawai_penyiasat_user_id || rec.assigned_user_id),
    nama: _reportText_(rec.report_nama_pegawai || p.full_name || p.name || rec.assigned_to || ""),
    jawatan: _reportText_(rec.report_jawatan_pegawai || p.jawatan || rec.assigned_role || ""),

    // Bahagian 2
    kategori: _reportText_(rec.report_kategori),
    parlimen: _reportText_(rec.report_parlimen),
    jenis_premis: _reportText_(rec.report_jenis_premis),
    subkategori_premis_makanan: _reportText_(rec.report_subkategori_premis_makanan),
    subkategori_produk_makanan: _reportText_(rec.report_subkategori_produk_makanan),
    status_pensijilan: _reportText_(rec.report_status_pensijilan),
    status_pemeriksaan_terdahulu: _reportText_(rec.report_status_pemeriksaan_terdahulu),
    markah_pemeriksaan_semasa: _reportText_(rec.report_markah_pemeriksaan_semasa),
    tindakan_penguatkuasaan: _reportText_(rec.report_tindakan_penguatkuasaan),

    // Bahagian 3
    findings: _readReportFindingsFromRecord_(rec),
    rumusan: _reportText_(rec.report_rumusan),
    kelemahan_list: _readKelemahanListFromRecord_(rec),

    // Bahagian 4
    kategori_penyelesaian_aduan: _reportText_(rec.report_kategori_penyelesaian_aduan),

    // Tandatangan / Lampiran
    tandatangan_url: sigUrl,
    tandatangan_file_id: sigRaw,
    gambar_hadapan_premis: _reportText_(rec.report_gambar_hadapan_premis),
    gambar_hadapan_premis_file_id: _reportText_(rec.report_gambar_hadapan_premis_file_id),
    weakness_images: _readWeaknessImagesFromRecord_(rec),

    report_status: _reportText_(rec.report_status)
  };
}

function getReportDraft(token, complaintId) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };

  const found = _findComplaintById_(complaintId);
  if (!found || !found.record) {
    return { ok: false, message: "Rekod aduan tidak dijumpai." };
  }

  const rec = found.record || {};
  if (!_reportText_(rec.assigned_to)) {
    return { ok: false, message: "Aduan ini belum diagihkan kepada pegawai penyiasat." };
  }

  if (!_canEditReport_(s, rec)) {
    return { ok: false, message: "Anda tidak mempunyai akses untuk mengemaskini laporan ini." };
  }

  const investigatorUserId = _reportText_(rec.assigned_user_id || s.userId);
  const investigatorProfile = investigatorUserId ? _findUserFullProfile_(investigatorUserId) : null;
  const draft = _buildReportDraft_(rec, investigatorProfile);

  return { ok: true, draft: draft };
}

function saveReportDraft(token, complaintId, payload) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };

  const found = _findComplaintById_(complaintId);
  if (!found || !found.record) {
    return { ok: false, message: "Rekod aduan tidak dijumpai." };
  }

  const rec = found.record || {};
  if (!_canEditReport_(s, rec)) {
    return { ok: false, message: "Anda tidak mempunyai akses untuk mengemaskini laporan ini." };
  }

  payload = payload || {};

  // Bahagian 1
  const idMaklumBalas = _reportText_(payload.id_maklumbalas || rec.id_maklumbalas);
  const jenisAduan = _reportText_(payload.jenis_aduan || rec.jenis_maklumbalas || rec.source);
  const jenisMaklumbalasAwam = _reportText_(payload.jenis_maklumbalas_awam);
  const tajuk = _reportText_(payload.tajuk || rec.tajuk);
  const tarikhTerimaRaw = _reportText_(payload.tarikh_terima || rec.tarikh_terima);
  const ringkasanButiran = _reportText_(payload.ringkasan_butiran || rec.ringkasan_butiran);
  const lokasi = _reportText_(payload.lokasi || rec.lokasi);
  const namaPengadu = _reportText_(payload.nama_pengadu || rec.nama_pengadu);
  const tarikhSiasatanRaw = _reportText_(payload.tarikh_siasatan);
  const pegawaiUserId = _reportText_(payload.pegawai_penyiasat_user_id || rec.assigned_user_id || s.userId);
  const pegawaiNama = _reportText_(payload.nama);
  const pegawaiJawatan = _reportText_(payload.jawatan);

  const tarikhTerimaIso = tarikhTerimaRaw ? _reportNormalizeDateOrBlank_(tarikhTerimaRaw) : "";
  const tarikhSiasatanIso = tarikhSiasatanRaw ? _reportNormalizeDateOrBlank_(tarikhSiasatanRaw) : "";

  if (!idMaklumBalas) return { ok: false, message: "ID Maklum Balas wajib diisi." };
  if (!jenisAduan) return { ok: false, message: "Jenis Aduan wajib diisi." };
  if (!jenisMaklumbalasAwam) return { ok: false, message: "Jenis Maklumbalas Awam wajib dipilih." };
  if (!tajuk) return { ok: false, message: "Tajuk Aduan wajib diisi." };
  if (!tarikhTerimaIso) return { ok: false, message: "Tarikh Aduan Diterima wajib diisi dalam format tarikh yang sah." };
  if (!ringkasanButiran) return { ok: false, message: "Ringkasan Butiran Aduan wajib diisi." };
  if (!lokasi) return { ok: false, message: "Premis Yang Di Adu wajib diisi." };
  if (!namaPengadu) return { ok: false, message: "Nama Pengadu wajib diisi." };
  if (!tarikhSiasatanIso) return { ok: false, message: "Tarikh Siasatan wajib diisi dalam format tarikh yang sah." };
  if (!pegawaiNama) return { ok: false, message: "Nama Pegawai Penyiasat wajib diisi." };
  if (!pegawaiJawatan) return { ok: false, message: "Jawatan Pegawai Penyiasat wajib diisi." };

  // Bahagian 2
  const kategori = _reportText_(payload.kategori);
  const parlimen = _reportText_(payload.parlimen);
  const jenisPremis = _reportText_(payload.jenis_premis);
  const subPremis = _reportText_(payload.subkategori_premis_makanan);
  const subProduk = _reportText_(payload.subkategori_produk_makanan);
  const statusPensijilan = _reportText_(payload.status_pensijilan);
  const statusPemeriksaanTerdahulu = _reportText_(payload.status_pemeriksaan_terdahulu);
  const markahPemeriksaanSemasa = _reportText_(payload.markah_pemeriksaan_semasa);
  const tindakanPenguatkuasaan = _reportText_(payload.tindakan_penguatkuasaan);

  if (!kategori) return { ok: false, message: "Kategori wajib dipilih." };
  if (!parlimen) return { ok: false, message: "Parlimen wajib dipilih." };
  if (!jenisPremis) return { ok: false, message: "Jenis Premis wajib dipilih." };
  if (!subPremis) return { ok: false, message: "Sub Kategori Premis Makanan wajib dipilih." };
  if (!subProduk) return { ok: false, message: "Sub Kategori Produk Makanan wajib dipilih." };
  if (!statusPensijilan) return { ok: false, message: "Status Pensijilan wajib dipilih." };
  if (!statusPemeriksaanTerdahulu) return { ok: false, message: "Status Pemeriksaan Premis Terdahulu wajib dipilih." };
  if (!markahPemeriksaanSemasa) return { ok: false, message: "Markah Pemeriksaan Semasa wajib diisi." };
  if (!/^(100(\.0{1,2})?|[0-9]?\d(\.\d{1,2})?)$/.test(markahPemeriksaanSemasa)) {
    return { ok: false, message: "Markah Pemeriksaan Semasa mesti antara 0.0 hingga 100.0." };
  }
  if (!tindakanPenguatkuasaan) return { ok: false, message: "Tindakan Yang Telah Diambil wajib diisi." };

  // Bahagian 3
  const findings = _normalizeReportFindings_(payload.findings || []);
  if (!findings.length) return { ok: false, message: "Sekurang-kurangnya satu Isu dan Penemuan wajib diisi." };

  for (let i = 0; i < findings.length; i++) {
    if (!findings[i].isu) return { ok: false, message: "Isu bagi item #" + (i + 1) + " wajib diisi." };
    if (!findings[i].penemuan) return { ok: false, message: "Penemuan bagi item #" + (i + 1) + " wajib diisi." };
  }

  const rumusan = _reportText_(payload.rumusan);
  if (!rumusan) return { ok: false, message: "Rumusan Aduan wajib dipilih." };

  const kelemahanList = _normalizeKelemahanList_(payload.kelemahan_list || []);
  if (!kelemahanList.length) return { ok: false, message: "Sekurang-kurangnya satu Kelemahan wajib diisi." };

  // Bahagian 4
  const kategoriPenyelesaianAduan = _reportText_(payload.kategori_penyelesaian_aduan);
  if (!kategoriPenyelesaianAduan) return { ok: false, message: "Kategori Penyelesaian Aduan wajib dipilih." };

  // Bahagian 5
  const weaknessImages = _normalizeWeaknessImages_(payload.weakness_images || []);
  for (let i = 0; i < weaknessImages.length; i++) {
    const item = weaknessImages[i];
    const hasAny = !!(item.url || item.file_id || item.keterangan);
    if (!hasAny) continue;

    if (!item.url && !item.file_id) {
      return { ok: false, message: "Gambar bagi item lampiran #" + (i + 1) + " belum dimuat naik." };
    }
    if (!item.keterangan) {
      return { ok: false, message: "Keterangan bagi item lampiran #" + (i + 1) + " wajib diisi." };
    }
  }

  const sh = found.sheet;
  const row = found.row;
  const map = _headerMap_(found.headers);
  const nowIso = new Date().toISOString();

  const investigatorUserId = _reportText_(rec.assigned_user_id || s.userId);
  const investigatorProfile = investigatorUserId ? _findUserFullProfile_(investigatorUserId) : null;

  const signFileId = _reportText_(
    rec.report_tandatangan_file_id ||
    (investigatorProfile && investigatorProfile.signature_file_id) ||
    ""
  );

  const signUrl = _reportText_(
    rec.report_tandatangan_url ||
    (signFileId ? ("https://drive.google.com/thumbnail?id=" + signFileId + "&sz=w1200") : "") ||
    ""
  );

  const firstFinding = findings[0] || { isu: "", penemuan: "" };

  _setCellByHeader_(sh, row, map, "report_status", "DRAF");
  _setCellByHeader_(sh, row, map, "report_updated_at", nowIso);

  // Bahagian 1
  _setCellByHeader_(sh, row, map, "report_id_maklumbalas", idMaklumBalas);
  _setCellByHeader_(sh, row, map, "report_jenis_aduan", jenisAduan);
  _setCellByHeader_(sh, row, map, "report_jenis_maklumbalas_awam", jenisMaklumbalasAwam);
  _setCellByHeader_(sh, row, map, "report_tajuk", tajuk);
  _setCellByHeader_(sh, row, map, "report_tarikh_terima", tarikhTerimaIso);
  _setCellByHeader_(sh, row, map, "report_ringkasan_butiran", ringkasanButiran);
  _setCellByHeader_(sh, row, map, "report_lokasi", lokasi);
  _setCellByHeader_(sh, row, map, "report_nama_pengadu", namaPengadu);
  _setCellByHeader_(sh, row, map, "report_tarikh_siasatan", tarikhSiasatanIso);
  _setCellByHeader_(sh, row, map, "report_pegawai_penyiasat_user_id", pegawaiUserId);
  _setCellByHeader_(sh, row, map, "report_nama_pegawai", pegawaiNama);
  _setCellByHeader_(sh, row, map, "report_jawatan_pegawai", pegawaiJawatan);

  // Bahagian 2
  _setCellByHeader_(sh, row, map, "report_kategori", kategori);
  _setCellByHeader_(sh, row, map, "report_parlimen", parlimen);
  _setCellByHeader_(sh, row, map, "report_jenis_premis", jenisPremis);
  _setCellByHeader_(sh, row, map, "report_subkategori_premis_makanan", subPremis);
  _setCellByHeader_(sh, row, map, "report_subkategori_produk_makanan", subProduk);
  _setCellByHeader_(sh, row, map, "report_status_pensijilan", statusPensijilan);
  _setCellByHeader_(sh, row, map, "report_status_pemeriksaan_terdahulu", statusPemeriksaanTerdahulu);
  _setCellByHeader_(sh, row, map, "report_markah_pemeriksaan_semasa", markahPemeriksaanSemasa);
  _setCellByHeader_(sh, row, map, "report_tindakan_penguatkuasaan", tindakanPenguatkuasaan);

  // Bahagian 3
  _setCellByHeader_(sh, row, map, "report_findings_json", JSON.stringify(findings));
  _setCellByHeader_(sh, row, map, "report_isu", firstFinding.isu || "");
  _setCellByHeader_(sh, row, map, "report_penemuan", firstFinding.penemuan || "");
  _setCellByHeader_(sh, row, map, "report_rumusan", rumusan);
  _setCellByHeader_(sh, row, map, "report_kelemahan_list_json", JSON.stringify(kelemahanList));
  _setCellByHeader_(sh, row, map, "report_kelemahan", _joinKelemahanForTemplate_(kelemahanList));

  // Bahagian 4
  _setCellByHeader_(sh, row, map, "report_kategori_penyelesaian_aduan", kategoriPenyelesaianAduan);

  // Tandatangan
  _setCellByHeader_(sh, row, map, "report_tandatangan_file_id", signFileId);
  _setCellByHeader_(sh, row, map, "report_tandatangan_url", signUrl);

  // Bahagian 5
  _setCellByHeader_(sh, row, map, "report_weakness_images_json", JSON.stringify(weaknessImages));
  _syncWeaknessLegacyColumns_(sh, row, map, weaknessImages);

  SpreadsheetApp.flush();

  return {
    ok: true,
    message: "Draf laporan berjaya disimpan."
  };
}

function uploadReportImage(token, complaintId, slot, base64Image, filename, mimeType, itemIndex) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };

  const found = _findComplaintById_(complaintId);
  if (!found || !found.record) {
    return { ok: false, message: "Rekod aduan tidak dijumpai." };
  }

  const rec = found.record || {};
  if (!_canEditReport_(s, rec)) {
    return { ok: false, message: "Anda tidak mempunyai akses untuk memuat naik gambar laporan ini." };
  }

  slot = _reportText_(slot);
  base64Image = _reportText_(base64Image);
  filename = _reportText_(filename) || ("report_" + slot + ".jpg");
  mimeType = _reportText_(mimeType) || "image/jpeg";
  itemIndex = Number(itemIndex || 0);

  if (!base64Image) return { ok: false, message: "Fail gambar tidak diterima." };
  if (!/^image\/(png|jpeg|jpg)$/i.test(mimeType)) {
    return { ok: false, message: "Format gambar tidak sah. Sila gunakan PNG/JPG/JPEG." };
  }

  let bytes;
  try {
    bytes = Utilities.base64Decode(base64Image);
  } catch (e) {
    return { ok: false, message: "Format fail gambar tidak sah." };
  }

  const sh = found.sheet;
  const row = found.row;
  const map = _headerMap_(found.headers);

  try {
    const folder = _getOrCreateReportImageFolder_();
    const ext = mimeType.toLowerCase().indexOf("png") >= 0 ? "png" : "jpg";
    const cleanName = "SLAM_REPORT_" + slot.toUpperCase() + "_" + Date.now() + "." + ext;
    const blob = Utilities.newBlob(bytes, mimeType, cleanName);
    const file = folder.createFile(blob);

    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {}

    const fileId = file.getId();
    const url = "https://drive.google.com/uc?export=view&id=" + fileId;

    if (slot === "gambar_hadapan_premis") {
      const oldFileId = _reportText_(rec.report_gambar_hadapan_premis_file_id);
      if (oldFileId) {
        try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch (e) {}
      }

      _setCellByHeader_(sh, row, map, "report_gambar_hadapan_premis", url);
      _setCellByHeader_(sh, row, map, "report_gambar_hadapan_premis_file_id", fileId);
    } else if (slot === "gambar_kelemahan_item") {
      const weaknessItems = _readWeaknessImagesFromRecord_(rec);

      while (weaknessItems.length <= itemIndex) {
        weaknessItems.push({ url: "", file_id: "", keterangan: "" });
      }

      const oldFileId = _reportText_(weaknessItems[itemIndex].file_id);
      if (oldFileId) {
        try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch (e) {}
      }

      weaknessItems[itemIndex].url = url;
      weaknessItems[itemIndex].file_id = fileId;

      _setCellByHeader_(sh, row, map, "report_weakness_images_json", JSON.stringify(weaknessItems));
      _syncWeaknessLegacyColumns_(sh, row, map, weaknessItems);
    } else {
      return { ok: false, message: "Slot gambar tidak sah." };
    }

    _setCellByHeader_(sh, row, map, "report_status", "DRAF");
    _setCellByHeader_(sh, row, map, "report_updated_at", new Date().toISOString());

    SpreadsheetApp.flush();

    return {
      ok: true,
      message: "Gambar berjaya dimuat naik.",
      slot: slot,
      item_index: itemIndex,
      file_id: fileId,
      url: url
    };
  } catch (e) {
    return {
      ok: false,
      message: "Gagal memuat naik gambar: " + (e && e.message ? e.message : e)
    };
  }
}