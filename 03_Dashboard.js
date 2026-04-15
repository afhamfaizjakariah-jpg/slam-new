/******************************
 * 03_Dashboard.gs
 ******************************/

function _dashboardSourceLabel_(rec) {
  const raw = String(rec.source || rec.report_jenis_aduan || rec.jenis_maklumbalas || "").trim();
  return raw || "-";
}

function _dashboardCardRow_(rec) {
  return {
    complaint_id: String(rec.complaint_id || "").trim(),
    id_maklumbalas: String(rec.id_maklumbalas || rec.complaint_id || "").trim(),
    tajuk: String(rec.tajuk || "").trim(),
    premis_nama: String(rec.premis_nama || _derivePremisName_(rec.lokasi, rec.tajuk) || "").trim(),
    tarikh_terima: _formatDateMalay_(rec.tarikh_terima || ""),
    tarikh_terima_iso: String(_normalizeDate_(rec.tarikh_terima || "") || "").trim(),
    bulan: _safeDate_(rec.tarikh_terima || rec.created_at || "") ? Utilities.formatDate(new Date(rec.tarikh_terima || rec.created_at), Session.getScriptTimeZone(), "MMMM") : _dashboardMonthName_(rec.tarikh_terima || rec.created_at || ""),
    source: _dashboardSourceLabel_(rec),
    status_card: String(_computeEffectiveCardStatus_(rec) || "Baharu").trim(),
    assigned_to: String(rec.assigned_to || "").trim(),
    rumusan: String(rec.report_rumusan || "").trim(),
    tahap_kesukaran: String(rec.tahap_kesukaran || "").trim(),
    nama_pengadu: String(rec.nama_pengadu || "").trim(),
    lokasi: String(rec.lokasi || "").trim()
  };
}

function _dashboardMonthName_(val) {
  const d = _safeDate_(val) || _safeDate_(_normalizeDate_(val));
  if (!d) return "";
  const bulan = ["Januari", "Februari", "Mac", "April", "Mei", "Jun", "Julai", "Ogos", "September", "Oktober", "November", "Disember"];
  return bulan[d.getMonth()];
}

function _dashboardSummaryFromRows_(rows) {
  const summary = { jumlah: 0, baru: 0, pending: 0, selesai: 0, lewat: 0, pindah: 0 };
  (rows || []).forEach(function(r) {
    summary.jumlah++;
    const st = String(r.status_card || "").trim().toUpperCase();
    if (st === "BAHARU" || st === "BARU") summary.baru++;
    else if (st === "PENDING" || st === "TELAH DIAGIHKAN") summary.pending++;
    else if (st === "SELESAI") summary.selesai++;
    else if (st === "LEWAT") summary.lewat++;
    else if (st === "PINDAH") summary.pindah++;
  });
  return summary;
}

function _dashboardTrendFromRows_(rows) {
  const monthMap = ["Jan", "Feb", "Mac", "Apr", "Mei", "Jun", "Jul", "Ogos", "Sep", "Okt", "Nov", "Dis"];
  const trend = monthMap.map(function(m) { return { m: m, c: 0 }; });

  (rows || []).forEach(function(r) {
    const dt = _safeDate_(r.tarikh_terima_iso || r.tarikh_terima || "");
    if (!dt) return;
    trend[dt.getMonth()].c++;
  });
  return trend;
}

function _dashboardStatusBreakdown_(rows) {
  const breakdown = {
    "Berasas": 0,
    "Tidak Berasas": 0,
    "Tidak Berkenaan": 0
  };

  (rows || []).forEach(function(r) {
    const raw = String(r.rumusan || r.report_rumusan || "").trim().toUpperCase();
    if (raw === "BERASAS") breakdown["Berasas"]++;
    else if (raw === "TIDAK BERASAS") breakdown["Tidak Berasas"]++;
    else if (raw === "TIDAK BERKAITAN" || raw === "TIDAK BERKENAAN") breakdown["Tidak Berkenaan"]++;
  });

  return [
    { status: "Berasas", count: breakdown["Berasas"] },
    { status: "Tidak Berasas", count: breakdown["Tidak Berasas"] },
    { status: "Tidak Berkenaan", count: breakdown["Tidak Berkenaan"] }
  ];
}

function getDashboardData(token) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };

  const rows = _getComplaintRows_().map(_dashboardCardRow_);
  const summary = _dashboardSummaryFromRows_(rows);
  const trend = _dashboardTrendFromRows_(rows);
  const status_breakdown = _dashboardStatusBreakdown_(rows);

  rows.sort(function(a, b) {
    const da = _safeDate_(a.tarikh_terima_iso || "");
    const db = _safeDate_(b.tarikh_terima_iso || "");
    return (db ? db.getTime() : 0) - (da ? da.getTime() : 0);
  });

  return {
    ok: true,
    summary: summary,
    trend: trend,
    status_breakdown: status_breakdown,
    cards: rows
  };
}
