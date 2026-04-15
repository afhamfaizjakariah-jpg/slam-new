/******************************
 * 11_Utils.gs
 ******************************/

function _formatMalayLongDate_(val) {
  const d = _safeDate_(val) || _safeDate_(_normalizeDate_(val));
  if (!d) return "";

  const bulan = [
    "Januari", "Februari", "Mac", "April", "Mei", "Jun",
    "Julai", "Ogos", "September", "Oktober", "November", "Disember"
  ];

  return d.getDate() + " " + bulan[d.getMonth()] + " " + d.getFullYear();
}

function _formatMalayLongDateUpper_(val) {
  const d = _safeDate_(val) || _safeDate_(_normalizeDate_(val));
  if (!d) return "";

  const bulan = [
    "JANUARI", "FEBRUARI", "MAC", "APRIL", "MEI", "JUN",
    "JULAI", "OGOS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DISEMBER"
  ];

  const hari = String(d.getDate()).padStart(2, "0");
  return hari + " " + bulan[d.getMonth()] + " " + d.getFullYear();
}

function _padNo_(n, len) {
  return String(n || 0).padStart(len || 4, "0");
}