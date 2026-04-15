/******************************
 * 01_EntryAuth.gs
 ******************************/

/** ========= ENTRY ========= **/
function doGet() {
  const t = HtmlService.createTemplateFromFile("Index");
  t.appName = CONFIG.APP_NAME;
  return t.evaluate()
    .setTitle(CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** ========= AUTH / SESSION ========= **/
function login(userId, password) {
  userId = String(userId || "").trim();
  password = String(password || "").trim();
  if (!userId || !password) return { ok: false, message: "Sila isi ID Pengguna dan Kata Laluan." };

  const user = _findUser(userId);
  if (!user) return { ok: false, message: "ID Pengguna tidak dijumpai." };

  const storedPw = String(user.password || "").trim();
  if (!storedPw || storedPw !== password) return { ok: false, message: "Kata Laluan tidak tepat." };

  const token = _newToken_();
  const session = {
    token,
    userId: user.userId,
    name: user.name || user.userId,
    role: user.role || "PENGGUNA",
    ts: Date.now()
  };
  CacheService.getScriptCache().put(_sessKey_(token), JSON.stringify(session), CONFIG.SESSION_TTL_SECONDS);

  return {
    ok: true,
    token,
    user: { userId: session.userId, name: session.name, role: session.role }
  };
}

function getSession(token) {
  return _getSession_(token) ? { ok: true } : { ok: false };
}

function getMe(token) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };
  return { ok: true, user: { userId: s.userId, name: s.name, role: s.role } };
}

function logout(token) {
  try { CacheService.getScriptCache().remove(_sessKey_(token)); } catch (e) {}
  return { ok: true };
}

/** ========= INTERNAL HELPERS ========= **/
function _withSession_(token, fn) {
  const s = _getSession_(token);
  if (!s) return { ok: false, message: "Sesi tamat. Sila log masuk semula." };

  try {
    return fn(s);
  } catch (e) {
    Logger.log("withSession error: " + (e && e.stack ? e.stack : e));
    return { ok: false, message: e && e.message ? e.message : String(e) };
  }
}

function _getSession_(token) {
  token = String(token || "").trim();
  if (!token) return null;
  const raw = CacheService.getScriptCache().get(_sessKey_(token));
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function _sessKey_(token) {
  return "SLAM_SESS_" + token;
}

function _newToken_() {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    Utilities.getUuid() + "|" + Date.now()
  );
  return Utilities.base64EncodeWebSafe(bytes).slice(0, 40);
}