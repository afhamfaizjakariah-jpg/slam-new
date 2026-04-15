/******************************
 * 10_Parsing.gs
 ******************************/

function _guessJenisFromFilename_(filename) {
  const f = String(filename || "").toUpperCase();
  if (f.indexOf("MOH") >= 0) return "MOH";
  if (f.indexOf("PCB") >= 0) return "PCB";
  if (f.indexOf("JPA") >= 0) return "JPA";
  if (f.indexOf("EMAIL") >= 0 || f.indexOf("EMEL") >= 0) return "EMEL";
  return "";
}

/** ====== Infer sumber ====== **/
function _inferSumberFromId_(textOrId) {
  const s = String(textOrId || "").toUpperCase().replace(/\u00A0/g, " ").trim();
  const m = s.match(/\b(MOH|PCB|JPA|EMEL)\b[\s\.\-:]*\s*(\d{0,})\b/i);
  if (m) return String(m[1]).toUpperCase();
  return "";
}

/** ====== Normalisasi ID ====== **/
function _normalizeIdMaklumBalas_(val) {
  let s = String(val || "");
  s = s.replace(/\u00A0/g, " ");
  s = s.replace(/[\u200B-\u200D\uFEFF]/g, "");
  s = s.replace(/\s+/g, " ").trim();

  const m = s.toUpperCase().match(/\b(MOH|PCB|JPA)\b[\s\.\-:]*\s*(\d{3,})\b/);
  if (m) return m[1] + "." + m[2];

  return s.toLowerCase();
}

function _cleanText_(t) {
  return String(t || "")
    .replace(/\u00A0/g, " ")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\r/g, "\n")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function _extractFirst_(t, re) {
  const m = String(t || "").match(re);
  return m ? String(m[1] || "").trim() : "";
}

function _extractSectionByLabels_(t, startLabelRe, endLabelRes) {
  const s = String(t || "");
  const startIdx = s.search(startLabelRe);
  if (startIdx === -1) return "";
  const sub = s.slice(startIdx);
  const m1 = sub.match(startLabelRe);
  if (!m1) return "";
  const after = sub.slice(m1[0].length);

  let endPos = -1;
  (endLabelRes || []).forEach(function(re) {
    const p = after.search(re);
    if (p !== -1) endPos = (endPos === -1 || p < endPos) ? p : endPos;
  });

  return String((endPos === -1 ? after : after.slice(0, endPos)) || "").trim();
}

/** ====== Smart extract BUTIRAN fallback ====== **/
function _extractButiranSmart_(rawText) {
  const t = _cleanText_(rawText);

  const starts = [
    /Butiran\s*Aduan\s*[:：]?\s*/i,
    /Butiran\s*[:：]?\s*/i,
    /Keterangan\s*[:：]?\s*/i,
    /Perihal\s*[:：]?\s*/i,
    /Aduan\s*[:：]?\s*/i
  ];
  const ends = [
    /(?:\n|\s)Lokasi\s*[:：]\s*/i,
    /(?:\n|\s)Premis\s*Yang\s*Di\s*Adu\s*[:：]\s*/i,
    /(?:\n|\s)Jabatan\s*[:：]\s*/i,
    /(?:\n|\s)Sumber\s*[:：]\s*/i,
    /(?:\n|\s)Isu\s*[:：]\s*/i,
    /(?:\n|\s)Sektor\s*[:：]\s*/i,
    /(?:\n|\s)Tahap\s*Kesukaran\s*[:：]\s*/i,
    /(?:\n|\s)Tahap\s*Sensitiviti\s*[:：]\s*/i,
    /(?:\n|\s)Kategori\s*[:：]\s*/i,
    /(?:\n|\s)Nama\s*[:：]\s*/i
  ];

  for (const st of starts) {
    const sec = _extractSectionByLabels_(t, st, ends);
    if (sec && sec.replace(/\s+/g, "").length >= 20) return sec;
  }
  return "";
}

function _parseComplaintText_(rawText, filename) {
  const parseNotes = [];
  const t = _cleanText_(rawText);

  let id = _extractFirst_(t, /ID\s*Maklum\s*Balas\s*:\s*([A-Z]{2,5}\s*[\.\-]?\s*\d{3,})/i);
  if (!id) {
    const m = t.match(/\b(MOH|PCB|JPA)\b[\s\.\-:]*\s*\d{3,}\b/i);
    if (m) id = m[0].toUpperCase().replace(/\s+/g, "");
  }
  if (id) {
    const mm = String(id).toUpperCase().match(/\b(MOH|PCB|JPA)\b[\s\.\-:]*\s*(\d{3,})\b/);
    if (mm) id = mm[1] + "." + mm[2];
  }

  let terima = _extractFirst_(t, /Terima\s*:\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/i);
  if (!terima) terima = _extractFirst_(t, /Tarikh\s*Terima\s*:\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/i);

  let jenis = _extractFirst_(t, /Jenis\s*Maklum\s*Balas\s*:\s*([^\n]+)/i);
  if (!jenis) jenis = _guessJenisFromFilename_(filename);

  let tajuk = _extractFirst_(t, /Tajuk\s*:\s*([^\n]+)/i);

  let butiran = _extractSectionByLabels_(
    t,
    /Butiran\s*:\s*/i,
    [
      /(?:\n|\s)Lokasi\s*:\s*/i,
      /(?:\n|\s)Jabatan\s*:\s*/i,
      /(?:\n|\s)Sumber\s*:\s*/i,
      /(?:\n|\s)Isu\s*:\s*/i,
      /(?:\n|\s)Sektor\s*:\s*/i,
      /(?:\n|\s)Tahap\s*Kesukaran\s*:\s*/i,
      /(?:\n|\s)Tahap\s*Sensitiviti\s*:\s*/i,
      /(?:\n|\s)Kategori\s*:\s*/i,
      /(?:\n|\s)Nama\s*:\s*/i
    ]
  );
  if (!String(butiran || "").trim()) {
    const b2 = _extractButiranSmart_(t);
    if (b2) {
      butiran = b2;
      parseNotes.push("Butiran smart extracted.");
    }
  }

  let lokasi = _extractFirst_(t, /Lokasi\s*:\s*([^\n]+)/i);
  if (!lokasi) lokasi = _extractFirst_(t, /Premis\s*Yang\s*Di\s*Adu\s*[:：]\s*([^\n]+)/i);

  let tahap = _extractFirst_(t, /Tahap\s*Kesukaran\s*[:：]\s*([^\n]+)/i);
  if (!tahap) tahap = "Biasa";
  const tahapUp = String(tahap).toUpperCase();
  if (tahapUp.indexOf("KOMPLEKS") >= 0) tahap = "Kompleks";
  else if (tahapUp.indexOf("BIASA") >= 0) tahap = "Biasa";

  let nama = _extractFirst_(t, /(?:^|\n)\s*Nama\s*:\s*([^\n]+)/i);
  if (nama) {
    nama = nama.replace(/\s+(No\.\s*Pengenalan|NRIC|Jenis\s*Pelanggan|Kategori|Jantina|Umur|Bangsa|Kewarganegaraan|Pekerjaan|Alamat|Poskod|Negara|Negeri|Bandar|No\.\s*Telefon|Telefon|Faks|E-?mel)\s*:.*$/i, "").trim();
  }

  if (id) id = String(id).replace(/\s+/g, "").trim();
  if (tajuk) tajuk = String(tajuk).replace(/\s+/g, " ").trim();
  if (lokasi) lokasi = String(lokasi).replace(/\s+/g, " ").trim();
  if (nama) nama = String(nama).replace(/\s+/g, " ").trim();

  return {
    id_maklumbalas: id || "",
    tarikh_terima: _normalizeDate_(terima) || "",
    jenis_maklumbalas: jenis || "",
    tajuk: tajuk || "",
    butiran: butiran || "",
    lokasi: lokasi || "",
    tahap_kesukaran: tahap || "Biasa",
    nama_pengadu: nama || "",
    parse_notes: parseNotes.join(" | ")
  };
}

function _makeNeutralSummary_(butiran) {
  const cleaned = _cleanButiranForSummary_(butiran);
  if (!cleaned) return "";

  const sentences = _splitToSentences_(cleaned);
  if (!sentences.length) return _finalizeSummaryParagraph_(cleaned);

  const scored = sentences.map((s, i) => ({ s, i, score: _scoreSentence_(s) }));
  scored.sort((a, b) => b.score - a.score);

  const picked = [];
  for (const item of scored) {
    if (picked.length >= 4) break;
    if (_looksTooGenericSentence_(item.s)) continue;

    const tooSimilar = picked.some(p => _jaccardSim_(p, item.s) >= 0.72);
    if (!tooSimilar) picked.push(item.s);
  }

  if (!picked.length) {
    picked.push(...sentences.slice(0, 3));
  }

  const ordered = picked
    .map(s => ({ s, idx: sentences.indexOf(s) }))
    .sort((a, b) => a.idx - b.idx)
    .map(x => x.s);

  return _finalizeSummaryParagraph_(ordered.join(" "));
}

function _finalizeSummaryParagraph_(text) {
  let s = String(text || "").trim();
  if (!s) return "";

  s = _stripImperatives_(s);
  s = _professionalizeMalay_(s);
  s = _forceThirdPerson_(s);

  const ss = _splitToSentences_(s);
  s = ss.slice(0, 4).join(" ").trim();

  const max = 650;
  if (s.length > max) s = s.slice(0, max).trim().replace(/[,\s]+$/, "") + "…";

  return _fixCasingPunctuation_(s);
}

function _cleanButiranForSummary_(butiran) {
  let s = String(butiran || "").trim();
  if (!s) return "";

  s = s.replace(/\u00A0/g, " ")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\r/g, "\n")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  s = s.replace(/\n?\s*(Lokasi|Jabatan|Sumber|Isu|Sektor|Tahap\s*Kesukaran|Tahap\s*Sensitiviti|Kategori|Nama)\s*:\s*[\s\S]*$/i, "").trim();
  s = s.split("\n").map(x => x.trim()).filter(Boolean).join(" ");
  s = s.replace(/\s{2,}/g, " ").trim();

  return s;
}

function _splitToSentences_(text) {
  const s = String(text || "").trim();
  if (!s) return [];
  const parts = s.split(/(?<=[.!?])\s+/g).map(x => x.trim()).filter(Boolean);
  if (parts.length >= 2) return parts;
  return s.split(/\s*,\s*/g).map(x => x.trim()).filter(Boolean).map(x => x.endsWith(".") ? x : (x + "."));
}

function _scoreSentence_(sentence) {
  const s = String(sentence || "").toLowerCase();
  const keywords = [
    "kotor","kebersihan","lalat","tandas","bau","busuk","najis","lipas","tikus",
    "makanan","masak","mentah","basi","keracunan","sakit perut","cirit",
    "pekerja","sarung tangan","penutup kepala","merokok","batuk","demam",
    "warga asing","vaksin","tb","batuk kering","peralatan","pinggan","air",
    "sinki","lantai","sampah","premis","gerai","restoran"
  ];
  let score = 0;
  for (const k of keywords) if (s.indexOf(k) >= 0) score += 3;
  if (/\d/.test(s)) score += 1;
  if (s.length >= 60 && s.length <= 220) score += 2;
  if (/kerana|apabila|namun|serta|dan|yang/.test(s)) score += 1;
  return score;
}

function _looksTooGenericSentence_(s) {
  const t = String(s || "").toLowerCase().trim();
  if (!t) return true;
  if (t.replace(/\s+/g, " ").length < 35) return true;
  if (/^(aduan|pengadu)\b.*\b(berkaitan|memaklumkan)\b.*$/i.test(s) && t.length < 70) return true;
  return false;
}

function _jaccardSim_(a, b) {
  const tok = (x) => String(x || "")
    .toLowerCase()
    .replace(/[^a-z0-9\u00C0-\u024F\s]/gi, " ")
    .split(/\s+/)
    .map(w => w.trim())
    .filter(w => w.length >= 4);

  const A = tok(a), B = tok(b);
  const setA = {}, setB = {}, all = {};
  A.forEach(w => { setA[w] = 1; all[w] = 1; });
  B.forEach(w => { setB[w] = 1; all[w] = 1; });

  let inter = 0;
  Object.keys(setA).forEach(k => { if (setB[k]) inter++; });
  const union = Object.keys(all).length;
  return union ? (inter / union) : 0;
}

function _stripImperatives_(text) {
  let s = String(text || "").trim();
  if (!s) return "";
  return s
    .replace(/\b(sila\s+check|sila\s+semak|tolong\s+check|tolong\s+semak|please\s+check|please\s+inspect)\b/gi, "")
    .replace(/\b(jangan\s+makan|jgn\s+makan)\b/gi, "")
    .replace(/!+/g, ".")
    .replace(/\s{2,}/g, " ")
    .trim();
}

function _professionalizeMalay_(text) {
  let s = String(text || "").trim();
  if (!s) return "";

  s = _normalizeEntities_(s);

  s = s
    .replace(/\b(besar kemungkinan|kemungkinan besar)\b/gi, "pengadu mengesyaki")
    .replace(/\b(betul\s*2|betul-betul)\b/gi, "dilaporkan")
    .replace(/\bmajority\b/gi, "kebanyakannya")
    .replace(/\btapi\b/gi, "namun")
    .replace(/\blepas tu\b/gi, "seterusnya")
    .replace(/\bskrg\b/gi, "kini")
    .replace(/\bni\b/gi, "ini")
    .replace(/\bbyr\b/gi, "bayaran")
    .replace(/\bpkrja\b/gi, "pekerja")
    .replace(/\borg\b/gi, "orang")
    .replace(/\bpremise\b/gi, "premis")
    .replace(/\brestoran\b/gi, "restoran")
    .replace(/\bgerai2\b/gi, "gerai-gerai")
    .replace(/\bwarga2\b/gi, "warga")
    .replace(/\bterlampau\b/gi, "amat")
    .replace(/\bteruk\b/gi, "kurang memuaskan");

  return s.replace(/\s{2,}/g, " ").trim();
}

function _normalizeEntities_(text) {
  let s = String(text || "").trim();
  if (!s) return "";
  s = s.replace(/\bgerai\s*2\b/gi, "gerai-gerai");
  s = s.replace(/\brestaurant\b/gi, "restoran");
  s = s.replace(/\borang\s+cina\b/gi, "orang Cina");
  s = s.replace(/nama\s+gerai\s+pakai\s+nama/gi, "nama gerai menggunakan nama");
  s = s.replace(/batuk\s+tidak\s+henti/gi, "batuk berterusan");
  s = s.replace(/\bTB\b/g, "batuk kering (TB)");
  return s.trim();
}

function _fixCasingPunctuation_(text) {
  let s = String(text || "").trim();
  if (!s) return "";
  s = s.replace(/\.{2,}/g, ".").replace(/\.\s*(\S)/g, ". $1").trim();
  s = s.charAt(0).toUpperCase() + s.slice(1);
  if (!/[.!?…]$/.test(s)) s += ".";
  return s.trim();
}

function _forceThirdPerson_(text) {
  let s = String(text || "").trim();
  if (!s) return "";
  s = s.replace(/\b(saya|aku|kami|kita)\b/gi, "pengadu").replace(/\b(pengadu pengadu)\b/gi, "pengadu");
  const startsOk = /^(Pengadu|Aduan|Kes|Laporan|Berdasarkan)/i.test(s);
  if (!startsOk) s = "Pengadu memaklumkan bahawa " + s;
  s = s.replace(/^Pengadu memaklumkan bahawa\s+(Aduan|Kes|Laporan)\b/i, "$1");
  return s.trim();
}

/** ========= Groq ========= **/
function _getGroqApiKey_() {
  return PropertiesService.getScriptProperties().getProperty("GROQ_API_KEY") || "";
}

function _summarizeButiranWithGroq_(butiranText, filename, opt) {
  opt = opt || {};
  const strict = !!opt.strict;
  const notes = Array.isArray(opt.notes) ? opt.notes : null;

  const apiKey = _getGroqApiKey_();
  if (!apiKey) return { ok: false, note: "GROQ_API_KEY tidak dijumpai (Script Properties)." };

  let input = String(butiranText || "").trim();
  if (!input) return { ok: false, note: "Input ringkasan kosong." };

  input = input
    .replace(/\u00A0/g, " ")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\r/g, "\n")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  if (input.length > CONFIG.GROQ_SUMMARY_MAX_INPUT_CHARS) {
    input = input.slice(0, CONFIG.GROQ_SUMMARY_MAX_INPUT_CHARS);
    if (notes) notes.push("Input truncated to " + CONFIG.GROQ_SUMMARY_MAX_INPUT_CHARS + " chars.");
  }

  const sysStrict =
    "Anda ialah pegawai kanan kesihatan awam. Tulis ringkasan profesional dalam Bahasa Melayu formal berdasarkan butiran aduan pengguna. " +
    "Gunakan gaya pihak ketiga yang neutral, tanpa menokok tambah fakta, tanpa nasihat, tanpa arahan, tanpa spekulasi di luar teks, " +
    "dan tanpa mengulang ayat yang sama. Fokus pada fakta pemerhatian, premis, kebersihan, pekerja, makanan, dan risiko yang dinyatakan. " +
    "Panjang sasaran 80 hingga 140 patah perkataan dalam satu perenggan padat.";

  const sysLoose =
    "Ringkaskan butiran aduan dalam Bahasa Melayu formal, neutral, dan profesional. " +
    "Tulis dalam bentuk satu perenggan pihak ketiga. Jangan guna bullet, jangan beri cadangan, jangan tambah fakta baru, jangan menyebut bahawa ini ialah ringkasan AI.";

  const userPrompt =
    "Nama fail: " + String(filename || "aduan.pdf") + "\n\n" +
    "Butiran aduan:\n" + input + "\n\n" +
    "Hasilkan ringkasan akhir sahaja.";

  const models = [
    String(CONFIG.GROQ_SUMMARY_MODEL || "").trim(),
    String(CONFIG.GROQ_SUMMARY_FALLBACK_MODEL || "").trim()
  ].filter(Boolean);

  const attemptMax = Math.max(1, Number(CONFIG.GROQ_RETRY_ATTEMPTS || 1));

  for (let mi = 0; mi < models.length; mi++) {
    const model = models[mi];

    for (let attempt = 1; attempt <= attemptMax; attempt++) {
      try {
        const payload = {
          model: model,
          temperature: strict ? 0.2 : 0.3,
          max_tokens: 240,
          messages: [
            { role: "system", content: strict ? sysStrict : sysLoose },
            { role: "user", content: userPrompt }
          ]
        };

        const res = UrlFetchApp.fetch("https://api.groq.com/openai/v1/chat/completions", {
          method: "post",
          contentType: "application/json",
          muteHttpExceptions: true,
          headers: {
            Authorization: "Bearer " + apiKey
          },
          payload: JSON.stringify(payload)
        });

        const code = res.getResponseCode();
        const body = res.getContentText() || "";

        if (code >= 200 && code < 300) {
          let data = {};
          try { data = JSON.parse(body); } catch (e) {}

          let out = "";
          try {
            out = String((((data || {}).choices || [])[0] || {}).message?.content || "").trim();
          } catch (e) {}

          out = out.replace(/^["'`\s]+|["'`\s]+$/g, "").trim();
          out = out.replace(/\n{2,}/g, "\n").replace(/\s{2,}/g, " ").trim();

          if (!out) {
            if (notes) notes.push("Groq empty content model=" + model + " attempt=" + attempt);
            continue;
          }

          out = _finalizeSummaryParagraph_(out);

          if (_isTooGenericSummary_(input, out)) {
            if (notes) notes.push("Groq generic summary rejected model=" + model + " attempt=" + attempt);
            continue;
          }

          return { ok: true, text: out, model: model, attempt: attempt };
        }

        if (notes) notes.push("Groq HTTP " + code + " model=" + model + " attempt=" + attempt);

        if (code === 429 || code >= 500) {
          Utilities.sleep((CONFIG.GROQ_RETRY_BASE_SLEEP_MS || 900) * attempt);
          continue;
        }
      } catch (e) {
        if (notes) notes.push("Groq error model=" + model + " attempt=" + attempt + ": " + (e && e.message ? e.message : e));
        Utilities.sleep((CONFIG.GROQ_RETRY_BASE_SLEEP_MS || 900) * attempt);
      }
    }
  }

  return { ok: false, note: "Semua percubaan ringkasan Groq gagal." };
}

function _isTooGenericSummary_(input, summary) {
  const stop = {
    "aduan":1,"pengadu":1,"melaporkan":1,"memaklumkan":1,"mengenai":1,"terhadap":1,"berkaitan":1,"isu":1,"tahap":1,
    "ini":1,"itu":1,"dan":1,"yang":1,"dengan":1,"untuk":1,"pada":1,"di":1,"ke":1,"dalam":1,"adalah":1,
    "the":1,"and":1,"with":1,"for":1,"that":1,"this":1,"are":1,"was":1,"were":1
  };

  const tok = (s) => String(s || "")
    .toLowerCase()
    .replace(/[^a-z0-9\u00C0-\u024F\s]/gi, " ")
    .split(/\s+/)
    .map(w => w.trim())
    .filter(w => w.length >= 3 && !stop[w]);

  const src = tok(input);
  const sum = tok(summary);

  if (!sum.length) return true;
  if (summary.length < 60) return true;

  const set = {};
  src.forEach(w => set[w] = 1);

  let overlap = 0;
  sum.forEach(w => { if (set[w]) overlap++; });

  return overlap < Math.max(3, Math.floor(sum.length * 0.18));
}

function _escapeForDocReplace_(text) {
  return String(text || "").replace(/\$/g, "$$$$");
}

function _buildAppointmentLetterRefNo_(rec, chosenUser) {
  const now = new Date();
  const year = now.getFullYear();
  const prefix = String(CONFIG.APPOINTMENT_REF_PREFIX || "PKK(S)100-9/3/").trim();
  const idRef = String(rec.id_maklumbalas || rec.complaint_id || "ADUAN").replace(/[^\w\-\.]+/g, "");
  return prefix + year + "/" + idRef;
}