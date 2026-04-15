/******************************
 * 00_Config.gs
 ******************************/

const CONFIG = {
  APP_NAME: "SLAM - Sistem Laporan Aduan Makanan",
  SPREADSHEET_ID: "18tISUD7IKbL0-EFUf210cDipvLcDcj1WbO-5GWYnfyY",
  USER_SHEET_CANDIDATES: ["pengguna", "Pengguna", "USERS", "Users", "USER", "User"],
  SESSION_TTL_SECONDS: 60 * 60 * 8,
  COMPLAINTS_SHEET: "COMPLAINTS",
  COMPLAINTS_SHEET_CANDIDATES: ["COMPLAINTS", "Complaints", "complaints"],
  COMPLAINTS_HEADERS: [
    "complaint_id","id_maklumbalas","tarikh_terima","jenis_maklumbalas","tajuk","ringkasan_butiran","lokasi","premis_nama","tahap_kesukaran","nama_pengadu","source","status","card_status","assigned_to","assigned_at","status_updated_at","generated_card_at","due_date","created_at","created_by","kad_id",
    "appointment_letter_ref_no","appointment_letter_generated_at","appointment_letter_pdf_file_id","appointment_letter_pdf_url","appointment_letter_doc_id",
    "raw_text","parse_notes","source_pdf_file_id","source_pdf_url","extraction_temp_pdf_file_id","extraction_temp_doc_id","assigned_user_id","assigned_role",
    "report_status","report_updated_at","report_submitted_at","report_id_maklumbalas","report_jenis_aduan","report_jenis_maklumbalas_awam","report_tajuk","report_tarikh_terima","report_ringkasan_butiran","report_lokasi","report_nama_pengadu","report_tarikh_siasatan","report_pegawai_penyiasat_user_id","report_nama_pegawai","report_jawatan_pegawai","report_kategori","report_parlimen","report_jenis_premis","report_subkategori_premis_makanan","report_subkategori_produk_makanan","report_status_pensijilan","report_status_pemeriksaan_terdahulu","report_markah_pemeriksaan_semasa","report_findings_json","report_isu","report_penemuan","report_rumusan","report_kelemahan","report_kelemahan_list_json","report_kategori_penyelesaian_aduan","report_tindakan_penguatkuasaan","report_tandatangan_url","report_tandatangan_file_id","report_gambar_hadapan_premis","report_gambar_hadapan_premis_file_id","report_weakness_images_json","report_gambar_kelemahan","report_gambar_kelemahan_file_id","report_keterangan_gambar","report_generated_at","report_doc_id","report_pdf_file_id","report_pdf_url"
  ].filter((v, i, a) => a.indexOf(v) === i),
  ROLE_ALLOW_TAMBAH: ["ADMIN", "PENTADBIR SISTEM", "PENTADBIR", "PENYELARAS"],
  ROLE_ALLOW_ASSIGN: ["ADMIN", "PENTADBIR SISTEM", "PEGAWAI PENYEMAK"],
  ROLE_ALLOW_DELETE: ["ADMIN", "PENTADBIR SISTEM"],
  ALLOW_TEMP_DRIVE_EXTRACTION: true,
  TEMP_PDF_FOLDER_NAME: "SLAM_TEMP_PDF",
  DRIVE_OCR_ENABLED: true,
  DRIVE_OCR_LANG: "ms",
  GROQ_SUMMARY_ENABLED: true,
  GROQ_SUMMARY_MODEL: "llama-3.3-70b-versatile",
  GROQ_SUMMARY_FALLBACK_MODEL: "llama-3.1-8b-instant",
  GROQ_SUMMARY_MAX_INPUT_CHARS: 9000,
  GROQ_RETRY_ATTEMPTS: 3,
  GROQ_RETRY_BASE_SLEEP_MS: 900,
  CARD_DEFAULT_DUE_DAYS: 3,
  PROFILE_SIGNATURE_FOLDER_NAME: "SLAM_SIGNATURES",
  REPORT_IMAGE_FOLDER_NAME: "SLAM_REPORT_IMAGES",
  REPORT_OUTPUT_FOLDER_NAME: "SLAM_REPORT_OUTPUT",
  REPORT_DOC_TEMPLATE_ID: "1tAim-uQqNifAz9RbXsvx7qIQIe3cQcUDlOzAGs5R51o",
  APPOINTMENT_DOC_TEMPLATE_ID: "1cxWonGRCKlMxTv2LN8AXIRpuZ56JBQS54VCCovZgr2Y",
  APPOINTMENT_OUTPUT_FOLDER_NAME: "SLAM_APPOINTMENT_LETTERS",
  APPOINTMENT_REF_PREFIX: "PKK(S)100-9/3/"
};

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
