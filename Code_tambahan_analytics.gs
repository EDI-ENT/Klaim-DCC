// ═══════════════════════════════════════════════════════════════
// TAMBAHAN UNTUK Code.gs  (Apps Script Klaim Entertain)
// 2 langkah: (1) tambah blok if di doGet, (2) tempel fungsi baru
// ═══════════════════════════════════════════════════════════════

// ── LANGKAH 1 ─────────────────────────────────────────────────
// Di dalam function doGet(e), taruh blok ini tepat SEBELUM baris:
//     return jsonRes({ error: 'Action tidak dikenali' });
//
//   // ── Analytics (khusus admin/finance, pakai sheet "Admin Dashboard" yang sama) ──
//   if (action === 'getAnalytics') {
//     if (!isDashboardEmailAllowed_(e.parameter.email || '')) {
//       return jsonRes({ error: 'Tidak ada akses.' });
//     }
//     return jsonRes(getAnalyticsData_());
//   }


// ── LANGKAH 2 ─────────────────────────────────────────────────
// Tempel fungsi ini di bagian bawah Code.gs (di luar fungsi lain).

/**
 * Kirim semua baris klaim dalam bentuk ringan. Seluruh perhitungan (ranking, bulanan,
 * tahunan, MoM, dll) dikerjakan di browser, jadi filter tahun/region/kategori terasa instan.
 *
 * Kolom sheet "Klaim" (0-indexed):
 * A=0 namaUser, B=1 idUser, C=2 spesialis, D=3 namaKlinik, G=6 districtArea,
 * H=7 tanggal, I=8 bulan, J=9 kategori, K=10 tempat, N=13 nominal,
 * O=14 averageSales, Q=16 region, S=18 namaPengisi
 */
function getAnalyticsData_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const ws = ss.getSheetByName('Klaim');
  const tz = ss.getSpreadsheetTimeZone();
  const rows = ws.getDataRange().getValues().slice(1).filter(function (r) { return r[0] !== ''; });

  return rows.map(function (r) {
    // Google Sheets kadang otomatis mengubah "2026-03-15" jadi objek Date — samakan jadi teks yyyy-MM-dd
    var tgl = r[7];
    if (tgl instanceof Date) tgl = Utilities.formatDate(tgl, tz, 'yyyy-MM-dd');
    else tgl = String(tgl || '').trim();

    return {
      namaUser    : String(r[0] || '').trim(),
      idUser      : String(r[1] || '').trim(),
      spesialis   : String(r[2] || '').trim(),
      namaKlinik  : String(r[3] || '').trim(),
      districtArea: String(r[6] || '').trim(),
      tanggal     : tgl,
      bulan       : String(r[8] || '').trim(),
      kategori    : String(r[9] || '').trim(),
      tempat      : String(r[10] || '').trim(),
      nominal     : Number(r[13]) || 0,
      averageSales: Number(r[14]) || 0,
      region      : String(r[16] || '').trim() || '(Tanpa Region)',
      namaPengisi : String(r[18] || '').trim()
    };
  });
}

// SETELAH ITU: Deploy → Manage deployments → ikon pensil → Version: "New version" → Deploy.
// (URL /exec tetap sama, tidak perlu ganti API_URL.)
