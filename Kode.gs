// ============================================
// HADIRAN RT — Google Apps Script Backend v7.3 (HARDENED + FAST)
// RT 004/006 · Tanah Baru Beji · Kota Depok
// Optimasi: PropertiesService, CacheService, batch getPageData()
// + hapusAbsensi (v7.1)
// + Audit fix v7.2:
//     - Validasi input simpanAbsensi (anti-corrupt-data)
//     - bayarTalangan + editTalangan pakai LockService + rekalk saldo
//     - _parseNominal tolak nominal negatif/NaN
//     - Rate-limit login + warning password default
//     - Helper _assertSheet (gagal fast jika sheet hilang)
//     - hapusAbsensi pakai tanggal Rekap sebelum rekap dihapus
// + Performance v7.3:
//     - _batchAppend: 1 setValues() panggilan per sheet (bukan N appendRow)
//     - _batchDeleteRows: kelompokkan baris kontigu → 1 deleteRows() per blok
//     - simpanAbsensi 5-8x lebih cepat untuk 20 anggota
// ============================================

// ── Konstanta iuran ──────────────────────────────────────────────
const IURAN_TOTAL = 50000;
const IURAN_SB    = 45000;
const IURAN_KAS   = 5000;

// ── Cache TTL ────────────────────────────────────────────────────
const CACHE_TTL_SHORT  = 30;
const CACHE_TTL_MEDIUM = 120;
const CACHE_TTL_LONG   = 600;
const CACHE_TTL_HTML   = 1800;

// ── Password ─────────────────────────────────────────────────────
// CATATAN KEAMANAN: password disimpan di Script Properties.
// Jika belum di-set, fallback ke 'rt2024' (HARAP SEGERA DIUBAH via setAdminPassword()).
function _getAdminPassword() {
  try {
    const pw = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
    if (!pw) {
      Logger.log('⚠️ KEAMANAN: ADMIN_PASSWORD belum di-set. Jalankan setAdminPassword("xxxx") sekali.');
    }
    return pw || 'rt2024';
  } catch(e) {
    Logger.log('_getAdminPassword error: '+e.message);
    return 'rt2024';
  }
}

// Rate-limit login global (anti brute-force): max 5 gagal per 5 menit
function _checkLoginAllowed() {
  try {
    const props = PropertiesService.getScriptProperties();
    const now = Date.now();
    const ts  = Number(props.getProperty('_LOGIN_FAIL_TS')) || 0;
    let fails = Number(props.getProperty('_LOGIN_FAIL_COUNT')) || 0;
    if (now - ts > 5*60*1000) { fails = 0; props.setProperty('_LOGIN_FAIL_TS', String(now)); props.setProperty('_LOGIN_FAIL_COUNT','0'); }
    if (fails >= 5) return { ok:false, message:'Terlalu banyak percobaan login. Tunggu 5 menit.' };
    return { ok:true };
  } catch(e) { return { ok:true }; }
}
function _recordLoginResult(success) {
  try {
    const props = PropertiesService.getScriptProperties();
    if (success) {
      props.setProperty('_LOGIN_FAIL_COUNT', '0');
    } else {
      const cur = Number(props.getProperty('_LOGIN_FAIL_COUNT')) || 0;
      props.setProperty('_LOGIN_FAIL_COUNT', String(cur+1));
      if (!Number(props.getProperty('_LOGIN_FAIL_TS'))) props.setProperty('_LOGIN_FAIL_TS', String(Date.now()));
    }
  } catch(e) {}
}

// ── Helper Sheet ─────────────────────────────────────────────────
function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}
// Gagal-fast jika sheet hilang (cegah crash di tengah transaksi)
function _assertSheet(name) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) throw new Error('Sheet "'+name+'" tidak ditemukan. Jalankan setupSheets().');
  return s;
}
function _tglMatch(cell, str) {
  if (!cell || !str) return false;
  if (cell === str) return true;
  if (cell instanceof Date)
    return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd') === str;
  return String(cell) === str;
}
function _tglStr(cell) {
  if (!cell) return '';
  if (typeof cell === 'string') return cell.substring(0, 10);
  if (cell instanceof Date)
    return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(cell);
}

// ── Setup Sheets ─────────────────────────────────────────────────
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const defs = {
    'MasterAnggota'  : ['No','Nama','No HP','Status'],
    'JadwalPertemuan': ['NoTarikan','Tanggal','SohibulBait','Status','Timestamp'],
    'Pertemuan'      : ['ID','NoTarikan','Tanggal','Nama','NoUrut','Hadir','Bayar','Titip','Catatan','Timestamp'],
    'KasRT'          : ['ID','Tanggal','Jenis','Keterangan','Pemasukan','Pengeluaran','Saldo','Timestamp'],
    'KasHadiran'     : ['NoUrut','NoTarikan','Tanggal','SohibulBait','TotalAnggota','JumlahHadir','JumlahTidakHadir','TotalBayarMasuk','TotalSohibulBait','KasMasuk','TalanganKeluar','NetKas','CatatanTalangan','Timestamp'],
    'RekapPertemuan' : ['NoUrut','NoTarikan','Tanggal','SohibulBait','JumlahHadir','JumlahTidakHadir','TotalAnggota','TotalBayarMasuk','TotalSohibulBait','KasMasuk','TalanganKeluar','NetKas'],
    'TalanganAnggota': ['ID','NoTarikan','Tanggal','NamaAnggota','Jumlah','Status','TanggalBayar','Timestamp']
  };
  Object.entries(defs).forEach(([name, headers]) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(headers);
      sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#e8f0fe');
    }
  });
  return {success: true, message: 'Setup v7.1 selesai!'};
}

// ── doGet ─────────────────────────────────────────────────────────
function doGet() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'HTML_INDEX';
  let htmlContent = cache.get(cacheKey);
  if (!htmlContent) {
    const tmpl = HtmlService.createHtmlOutputFromFile('index');
    htmlContent = tmpl.getContent();
    try { cache.put(cacheKey, htmlContent, CACHE_TTL_HTML); } catch(e) {}
  }
  return HtmlService.createHtmlOutput(htmlContent)
    .setTitle('Hadiran RT')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, viewport-fit=cover');
}

// ── verifyPassword ────────────────────────────────────────────────
function verifyPassword(pw) {
  const gate = _checkLoginAllowed();
  if (!gate.ok) return false; // (frontend kompatibel: tetap return boolean)
  const ok = String(pw || '').trim() === _getAdminPassword();
  _recordLoginResult(ok);
  return ok;
}
// Versi baru dengan pesan eksplisit (opsional untuk dipanggil frontend)
function verifyPasswordDetailed(pw) {
  const gate = _checkLoginAllowed();
  if (!gate.ok) return { ok:false, locked:true, message: gate.message };
  const ok = String(pw || '').trim() === _getAdminPassword();
  _recordLoginResult(ok);
  return ok ? { ok:true } : { ok:false, message:'Password salah' };
}

// ── getAnggota ────────────────────────────────────────────────────
function getAnggota() {
  const cache = CacheService.getScriptCache();
  const key = 'DATA_ANGGOTA';
  const cached = cache.get(key);
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  const sh = getSheet('MasterAnggota'); if (!sh) return [];
  const data = sh.getDataRange().getValues(); if (data.length <= 1) return [];
  const result = data.slice(1).filter(r => r[1] && r[3] === 'Aktif').map((r, i) => ({
    no: i + 1, nama: String(r[1]).trim(), noHp: r[2], status: r[3]
  }));
  try { cache.put(key, JSON.stringify(result), CACHE_TTL_LONG); } catch(e) {}
  return result;
}

// ── getJadwalPertemuan ────────────────────────────────────────────
function getJadwalPertemuan() {
  const cache = CacheService.getScriptCache();
  const key = 'DATA_JADWAL';
  const cached = cache.get(key);
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  const sh = getSheet('JadwalPertemuan'); if (!sh) return [];
  const data = sh.getDataRange().getValues(); if (data.length <= 1) return [];
  const result = data.slice(1).map(r => ({
    noTarikan: Number(r[0])||0, tanggal: _tglStr(r[1]),
    sohibulBait: String(r[2]||'').trim(), status: String(r[3]||'Terjadwal').trim()
  })).filter(r => r.noTarikan > 0);
  try { cache.put(key, JSON.stringify(result), CACHE_TTL_MEDIUM); } catch(e) {}
  return result;
}

// ── getNextTarikan ────────────────────────────────────────────────
function getNextTarikan() {
  const jadwal = getJadwalPertemuan(); const rekap = getRekapPertemuan();
  const allNos = [...jadwal.map(j => j.noTarikan), ...rekap.map(r => r.noTarikan)];
  return allNos.length === 0 ? 1 : Math.max(...allNos) + 1;
}

// ── tambahJadwal ──────────────────────────────────────────────────
function tambahJadwal(data) {
  const sh = getSheet('JadwalPertemuan'); if (!sh) return {success: false};
  sh.appendRow([data.noTarikan || getNextTarikan(), data.tanggal, data.sohibulBait, 'Terjadwal', new Date().toISOString()]);
  _invalidateCache(['DATA_JADWAL', 'PAGE_dash', 'PAGE_jadwal']);
  return {success: true};
}

// ── updateJadwal ──────────────────────────────────────────────────
function updateJadwal(data) {
  const sh = getSheet('JadwalPertemuan'); if (!sh) return {success: false};
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (Number(rows[i][0]) === Number(data.noTarikan)) {
      if (data.tanggal) sh.getRange(i + 1, 2).setValue(data.tanggal);
      if (data.sohibulBait) sh.getRange(i + 1, 3).setValue(data.sohibulBait);
      _invalidateCache(['DATA_JADWAL', 'PAGE_dash', 'PAGE_jadwal']);
      return {success: true};
    }
  }
  return {success: false, message: 'Jadwal tidak ditemukan'};
}

// ── updateJadwalStatus ────────────────────────────────────────────
function updateJadwalStatus(noTarikan, status) {
  const sh = getSheet('JadwalPertemuan'); if (!sh) return;
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === Number(noTarikan)) {
      sh.getRange(i + 1, 4).setValue(status);
      _invalidateCache(['DATA_JADWAL', 'PAGE_dash', 'PAGE_jadwal']);
      return;
    }
  }
}

// ── getSaldoKas ───────────────────────────────────────────────────
function getSaldoKas() {
  const sh = getSheet('KasRT'); if (!sh) return 0;
  const last = sh.getLastRow(); if (last <= 1) return 0;
  const data = sh.getRange(2, 5, last - 1, 2).getValues();
  let saldo = 0;
  data.forEach(r => saldo += (Number(r[0])||0) - (Number(r[1])||0));
  return saldo;
}

// ── getKasHistory ─────────────────────────────────────────────────
function getKasHistory() {
  const cache = CacheService.getScriptCache();
  const key = 'DATA_KAS_HISTORY';
  const cached = cache.get(key);
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  const sh = getSheet('KasRT'); if (!sh) return [];
  const data = sh.getDataRange().getValues(); if (data.length <= 1) return [];
  const result = data.slice(1).map((r, i) => ({
    no: i+1, id: r[0], tanggal: _tglStr(r[1]), jenis: r[2],
    keterangan: r[3], pemasukan: Number(r[4])||0,
    pengeluaran: Number(r[5])||0, saldo: Number(r[6])||0
  })).reverse();
  try { cache.put(key, JSON.stringify(result), CACHE_TTL_SHORT); } catch(e) {}
  return result;
}

// ── getKasHadiran ─────────────────────────────────────────────────
function getKasHadiran() {
  const cache = CacheService.getScriptCache();
  const key = 'DATA_KAS_HADIRAN';
  const cached = cache.get(key);
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  const sh = getSheet('KasHadiran'); if (!sh) return [];
  const data = sh.getDataRange().getValues(); if (data.length <= 1) return [];
  const result = data.slice(1).map((r, i) => ({
    noUrut: i+1, noTarikan: Number(r[1])||0, tanggal: _tglStr(r[2]),
    sohibulBait: String(r[3]||'').trim(), totalAnggota: Number(r[4])||0,
    jumlahHadir: Number(r[5])||0, jumlahTidakHadir: Number(r[6])||0,
    totalBayarMasuk: Number(r[7])||0, totalSohibulBait: Number(r[8])||0,
    kasMasuk: Number(r[9])||0, talanganKeluar: Number(r[10])||0, netKas: Number(r[11])||0
  }));
  try { cache.put(key, JSON.stringify(result), CACHE_TTL_MEDIUM); } catch(e) {}
  return result;
}

// ── simpanAbsensi (PATCHED v7.2 — validasi ketat + idempotent) ─────
function simpanAbsensi(payload) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch(e) {
    return {success: false, message: 'Server sedang memproses, coba lagi sebentar.'};
  }
  try {
    // ════ VALIDASI INPUT KETAT (v7.2) ════
    if (!payload || typeof payload !== 'object')
      return {success:false, message:'Payload kosong'};
    const noTarikan   = Number(payload.noTarikan)||0;
    const tanggal     = String(payload.tanggal||'').trim();
    const sohibulBait = String(payload.sohibulBait||'').trim();
    const absensi     = payload.absensi;
    const ts          = new Date().toISOString();

    if (!noTarikan || noTarikan < 1)
      return {success:false, message:'NoTarikan tidak valid (harus > 0)'};
    if (!/^\d{4}-\d{2}-\d{2}$/.test(tanggal))
      return {success:false, message:'Tanggal harus format YYYY-MM-DD'};
    if (!sohibulBait)
      return {success:false, message:'Sohibul Bait wajib diisi'};
    if (!Array.isArray(absensi) || absensi.length === 0)
      return {success:false, message:'Daftar absensi kosong'};

    // Cek setiap anggota punya field nama
    for (let i = 0; i < absensi.length; i++) {
      const a = absensi[i];
      if (!a || typeof a !== 'object' || !String(a.nama||'').trim())
        return {success:false, message:'Baris absensi #'+(i+1)+' tidak valid (nama kosong)'};
    }

    // Cek sohibulBait ada di daftar absensi (cegah hitungan SB ngawur)
    const namaSet = absensi.map(a => String(a.nama).trim().toLowerCase());
    if (namaSet.indexOf(sohibulBait.toLowerCase()) < 0)
      return {success:false, message:'Sohibul Bait "'+sohibulBait+'" tidak ada di daftar absensi'};

    // Cek nama anggota valid (status Aktif di master)
    const aktifSet = getAnggota().map(x => String(x.nama).trim().toLowerCase());
    const tidakValid = absensi.filter(a => aktifSet.indexOf(String(a.nama).trim().toLowerCase()) < 0);
    if (tidakValid.length > 0)
      return {success:false, message:'Anggota tidak terdaftar/non-aktif: '+tidakValid.map(a=>a.nama).join(', ')};

    let shP, shK, shKH, shR, shT;
    try {
      shP  = _assertSheet('Pertemuan');
      shK  = _assertSheet('KasRT');
      shKH = _assertSheet('KasHadiran');
      shR  = _assertSheet('RekapPertemuan');
      shT  = _assertSheet('TalanganAnggota');
    } catch(e) { return {success:false, message:e.message}; }

    // Idempotent: kalau noTarikan sudah ada, hapus dulu lalu tulis ulang
    _hapusRowsByNoTarikan(shP, 1, noTarikan);
    _hapusRowsByNoTarikan(shKH, 1, noTarikan);
    _hapusRowsByNoTarikan(shR, 1, noTarikan);
    _hapusKasByTanggalJenis(shK, tanggal, 'Iuran Mingguan');
    _hapusKasByTanggalJenis(shK, tanggal, 'Talangan Sohibul Bait');

    const rekapCount       = shR.getLastRow();
    const totalAnggota     = absensi.length;
    const hadirList        = absensi.filter(a => a.hadir || a.titip);
    const tidakHadirList   = absensi.filter(a => !a.hadir && !a.titip);
    const jumlahHadir      = hadirList.length;
    const jumlahTidakHadir = tidakHadirList.length;
    const totalBayarMasuk  = jumlahHadir * IURAN_TOTAL;
    const sbDiterima       = (totalAnggota - 1) * IURAN_SB;
    const kasDariHadir     = (jumlahHadir - 1) * IURAN_KAS;
    const talanganKeluar   = jumlahTidakHadir * IURAN_TOTAL;
    const kasDariTalangan  = jumlahTidakHadir * IURAN_KAS;
    const totalKasMasuk    = kasDariHadir + kasDariTalangan;
    const netKas           = totalKasMasuk - talanganKeluar;

    // ════ BATCH WRITE v7.2 (1 round-trip per sheet, bukan N) ════
    // Pertemuan: tulis semua anggota sekaligus
    const pertemuanRows = absensi.map((a, i) => [
      tanggal+'_'+(i+1), noTarikan, tanggal, a.nama, (i+1),
      a.hadir?'Ya':'Tidak', a.bayar?'Ya':'Tidak', a.titip?'Ya':'Tidak',
      a.catatan||'-', ts
    ]);
    _batchAppend(shP, pertemuanRows);

    // KasRT: 1-2 baris (iuran + opsional talangan keluar)
    const kasRows = [[
      'KAS_IN_'+tanggal, tanggal, 'Iuran Mingguan',
      'Iuran '+(jumlahHadir-1)+' hadir (excl.SB) + '+jumlahTidakHadir+' talangan → SB: '+sohibulBait + ' [tarikan-'+noTarikan+']',
      totalKasMasuk, 0, 0, ts
    ]];
    if (talanganKeluar > 0) {
      kasRows.push([
        'KAS_OUT_'+tanggal, tanggal, 'Talangan Sohibul Bait',
        'Talangan '+jumlahTidakHadir+' anggota (50k/org) → '+sohibulBait + ' [tarikan-'+noTarikan+']',
        0, talanganKeluar, 0, ts
      ]);
    }
    _batchAppend(shK, kasRows);

    // KasHadiran + RekapPertemuan: 1 baris masing-masing
    _batchAppend(shKH, [[rekapCount, noTarikan, tanggal, sohibulBait, totalAnggota,
      jumlahHadir, jumlahTidakHadir, totalBayarMasuk, sbDiterima,
      totalKasMasuk, talanganKeluar, netKas, '', ts]]);
    _batchAppend(shR, [[rekapCount, noTarikan, tanggal, sohibulBait, jumlahHadir,
      jumlahTidakHadir, totalAnggota, totalBayarMasuk, sbDiterima,
      totalKasMasuk, talanganKeluar, netKas]]);

    // Talangan: hapus dulu (batch) lalu tulis semua sekaligus
    _hapusTalanganByTarikan(shT, noTarikan);
    if (tidakHadirList.length > 0) {
      const talanganRows = tidakHadirList.map(a => [
        'TAL_'+noTarikan+'_'+a.nama, noTarikan, tanggal, a.nama,
        IURAN_TOTAL, 'Belum Lunas', '', ts
      ]);
      _batchAppend(shT, talanganRows);
    }

    if (noTarikan > 0) updateJadwalStatus(noTarikan, 'Selesai');

    // ✅ PATCH: rekalkulasi saldo berjalan KasRT setelah semua baris ditulis
    _rekalkSaldoOtomatis();

    // Ambil saldo akhir yang sudah benar setelah rekalkulasi
    const saldoAkhir = getSaldoKas();

    _invalidateCache(['DATA_KAS_HISTORY','DATA_KAS_HADIRAN',
      'PAGE_dash','PAGE_absensi','PAGE_kas','PAGE_talangan']);

    return {
      success: true, noTarikan, totalAnggota, jumlahHadir, jumlahTidakHadir,
      totalBayarMasuk, sbDiterima, kasMasuk: totalKasMasuk,
      talanganKeluar, netKas, saldoAkhir: saldoAkhir
    };
  } finally { lock.releaseLock(); }
}

// ══════════════════════════════════════════════════════════════════
//  hapusAbsensi — hapus semua data tarikan by noTarikan (v7.1 NEW)
// ══════════════════════════════════════════════════════════════════
function hapusAbsensi(payload) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e) {
    return { success: false, message: 'Server sedang memproses, coba lagi.' };
  }
  try {
    const noTarikan = Number(payload && payload.noTarikan) || 0;
    if (!noTarikan) return { success: false, message: 'NoTarikan tidak valid' };

    const shP  = getSheet('Pertemuan');
    const shKH = getSheet('KasHadiran');
    const shR  = getSheet('RekapPertemuan');
    const shT  = getSheet('TalanganAnggota');
    const shK  = getSheet('KasRT');

    // ✅ FIX v7.2: AMBIL TANGGAL DULU sebelum rekap dihapus (utk fallback KasRT cleanup)
    let tanggalTarikan = '';
    if (shR) {
      const rData = shR.getDataRange().getValues();
      for (let i = 1; i < rData.length; i++) {
        if (Number(rData[i][1]) === noTarikan) {
          tanggalTarikan = _tglStr(rData[i][2]);
          break;
        }
      }
    }

    if (shP)  _hapusRowsByNoTarikan(shP, 1, noTarikan);
    if (shKH) _hapusRowsByNoTarikan(shKH, 1, noTarikan);
    if (shR)  _hapusRowsByNoTarikan(shR, 1, noTarikan);
    if (shT)  _hapusTalanganByTarikan(shT, noTarikan);

    // Hapus dari KasRT: Iuran + Talangan tarikan ini
    if (shK) {
      const data = shK.getDataRange().getValues();
      const toDelete = [];
      for (let i = data.length - 1; i >= 1; i--) {
        const jenis = String(data[i][2] || '');
        const ket   = String(data[i][3] || '').toLowerCase();
        const cocokKet = (jenis === 'Iuran Mingguan' || jenis === 'Talangan Sohibul Bait') &&
            (ket.indexOf('tarikan-' + noTarikan) >= 0 ||
             ket.indexOf('#' + noTarikan) >= 0 ||
             ket.indexOf('ke-' + noTarikan) >= 0 ||
             ket.indexOf('tarikan ke ' + noTarikan) >= 0);
        // ✅ FIX v7.2: fallback hapus by tanggal jika keterangan tidak punya noTarikan
        const cocokTgl = tanggalTarikan && (jenis === 'Iuran Mingguan' || jenis === 'Talangan Sohibul Bait')
                         && _tglMatch(data[i][1], tanggalTarikan);
        if (cocokKet || cocokTgl) toDelete.push(i + 1);
      }
      if (!toDelete.length) {
        Logger.log('hapusAbsensi: tidak ada baris KasRT yang cocok untuk noTarikan=' + noTarikan + ' tgl=' + tanggalTarikan);
      }
      toDelete.sort((a, b) => b - a).forEach(r => shK.deleteRow(r));
      // Rekalkulasi saldo setelah hapus
      if (toDelete.length) _rekalkSaldoOtomatis();
    }

    // Kembalikan status jadwal → Terjadwal
    updateJadwalStatus(noTarikan, 'Terjadwal');

    SpreadsheetApp.flush();
    _invalidateCache(['DATA_KAS_HISTORY','DATA_KAS_HADIRAN','DATA_REKAP',
      'PAGE_dash','PAGE_absensi','PAGE_kas','PAGE_talangan']);

    return { success: true, noTarikan: noTarikan, hapusBarisKas: 0 };
  } catch(e) {
    Logger.log('hapusAbsensi error: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── getTalanganList ───────────────────────────────────────────────
function getTalanganList() {
  const cache = CacheService.getScriptCache();
  const key = 'DATA_TALANGAN';
  const cached = cache.get(key);
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  const sh = getSheet('TalanganAnggota'); if (!sh) return [];
  const data = sh.getDataRange().getValues(); if (data.length <= 1) return [];
  const result = data.slice(1).map(r => ({
    id: r[0], noTarikan: Number(r[1])||0, tanggal: _tglStr(r[2]),
    namaAnggota: String(r[3]).trim(), jumlah: Number(r[4])||0,
    status: String(r[5]).trim(), tanggalBayar: _tglStr(r[6])
  })).reverse();
  try { cache.put(key, JSON.stringify(result), CACHE_TTL_SHORT); } catch(e) {}
  return result;
}

// ── bayarTalangan (v7.2 — Lock + rekalk saldo, anti race) ────────
function bayarTalangan(payload) {
  // Validasi input
  if (!payload || !payload.id)         return {success:false, message:'ID talangan kosong'};
  if (!payload.tanggalBayar)           return {success:false, message:'Tanggal bayar wajib diisi'};
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(payload.tanggalBayar)))
                                       return {success:false, message:'Tanggal format YYYY-MM-DD'};

  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e) {
    return {success:false, message:'Server sibuk, coba lagi 5 detik.'};
  }
  try {
    const sh = getSheet('TalanganAnggota'), shKas = getSheet('KasRT');
    if (!sh||!shKas) return {success: false, message: 'Sheet tidak ada'};
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === payload.id && data[i][5] === 'Belum Lunas') {
        const jumlah = Number(data[i][4])||0, nama = data[i][3];
        if (jumlah <= 0) return {success:false, message:'Jumlah talangan tidak valid'};
        sh.getRange(i + 1, 6).setValue('Lunas');
        sh.getRange(i + 1, 7).setValue(payload.tanggalBayar);
        // Tulis dengan saldo sementara 0 — saldo final dihitung ulang setelah append
        shKas.appendRow([
          'KAS_IN_TAL_'+Date.now(), payload.tanggalBayar,
          'Pembayaran Talangan', 'Talangan lunas oleh ' + nama,
          jumlah, 0, 0, new Date().toISOString()
        ]);
        _rekalkSaldoOtomatis();  // ← FIX: hindari TOCTOU saldo
        const saldoBaru = getSaldoKas();
        _invalidateCache(['DATA_TALANGAN','DATA_KAS_HISTORY','PAGE_talangan','PAGE_kas','PAGE_dash']);
        return {success: true, saldoBaru: saldoBaru};
      }
    }
    return {success: false, message: 'Talangan tidak ditemukan atau sudah lunas'};
  } finally { lock.releaseLock(); }
}

// ── editTalangan (v7.2 — Lock + rekalk saldo, anti race) ─────────
function editTalangan(params) {
  if (!params || typeof params !== 'object') {
    return { success: false, message: 'Parameter tidak valid' };
  }
  var id           = String(params.id || '').trim();
  var statusBaru   = String(params.status || '').trim();
  var tanggalBayar = String(params.tanggalBayar || '').trim();
  var jumlahBaru   = Number(params.jumlah) || 0;

  if (!id)         return { success: false, message: 'ID talangan kosong' };
  if (!statusBaru) return { success: false, message: 'Status kosong' };
  if (statusBaru === 'Lunas' && !tanggalBayar)
                   return { success: false, message: 'Tanggal bayar wajib diisi jika Lunas' };
  if (statusBaru === 'Lunas' && !/^\d{4}-\d{2}-\d{2}$/.test(tanggalBayar))
                   return { success: false, message: 'Tanggal harus format YYYY-MM-DD' };
  if (jumlahBaru <= 0) return { success: false, message: 'Jumlah harus > 0' };
  if (['Lunas','Belum Lunas'].indexOf(statusBaru) < 0)
                   return { success: false, message: 'Status tidak valid' };

  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e) {
    return { success:false, message:'Server sibuk, coba lagi 5 detik.' };
  }
  try {
    var sheet = getSheet('TalanganAnggota');
    if (!sheet) return { success: false, message: 'Sheet TalanganAnggota tidak ditemukan' };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'Belum ada data talangan' };

    var allData = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    var rowFound = -1;
    for (var i = 0; i < allData.length; i++) {
      if (String(allData[i][0] || '').trim() === id) { rowFound = i + 2; break; }
    }
    if (rowFound === -1) return { success: false, message: 'Talangan "' + id + '" tidak ditemukan' };

    // CATATAN INDEX KOLOM (header: ID, NoTarikan, Tanggal, NamaAnggota, Jumlah, Status, TanggalBayar, Timestamp)
    var rowData     = allData[rowFound - 2];
    var noTarikan   = Number(rowData[1] || 0);          // col B = NoTarikan
    var namaAnggota = String(rowData[3] || '').trim();  // col D = NamaAnggota
    var oldJumlah   = Number(rowData[4] || 0);          // col E = Jumlah
    var oldStatus   = String(rowData[5] || '').trim();  // col F = Status

    sheet.getRange(rowFound, 6).setValue(statusBaru);   // col F = Status
    sheet.getRange(rowFound, 5).setValue(jumlahBaru);   // col E = Jumlah

    if (statusBaru === 'Lunas' && tanggalBayar) {
      var p = tanggalBayar.split('-');
      sheet.getRange(rowFound, 7).setValue(p.length === 3 ? new Date(+p[0], +p[1]-1, +p[2]) : tanggalBayar);
    } else if (statusBaru === 'Belum Lunas') {
      sheet.getRange(rowFound, 7).clearContent();
    }

    var shKas = getSheet('KasRT');
    if (shKas) {
      var ts = new Date().toISOString();
      var tglVal = tanggalBayar || _tglStr(new Date());
      // Tulis dengan saldo 0 sementara — biarkan _rekalkSaldoOtomatis() menghitung
      if (statusBaru === 'Lunas' && oldStatus !== 'Lunas') {
        shKas.appendRow(['KAS_IN_TAL_EDIT_'+Date.now(), tglVal, 'Pembayaran Talangan',
          'Talangan lunas (edit) — '+namaAnggota+' Tarikan #'+noTarikan, jumlahBaru, 0, 0, ts]);
      } else if (statusBaru === 'Belum Lunas' && oldStatus === 'Lunas') {
        shKas.appendRow(['KAS_REV_TAL_'+Date.now(), _tglStr(new Date()), 'Koreksi Talangan',
          'Batalkan pembayaran talangan — '+namaAnggota+' Tarikan #'+noTarikan, 0, oldJumlah, 0, ts]);
      }
      _rekalkSaldoOtomatis();
    }

    SpreadsheetApp.flush();
    _invalidateCache(['DATA_TALANGAN','DATA_KAS_HISTORY','PAGE_talangan','PAGE_kas','PAGE_dash']);
    return { success: true, message: 'Talangan diperbarui', statusBaru: statusBaru };
  } catch(e) {
    Logger.log('editTalangan error: ' + e.message);
    return { success: false, message: 'Error: ' + e.message };
  } finally { lock.releaseLock(); }
}

// ── getRekapPertemuan ─────────────────────────────────────────────
function getRekapPertemuan() {
  const cache = CacheService.getScriptCache();
  const key = 'DATA_REKAP';
  const cached = cache.get(key);
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  const sh = getSheet('RekapPertemuan'); if (!sh) return [];
  const data = sh.getDataRange().getValues(); if (data.length <= 1) return [];
  const result = data.slice(1).map((r, i) => ({
    noUrut: i+1, noTarikan: Number(r[1])||0, tanggal: _tglStr(r[2]),
    sohibulBait: String(r[3]||'').trim(), jumlahHadir: Number(r[4])||0,
    jumlahTidakHadir: Number(r[5])||0, totalAnggota: Number(r[6])||0,
    totalBayarMasuk: Number(r[7])||0, totalSohibulBait: Number(r[8])||0,
    kasMasuk: Number(r[9])||0, talanganKeluar: Number(r[10])||0, netKas: Number(r[11])||0
  }));
  try { cache.put(key, JSON.stringify(result), CACHE_TTL_MEDIUM); } catch(e) {}
  return result;
}

// ── getDashboardData ──────────────────────────────────────────────
function getDashboardData() {
  try {
    const anggota = getAnggota(), rekap = getRekapPertemuan(),
          kasHistory = getKasHistory(), jadwal = getJadwalPertemuan(),
          talangan = getTalanganList(), saldo = getSaldoKas();
    const totalTerkumpul   = kasHistory.filter(k => k.jenis === 'Iuran Mingguan').reduce((s,k) => s+k.pemasukan, 0);
    const totalPengeluaran = kasHistory.filter(k => k.jenis === 'Pengeluaran').reduce((s,k) => s+k.pengeluaran, 0);
    const sisaTalangan     = talangan.filter(t => t.status === 'Belum Lunas').reduce((s,t) => s+t.jumlah, 0);
    return {
      totalAnggota: anggota.length, totalPertemuan: rekap.length,
      saldoKas: saldo, totalTerkumpul, totalPengeluaran, sisaTalangan,
      pertemuanTerakhir: rekap[0] || null,
      recentKas: kasHistory.slice(0, 5),
      jadwalBerikutnya: jadwal.filter(j => j.status === 'Terjadwal').sort((a,b) => a.noTarikan-b.noTarikan),
      nextTarikan: getNextTarikan()
    };
  } catch(e) {
    return {
      totalAnggota:0, totalPertemuan:0, saldoKas:0, totalTerkumpul:0,
      totalPengeluaran:0, sisaTalangan:0, pertemuanTerakhir:null,
      recentKas:[], jadwalBerikutnya:[], nextTarikan:1, error: e.message
    };
  }
}

// ── getRincianTarikan ─────────────────────────────────────────────
function getRincianTarikan(noTarikan) {
  const shP = getSheet('Pertemuan'), shR = getSheet('RekapPertemuan');
  if (!shP || !shR) return {success: false, message: 'Sheet tidak lengkap'};
  const rData = shR.getDataRange().getValues();
  let sb = '', tgl = '';
  for (let i = 1; i < rData.length; i++) {
    if (Number(rData[i][1]) === Number(noTarikan)) {
      tgl = _tglStr(rData[i][2]); sb = String(rData[i][3] || '').trim(); break;
    }
  }
  const pData = shP.getDataRange().getValues();
  let anggota = [], sbFound = false;
  for (let i = 1; i < pData.length; i++) {
    if (Number(pData[i][1]) === Number(noTarikan)) {
      const nama = String(pData[i][3]).trim();
      const isSB = sb && nama.toLowerCase() === sb.toLowerCase();
      if (isSB) sbFound = true;
      anggota.push({
        nama, hadir: String(pData[i][5]).trim() === 'Ya',
        bayar: String(pData[i][6]).trim() === 'Ya',
        titip: String(pData[i][7]).trim() === 'Ya',
        catatan: String(pData[i][8] || '').trim(), isSB
      });
    }
  }
  if (!sbFound && sb) {
    anggota.unshift({ nama: sb, hadir: true, bayar: false, titip: false, catatan: 'Sohibul Bait', isSB: true });
  }
  anggota.sort((a, b) => {
    if (a.isSB && !b.isSB) return -1;
    if (!a.isSB && b.isSB) return 1;
    return a.nama.localeCompare(b.nama);
  });
  return {success: true, noTarikan, tanggal: tgl, sohibulBait: sb, anggota};
}

// ── Helper internal ───────────────────────────────────────────────
// _batchAppend: append banyak baris sekaligus dengan 1 round-trip
// (jauh lebih cepat dari N kali appendRow ketika N>3)
function _batchAppend(sheet, rows) {
  if (!sheet || !rows || rows.length === 0) return;
  const startRow = Math.max(2, sheet.getLastRow() + 1);
  const numCols  = rows[0].length;
  sheet.getRange(startRow, 1, rows.length, numCols).setValues(rows);
}

// _batchDeleteRows: hapus banyak baris dengan deteksi range kontigu
// (hasil identik dengan deleteRow per baris, tapi 1 panggilan per blok kontigu)
function _batchDeleteRows(sheet, rowNumbers) {
  if (!sheet || !rowNumbers || rowNumbers.length === 0) return 0;
  // Urutkan dari terbesar → terkecil supaya hapus dari bawah ke atas (index tetap valid)
  const sorted = rowNumbers.slice().sort((a, b) => b - a);
  let deleted = 0;
  let i = 0;
  while (i < sorted.length) {
    let runEnd = i;
    // Cari run kontigu menurun: sorted[i], sorted[i]-1, sorted[i]-2, ...
    while (runEnd + 1 < sorted.length && sorted[runEnd + 1] === sorted[runEnd] - 1) runEnd++;
    const top = sorted[runEnd];       // baris paling atas dari run
    const len = runEnd - i + 1;       // jumlah baris kontigu
    sheet.deleteRows(top, len);
    deleted += len;
    i = runEnd + 1;
  }
  return deleted;
}

function _hapusRowsByNoTarikan(sheet, colNoTarikan, noTarikan) {
  const data = sheet.getDataRange().getValues();
  const rowsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (Number(data[i][colNoTarikan]) === Number(noTarikan)) rowsToDelete.push(i + 1);
  }
  return _batchDeleteRows(sheet, rowsToDelete);
}
function _hapusKasByTanggalJenis(sheet, tanggal, jenis) {
  const data = sheet.getDataRange().getValues();
  const rowsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (_tglMatch(data[i][1], tanggal) && data[i][2] === jenis) rowsToDelete.push(i + 1);
  }
  return _batchDeleteRows(sheet, rowsToDelete);
}
function _hapusTalanganByTarikan(sheet, noTarikan) {
  const data = sheet.getDataRange().getValues();
  const rowsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (Number(data[i][1]) === Number(noTarikan)) rowsToDelete.push(i + 1);
  }
  return _batchDeleteRows(sheet, rowsToDelete);
}

// ── Cross-spreadsheet: Laporan Kas RT ────────────────────────────
var LAPORAN_KAS_ID = '1wlu_bS_CorBIwidL6pb4LJYA5TbkW4W2-duxqXzbmk4';
var SH_PEMASUKAN   = 'PEMASUKAN';
var SH_PENGELUARAN = 'PENGELUARAN';

function _openLaporanKas() {
  try { return SpreadsheetApp.openById(LAPORAN_KAS_ID); }
  catch(e) { Logger.log('openLaporanKas error: '+e.message); return null; }
}
function _findSheetByHeader(ss, hint) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var hdr = String(sheets[i].getRange(1,1,1,sheets[i].getLastColumn()||1).getValues()[0]||'').toLowerCase();
    if (hdr.indexOf(hint.toLowerCase()) >= 0) return sheets[i];
  }
  return null;
}

function getLaporanKasBesar() {
  try {
    var ss = _openLaporanKas();
    if (!ss) return { error: 'Tidak bisa akses spreadsheet Laporan Kas RT.', records: [], summary: {} };
    var shM = ss.getSheetByName(SH_PEMASUKAN)   || _findSheetByHeader(ss, 'Keterangan');
    var shK = ss.getSheetByName(SH_PENGELUARAN) || _findSheetByHeader(ss, 'Keterangan Detail');
    var records = [], totMasuk = 0, totKeluar = 0;
    if (shM) {
      var dM = shM.getDataRange().getValues();
      for (var i = 1; i < dM.length; i++) {
        var r = dM[i]; if (!r[0]) continue;
        var nominal = _parseNominal(r[4]);
        var kategori = String(r[2]||'').toUpperCase();
        if (kategori !== 'SALDO_AWAL') totMasuk += nominal;
        records.push({ id: String(r[0]||''), tanggal: _fmtTglID(r[1]), kategori,
          keterangan: String(r[3]||''), nominal, jenis: 'masuk' });
      }
    }
    if (shK) {
      var dK = shK.getDataRange().getValues();
      for (var j = 1; j < dK.length; j++) {
        var rk = dK[j]; if (!rk[0]) continue;
        var nom = _parseNominal(rk[4]);
        totKeluar += nom;
        records.push({ id: String(rk[0]||''), tanggal: _fmtTglID(rk[1]),
          kategori: String(rk[2]||'').toUpperCase(),
          keterangan: String(rk[3]||''), nominal: nom, jenis: 'keluar' });
      }
    }
    records.sort(function(a,b){ return String(b.tanggal).localeCompare(String(a.tanggal)); });
    return { records, summary: { totalMasuk: totMasuk, totalKeluar: totKeluar, saldoBersih: totMasuk-totKeluar }, success: true };
  } catch(e) { return { error: e.message, records: [], summary: {} }; }
}

function _parseNominal(v) {
  if (typeof v === 'number') {
    if (!isFinite(v) || v < 0) return 0;
    return v;
  }
  // Buang semua non-digit (kecuali minus di depan), lalu pisah ribuan/desimal
  var raw = String(v||'').trim();
  if (!raw) return 0;
  var neg = /^-/.test(raw);
  var s = raw.replace(/[^0-9]/g, '');
  var n = Number(s);
  if (!isFinite(n) || isNaN(n)) return 0;
  if (neg) return 0; // tolak negatif (cegah manipulasi finansial)
  return n;
}
function _fmtTglID(d) {
  if (!d) return '';
  if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var s = String(d), m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (m) return m[3]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[1]).slice(-2);
  return s.substring(0,10);
}

function setorKasBesar(data) {
  try {
    var ss = _openLaporanKas();
    if (!ss) return { success: false, message: 'Tidak bisa akses spreadsheet Laporan Kas RT.' };
    var sh = ss.getSheetByName(SH_PEMASUKAN) || _findSheetByHeader(ss, 'Keterangan');
    if (!sh) return { success: false, message: 'Sheet PEMASUKAN tidak ditemukan.' };
    var p = String(data.tanggal||'').split('-');
    var d = p.length===3 ? new Date(+p[0], +p[1]-1, +p[2]) : new Date();
    var id = 'PMS-HDR-' + Utilities.formatDate(new Date(), 'GMT+7', 'yyMMddHHmmss');
    var nominal = Number(data.jumlah)||0;
    sh.appendRow([id, d, 'IURAN_ANGGOTA', String(data.keterangan||''), nominal, d.getMonth()+1, d.getFullYear(), new Date()]);
    var shKas = getSheet('KasRT');
    if (shKas) {
      var saldoBaru = getSaldoKas() - nominal;
      shKas.appendRow([id, data.tanggal, 'Setor Kas Besar RT', String(data.keterangan||''), 0, nominal, saldoBaru, new Date().toISOString()]);
    }
    _invalidateCache(['DATA_KAS_HISTORY','PAGE_kas','PAGE_kasrt','PAGE_dash']);
    return { success: true, id };
  } catch(e) { return { success: false, message: e.message }; }
}

function updateSetoranKasBesar(data) {
  try {
    var ss = _openLaporanKas(); if (!ss) return { success: false, message: 'Tidak bisa akses spreadsheet.' };
    var sh = data.jenis==='keluar'
      ? (ss.getSheetByName(SH_PENGELUARAN) || _findSheetByHeader(ss,'Keterangan Detail'))
      : (ss.getSheetByName(SH_PEMASUKAN)   || _findSheetByHeader(ss,'Keterangan'));
    if (!sh) return { success: false, message: 'Sheet tidak ditemukan.' };
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        var p = String(data.tanggal||'').split('-');
        var d = p.length===3 ? new Date(+p[0], +p[1]-1, +p[2]) : new Date();
        sh.getRange(i+1, 2).setValue(d);
        sh.getRange(i+1, 4).setValue(data.keterangan);
        sh.getRange(i+1, 5).setValue(Number(data.nominal)||0);
        sh.getRange(i+1, 6).setValue(d.getMonth()+1);
        sh.getRange(i+1, 7).setValue(d.getFullYear());
        _invalidateCache(['PAGE_kasrt']);
        return { success: true };
      }
    }
    return { success: false, message: 'ID '+data.id+' tidak ditemukan' };
  } catch(e) { return { success: false, message: e.message }; }
}

function hapusSetoranKasBesar(data) {
  try {
    var ss = _openLaporanKas(); if (!ss) return { success: false, message: 'Tidak bisa akses spreadsheet.' };
    var sh = data.jenis==='keluar'
      ? (ss.getSheetByName(SH_PENGELUARAN) || _findSheetByHeader(ss,'Keterangan Detail'))
      : (ss.getSheetByName(SH_PEMASUKAN)   || _findSheetByHeader(ss,'Keterangan'));
    if (!sh) return { success: false, message: 'Sheet tidak ditemukan.' };
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        sh.deleteRow(i+1);
        var shKas = getSheet('KasRT');
        if (shKas) {
          var rk = shKas.getDataRange().getValues();
          for (var j = rk.length-1; j >= 1; j--) {
            if (String(rk[j][0]) === String(data.id)) { shKas.deleteRow(j+1); break; }
          }
        }
        _invalidateCache(['DATA_KAS_HISTORY','PAGE_kasrt','PAGE_kas','PAGE_dash']);
        return { success: true };
      }
    }
    return { success: false, message: 'ID '+data.id+' tidak ditemukan' };
  } catch(e) { return { success: false, message: e.message }; }
}

function tambahPemasukanKasBesar(params) {
  try {
    var kat = String(params.kategori||'LAIN_LAIN').toUpperCase();
    var tgl = String(params.tanggal||'');
    var ket = String(params.keterangan||'').trim();
    var jml = Number(params.jumlah)||0;
    if (!tgl||!ket||!jml) return {success:false, message:'Lengkapi semua field'};
    var ss = _openLaporanKas(); if (!ss) return {success:false, message:'Tidak bisa akses Laporan Kas RT'};
    var sh = ss.getSheetByName(SH_PEMASUKAN); if (!sh) return {success:false, message:'Sheet PEMASUKAN tidak ditemukan'};
    var p = tgl.split('-'), d = new Date(+p[0],+p[1]-1,+p[2]);
    var kode = 'KMS-' + new Date().getTime().toString(36).toUpperCase();
    sh.appendRow([kode, d, kat, ket, jml, d.getMonth()+1, d.getFullYear(), new Date()]);
    _invalidateCache(['PAGE_kasrt']);
    return {success:true, message:'Pemasukan tercatat', id:kode};
  } catch(e) { return {success:false, message:'Error: '+e.message}; }
}

function tambahPengeluaranKasBesar(params) {
  try {
    var kat = String(params.kategori||'LAIN_LAIN').toUpperCase();
    var tgl = String(params.tanggal||'');
    var ket = String(params.keterangan||'').trim();
    var jml = Number(params.jumlah)||0;
    if (!tgl||!ket||!jml) return {success:false, message:'Lengkapi semua field'};
    var ss = _openLaporanKas(); if (!ss) return {success:false, message:'Tidak bisa akses Laporan Kas RT'};
    var sh = ss.getSheetByName(SH_PENGELUARAN); if (!sh) return {success:false, message:'Sheet PENGELUARAN tidak ditemukan'};
    var p = tgl.split('-'), d = new Date(+p[0],+p[1]-1,+p[2]);
    var kode = 'KLR-' + new Date().getTime().toString(36).toUpperCase();
    sh.appendRow([kode, d, kat, ket, jml, d.getMonth()+1, d.getFullYear(), new Date()]);
    _invalidateCache(['PAGE_kasrt']);
    return {success:true, message:'Pengeluaran tercatat', id:kode};
  } catch(e) { return {success:false, message:'Error: '+e.message}; }
}

// ── tambahAnggota ─────────────────────────────────────────────────
function tambahAnggota(data) {
  try {
    var sh = getSheet('MasterAnggota'); if (!sh) return {success:false, message:'Sheet tidak ada'};
    var d = sh.getDataRange().getValues();
    var maxNo = 0;
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][1]||'').trim().toLowerCase()===String(data.nama||'').trim().toLowerCase())
        return {success:false, message:'Nama "'+data.nama+'" sudah ada'};
      if ((Number(d[i][0])||0)>maxNo) maxNo=Number(d[i][0])||0;
    }
    sh.appendRow([maxNo+1, String(data.nama).trim(), String(data.noHp||''), data.status||'Aktif']);
    _invalidateCache(['DATA_ANGGOTA','PAGE_absensi','PAGE_dash']);
    return {success:true, no:maxNo+1, nama:data.nama};
  } catch(e) { return {success:false, message:e.message}; }
}

function refreshAnggota() {
  _invalidateCache(['DATA_ANGGOTA','PAGE_absensi','PAGE_dash']);
  return getAnggota();
}

// ── Batch getPageData ─────────────────────────────────────────────
function getPageData(page) {
  const cache = CacheService.getScriptCache();
  const key   = 'PAGE_' + String(page || 'dash');
  const cached = cache.get(key);
  if (cached) {
    try { const p = JSON.parse(cached); p._fromCache = true; return p; } catch(e) {}
  }
  let result = {}, ttl = CACHE_TTL_SHORT;
  try {
    switch (String(page)) {
      case 'dash':     result = getDashboardData(); ttl = CACHE_TTL_SHORT; break;
      case 'talangan': result = { talangan: getTalanganList(), kasHadiran: getKasHadiran() }; ttl = CACHE_TTL_SHORT; break;
      case 'kas':      result = { kasHistory: getKasHistory(), kasHadiran: getKasHadiran(), saldo: getSaldoKas(), talangan: getTalanganList() }; ttl = CACHE_TTL_SHORT; break;
      case 'kasrt':    result = getLaporanKasBesar(); ttl = CACHE_TTL_MEDIUM; break;
      case 'absensi':  result = { anggota: getAnggota(), jadwal: getJadwalPertemuan(), rekap: getRekapPertemuan() }; ttl = CACHE_TTL_MEDIUM; break;
      default:         result = getDashboardData(); ttl = CACHE_TTL_SHORT;
    }
    result._fromCache = false;
    try { cache.put(key, JSON.stringify(result), ttl); } catch(e) {}
  } catch(err) { result = { error: err.message, _fromCache: false }; }
  return result;
}

// ── _invalidateCache ──────────────────────────────────────────────
function _invalidateCache(keys) {
  try {
    const cache = CacheService.getScriptCache();
    if (Array.isArray(keys)) cache.removeAll(keys);
    else cache.remove(keys);
  } catch(e) { Logger.log('_invalidateCache error: '+e.message); }
}

// ── setAdminPassword ──────────────────────────────────────────────
function setAdminPassword(newPw) {
  if (!newPw || newPw.length < 4) return {success:false, message:'Password minimal 4 karakter'};
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', String(newPw));
  return {success:true, message:'Password berhasil disimpan'};
}
function hapusTarikan(data) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch(e) { return { success:false, message:'Server sedang sibuk, coba lagi 5 detik' }; }
 
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const noTar = Number(data && data.noTarikan);
    if (!noTar || noTar < 1) {
      return { success:false, message:'Nomor tarikan tidak valid' };
    }
 
    const log = [];
 
    /* ── STEP 1: Baca info tarikan SEBELUM dihapus (untuk match KasRT) ── */
    let tanggalTar = '', sohibulBait = '';
    const shR = ss.getSheetByName('RekapPertemuan');
    if (shR && shR.getLastRow() > 1) {
      const rData = shR.getDataRange().getValues();
      for (let i = 1; i < rData.length; i++) {
        if (Number(rData[i][1]) === noTar) {
          // kolom: 0=id, 1=noTarikan, 2=tanggal, 3=sohibulBait, ...
          const tg = rData[i][2];
          tanggalTar = (tg instanceof Date)
            ? Utilities.formatDate(tg, Session.getScriptTimeZone(), 'yyyy-MM-dd')
            : String(tg);
          sohibulBait = String(rData[i][3] || '').trim();
          break;
        }
      }
    }
 
    /* ── STEP 2: Hapus baris di KasRT berdasarkan tanggal + jenis + SB ── */
    /* Ini WAJIB pertama, sebelum hapus RekapPertemuan, karena butuh info */
    const shK = ss.getSheetByName('KasRT');
    if (shK && shK.getLastRow() > 1 && tanggalTar) {
      const kData = shK.getDataRange().getValues();
      const toDelete = [];
 
      for (let i = kData.length - 1; i >= 1; i--) {
        const id     = String(kData[i][0] || '');
        const tgl    = kData[i][1];
        const jenis  = String(kData[i][2] || '').trim();
        const ket    = String(kData[i][3] || '');
 
        // Normalisasi tanggal baris ini → yyyy-MM-dd
        const tglStr = (tgl instanceof Date)
          ? Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : String(tgl);
 
        // Match jika: tanggal sama AND (jenis Iuran/Talangan terkait absensi)
        //              AND (keterangan mengandung sohibulBait)
        const isIuranTarikan = (jenis === 'Iuran Mingguan'
                             || jenis === 'Talangan Sohibul Bait');
        const ketHasSB = sohibulBait
          ? (ket.toLowerCase().indexOf(sohibulBait.toLowerCase()) >= 0)
          : false;
 
        // Fallback: match by ID prefix (KAS_IN_<tanggal>, KAS_OUT_<tanggal>)
        const isIdMatch = id === 'KAS_IN_' + tanggalTar
                       || id === 'KAS_OUT_' + tanggalTar;
 
        if ((tglStr === tanggalTar && isIuranTarikan && ketHasSB) || isIdMatch) {
          toDelete.push(i + 1); // simpan nomor baris (1-indexed)
        }
      }
 
      // Hapus dari bawah ke atas agar index tidak bergeser
      toDelete.sort((a, b) => b - a).forEach(r => shK.deleteRow(r));
      log.push('KasRT: ' + toDelete.length + ' baris dihapus');
 
      /* ── STEP 2b: Rekalkulasi saldo berjalan kolom G ── */
      if (toDelete.length > 0 && shK.getLastRow() > 1) {
        _rekalkSaldoKasRT(shK);
        log.push('KasRT: saldo berjalan dihitung ulang');
      }
    } else if (!tanggalTar) {
      log.push('KasRT: skip (tarikan tidak ditemukan di RekapPertemuan)');
    }
 
    /* ── STEP 3: Hapus baris di sheet utama berdasarkan noTarikan ── */
    const hapusSheets = [
      { name:'Pertemuan',        colNoTar:1 }, // B = noTarikan
      { name:'RekapPertemuan',   colNoTar:1 },
      { name:'KasHadiran',       colNoTar:1 },
      { name:'TalanganAnggota',  colNoTar:1 },
      { name:'RekapTuanRumah',   colNoTar:-1 } // tidak ada noTarikan, skip
    ];
 
    hapusSheets.forEach(function(cfg) {
      if (cfg.colNoTar < 0) return;
      const sheet = ss.getSheetByName(cfg.name);
      if (!sheet) { log.push(cfg.name + ': sheet tidak ada'); return; }
      const last = sheet.getLastRow();
      if (last < 2) { log.push(cfg.name + ': kosong'); return; }
 
      // Baca seluruh data sekali (lebih cepat dari getValue per cell)
      const cols = sheet.getLastColumn();
      const vals = sheet.getRange(1, 1, last, cols).getValues();
      let del = 0;
 
      // Hapus dari bawah ke atas
      for (let i = last; i >= 2; i--) {
        const v = vals[i - 1][cfg.colNoTar];
        if (Number(v) === noTar) {
          sheet.deleteRow(i);
          del++;
        }
      }
      log.push(cfg.name + ': ' + del + ' baris dihapus');
    });
 
    /* ── STEP 4: Reset status JadwalPertemuan → 'Terjadwal' ── */
    const shJ = ss.getSheetByName('JadwalPertemuan');
    if (shJ && shJ.getLastRow() > 1) {
      const jData = shJ.getDataRange().getValues();
      const headers = jData[0].map(h => String(h).toLowerCase());
      let colNo = -1, colSt = -1;
      headers.forEach((h, idx) => {
        if (h.indexOf('notarikan') >= 0 || h === 'no tarikan' || h === 'no') colNo = idx;
        if (h.indexOf('status') >= 0) colSt = idx;
      });
 
      if (colNo >= 0 && colSt >= 0) {
        for (let j = 1; j < jData.length; j++) {
          if (Number(jData[j][colNo]) === noTar) {
            shJ.getRange(j + 1, colSt + 1).setValue('Terjadwal');
            log.push('JadwalPertemuan: status direset → Terjadwal');
            break;
          }
        }
      }
    }
 
    /* ── STEP 5: Invalidate SEMUA cache supaya frontend dapat data fresh ── */
    try {
      const cache = CacheService.getScriptCache();
      cache.removeAll([
        'DATA_KAS_HISTORY', 'DATA_KAS_HADIRAN', 'DATA_REKAP',
        'DATA_TALANGAN', 'DATA_PERTEMUAN', 'DATA_JADWAL', 'DATA_ANGGOTA',
        'PAGE_dash', 'PAGE_absensi', 'PAGE_kas', 'PAGE_talangan',
        'PAGE_kasrt', 'PAGE_jadwal', 'PAGE_rekap'
      ]);
      log.push('Cache: semua key di-invalidate');
    } catch(e) {
      log.push('Cache: gagal invalidate (' + e.message + ')');
    }
 
    SpreadsheetApp.flush();
 
    return {
      success: true,
      noTarikan: noTar,
      message: 'Tarikan ke-' + noTar + (sohibulBait ? ' ('+sohibulBait+')' : '') + ' berhasil dihapus',
      detail: log.join(' · ')
    };
 
  } catch(e) {
    Logger.log('hapusTarikan error: ' + e.message + '\n' + e.stack);
    return { success:false, message:'Error: ' + e.message };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}
 
/* ── Helper: rekalkulasi saldo berjalan KasRT ── */
function _rekalkSaldoKasRT(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return;
 
  // Kolom: A=id, B=tanggal, C=jenis, D=keterangan, E=pemasukan, F=pengeluaran, G=saldo, H=ts
  const data = sheet.getRange(2, 1, last - 1, 8).getValues();
 
  // Sort by tanggal kemudian timestamp agar saldo akurat
  data.sort(function(a, b) {
    const tA = a[1] instanceof Date ? a[1].getTime() : new Date(a[1]).getTime();
    const tB = b[1] instanceof Date ? b[1].getTime() : new Date(b[1]).getTime();
    if (tA !== tB) return tA - tB;
    const tsA = a[7] ? new Date(a[7]).getTime() : 0;
    const tsB = b[7] ? new Date(b[7]).getTime() : 0;
    return tsA - tsB;
  });
 
  // Hitung ulang saldo berjalan
  let saldo = 0;
  for (let i = 0; i < data.length; i++) {
    const pemasukan  = Number(data[i][4]) || 0;
    const pengeluaran = Number(data[i][5]) || 0;
    saldo += pemasukan - pengeluaran;
    data[i][6] = saldo;
  }
 
  // Tulis kembali (sekali batch write)
  sheet.getRange(2, 1, data.length, 8).setValues(data);
}

/* ══════════════════════════════════════════════════════════════════════════
   HADIRAN RT — RAPIKAN DATA (DATA RECOVERY TOOL)
   ──────────────────────────────────────────────────────────────────────────
   Untuk: kasus data ngaco setelah "Hitung Ulang" dengan anggota berubah

   Strategi:
   - Sheet "Pertemuan" = SUMBER KEBENARAN (atomic data per orang per tarikan)
   - Sheet lain (RekapPertemuan, KasHadiran, KasRT iuran, TalanganAnggota)
     = DERIVATIF — bisa dibangun ulang dari Pertemuan

   3 Fungsi yang disediakan:
   1. cekData()           → DIAGNOSTIK (read-only, aman, lihat dulu)
   2. previewRapikan()    → SIMULASI rebuild (read-only, lihat hasilnya)
   3. rapikanData()       → EKSEKUSI rebuild (WRITE — bersih-bersih beneran)

   Cara pakai:
   1. Paste seluruh kode di bawah ke Code.gs (di paling bawah, jangan timpa)
   2. Klik  Save
   3. Di Apps Script Editor, pilih function `cekData` → klik ▶ Run
   4. Lihat hasil di View → Executions → klik run terakhir → Logger output
   5. Kalau setuju dengan rencana, jalankan `previewRapikan` untuk lihat detail
   6. Kalau sudah yakin, jalankan `rapikanData` untuk eksekusi
   ══════════════════════════════════════════════════════════════════════════ */

/* ═══════════════════════════════════════════════════════════════════════
   FUNGSI 1: cekData() — Diagnostik
   ─────────────────────────────────────────────────────────────────────
   Cek konsistensi data tanpa mengubah apapun.
   Output: log di Apps Script editor.
   ═══════════════════════════════════════════════════════════════════════ */

function cekData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];

  log.push('════════════════════════════════════════════');
  log.push('  CEK DATA HADIRAN — ' + new Date().toLocaleString('id-ID'));
  log.push('════════════════════════════════════════════');

  // ── 1. MasterAnggota ──
  const shMA = ss.getSheetByName('MasterAnggota');
  const ma = shMA ? shMA.getDataRange().getValues().slice(1) : [];
  const angAktif = ma.filter(r => String(r[3] || '').trim().toLowerCase() === 'aktif');
  log.push('\n MasterAnggota:');
  log.push('   Total tercatat: ' + ma.length);
  log.push('   Status Aktif: ' + angAktif.length);
  log.push('   Status lain: ' + (ma.length - angAktif.length));

  // ── 2. Pertemuan (sumber kebenaran) ──
  const shP = ss.getSheetByName('Pertemuan');
  const pData = shP ? shP.getDataRange().getValues().slice(1) : [];
  const tarikanMap = {}; // noTarikan → list of records
  pData.forEach(r => {
    const no = Number(r[1]);
    if (!no) return;
    if (!tarikanMap[no]) tarikanMap[no] = [];
    tarikanMap[no].push({ nama: String(r[3]).trim(), tgl: r[2] });
  });
  const nomorTarikan = Object.keys(tarikanMap).map(Number).sort((a,b) => a-b);
  log.push('\n Pertemuan (sumber data):');
  log.push('   Total baris: ' + pData.length);
  log.push('   Tarikan terdeteksi: ' + nomorTarikan.length);
  nomorTarikan.forEach(no => {
    log.push('   • Tarikan ' + no + ': ' + tarikanMap[no].length + ' anggota');
  });

  // ── 3. RekapPertemuan ──
  const shR = ss.getSheetByName('RekapPertemuan');
  const rData = shR ? shR.getDataRange().getValues().slice(1) : [];
  log.push('\n RekapPertemuan: ' + rData.length + ' baris');

  // ── 4. KasHadiran ──
  const shKH = ss.getSheetByName('KasHadiran');
  const khData = shKH ? shKH.getDataRange().getValues().slice(1) : [];
  log.push(' KasHadiran: ' + khData.length + ' baris');

  // ── 5. KasRT ──
  const shKR = ss.getSheetByName('KasRT');
  const krData = shKR ? shKR.getDataRange().getValues().slice(1) : [];
  const jenisCount = {};
  krData.forEach(r => {
    const j = String(r[2] || 'Tanpa jenis').trim();
    jenisCount[j] = (jenisCount[j] || 0) + 1;
  });
  log.push('\n KasRT: ' + krData.length + ' baris total');
  Object.keys(jenisCount).forEach(j => {
    log.push('   • ' + j + ': ' + jenisCount[j] + ' baris');
  });

  // ── 6. Cek saldo berjalan KasRT ──
  if (krData.length > 0) {
    let saldoHitung = 0;
    let mismatch = 0;
    krData.forEach((r, i) => {
      saldoHitung += (Number(r[4]) || 0) - (Number(r[5]) || 0);
      const saldoTercatat = Number(r[6]) || 0;
      if (Math.abs(saldoHitung - saldoTercatat) > 0.01) mismatch++;
    });
    const saldoAkhirHitung = saldoHitung;
    const saldoAkhirTercatat = Number(krData[krData.length - 1][6]) || 0;
    log.push('\n Validasi Saldo Berjalan KasRT:');
    log.push('   Baris dengan saldo tidak konsisten: ' + mismatch);
    log.push('   Saldo akhir hitung ulang: Rp ' + _fmt(saldoAkhirHitung));
    log.push('   Saldo akhir di sheet:    Rp ' + _fmt(saldoAkhirTercatat));
    if (Math.abs(saldoAkhirHitung - saldoAkhirTercatat) > 0.01) {
      log.push('   ⚠️ SELISIH: Rp ' + _fmt(saldoAkhirHitung - saldoAkhirTercatat) + ' — DATA NGACO!');
    } else if (mismatch > 0) {
      log.push('   ⚠️ Saldo akhir benar tapi langkah-langkah berjalannya salah');
    } else {
      log.push('   ✅ Konsisten');
    }
  }

  // ── 7. Cek talangan orphan (anggota yang sudah dihapus dari MasterAnggota) ──
  const shT = ss.getSheetByName('TalanganAnggota');
  const tData = shT ? shT.getDataRange().getValues().slice(1) : [];
  const namaAktif = new Set(ma.map(r => String(r[1] || '').trim().toLowerCase()));
  const orphanTalangan = tData.filter(r => {
    const nama = String(r[3] || '').trim().toLowerCase();
    return nama && !namaAktif.has(nama);
  });
  log.push('\n TalanganAnggota: ' + tData.length + ' baris');
  if (orphanTalangan.length > 0) {
    log.push('   ⚠️ ' + orphanTalangan.length + ' baris untuk anggota yang sudah dihapus:');
    const orphanNames = [...new Set(orphanTalangan.map(r => r[3]))];
    orphanNames.forEach(n => log.push('      • ' + n));
  } else {
    log.push('   ✅ Tidak ada talangan orphan');
  }

  // ── 8. Cek konsistensi: Pertemuan vs RekapPertemuan ──
  log.push('\n Konsistensi Pertemuan ↔ RekapPertemuan:');
  let inconsistent = 0;
  nomorTarikan.forEach(no => {
    const pCount = tarikanMap[no].length;
    const rRow = rData.find(r => Number(r[1]) === no);
    if (!rRow) {
      log.push('   ⚠️ Tarikan ' + no + ': ada di Pertemuan tapi TIDAK ADA di RekapPertemuan');
      inconsistent++;
    } else {
      const totalAng = Number(rRow[6]) || 0;
      if (totalAng !== pCount) {
        log.push('   ⚠️ Tarikan ' + no + ': Pertemuan punya ' + pCount + ' baris, RekapPertemuan totalAnggota=' + totalAng);
        inconsistent++;
      }
    }
  });
  if (inconsistent === 0) log.push('   ✅ Konsisten');

  log.push('\n════════════════════════════════════════════');
  log.push('  KESIMPULAN:');
  if (orphanTalangan.length > 0 || inconsistent > 0) {
    log.push('  ⚠️ DATA NGACO — perlu rapikanData()');
    log.push('  Langkah: jalankan previewRapikan() dulu untuk simulasi');
  } else {
    log.push('  ✅ Data dalam kondisi konsisten');
  }
  log.push('════════════════════════════════════════════');

  const output = log.join('\n');
  Logger.log(output);
  return output;
}

/* ═══════════════════════════════════════════════════════════════════════
   FUNGSI 2: previewRapikan() — SIMULASI tanpa write
   ═══════════════════════════════════════════════════════════════════════ */

function previewRapikan() {
  return _rapikanData(true);
}

/* ═══════════════════════════════════════════════════════════════════════
   FUNGSI 3: rapikanData() — EKSEKUSI rebuild
   ═══════════════════════════════════════════════════════════════════════ */

function rapikanData() {
  return _rapikanData(false);
}

/* ═══════════════════════════════════════════════════════════════════════
   INTERNAL: _rapikanData(dryRun)
   Rebuild semua data turunan dari Pertemuan.
   Manual entries di KasRT (Pembayaran Talangan, Pengeluaran, Setoran)
   AKAN DIPERTAHANKAN — hanya 'Iuran Mingguan' & 'Talangan Sohibul Bait'
   yang di-rebuild.
   ═══════════════════════════════════════════════════════════════════════ */

function _rapikanData(dryRun) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch(e) { return { success:false, message:'Server sibuk, coba lagi' }; }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];
  const stats = { deleted:0, rebuilt:0, kept:0 };

  log.push('════════════════════════════════════════════');
  log.push('  ' + (dryRun ? 'PREVIEW' : 'EKSEKUSI') + ' RAPIKAN DATA');
  log.push('════════════════════════════════════════════');

  try {
    /* ── KONSTANTA — sesuaikan dengan setup spreadsheet ── */
    const IURAN_TOTAL = 50000;
    const IURAN_SB    = 45000;
    const IURAN_KAS   = 5000;

    /* ── STEP 1: Baca Pertemuan (sumber kebenaran) ── */
    const shP = ss.getSheetByName('Pertemuan');
    if (!shP) throw new Error('Sheet Pertemuan tidak ditemukan');
    const pData = shP.getDataRange().getValues().slice(1);

    // Group by noTarikan
    const tarikanMap = {};
    pData.forEach(r => {
      const no = Number(r[1]);
      if (!no) return;
      if (!tarikanMap[no]) {
        tarikanMap[no] = {
          noTarikan: no,
          tanggal: _toTglStr(r[2]),
          anggota: []
        };
      }
      tarikanMap[no].anggota.push({
        nama: String(r[3]).trim(),
        hadir: String(r[5]).trim() === 'Ya',
        bayar: String(r[6]).trim() === 'Ya',
        titip: String(r[7]).trim() === 'Ya',
        catatan: String(r[8] || '').trim()
      });
    });
    const nomorTarikan = Object.keys(tarikanMap).map(Number).sort((a,b) => a-b);
    log.push('\n[1] Baca Pertemuan: ' + nomorTarikan.length + ' tarikan ditemukan');

    /* ── STEP 2: Backup KasRT manual entries ── */
    const shKR = ss.getSheetByName('KasRT');
    if (!shKR) throw new Error('Sheet KasRT tidak ditemukan');
    const krData = shKR.getDataRange().getValues();
    const krHeader = krData[0];
    const krRows = krData.slice(1);

    const manualEntries = krRows.filter(r => {
      const jenis = String(r[2] || '').trim();
      // Yang bukan iuran/talangan otomatis = entri manual yang harus dipertahankan
      return jenis !== 'Iuran Mingguan' && jenis !== 'Talangan Sohibul Bait';
    });
    log.push('[2] KasRT entries:');
    log.push('    Manual (dipertahankan): ' + manualEntries.length);
    log.push('    Iuran/Talangan (akan rebuild): ' + (krRows.length - manualEntries.length));
    stats.kept += manualEntries.length;

    /* ── STEP 3: Rebuild RekapPertemuan & KasHadiran ── */
    const shR  = ss.getSheetByName('RekapPertemuan');
    const shKH = ss.getSheetByName('KasHadiran');
    if (!shR || !shKH) throw new Error('Sheet RekapPertemuan atau KasHadiran tidak ditemukan');

    const newRekap = [];
    const newKasHadiran = [];
    const newKasIuran = []; // akan digabung dengan manualEntries
    const ts = new Date().toISOString();

    nomorTarikan.forEach((no, idx) => {
      const tar = tarikanMap[no];
      const sbRecord = tar.anggota.find(a =>
        String(a.catatan).toLowerCase().indexOf('sohibul') >= 0
      );
      const sb = sbRecord ? sbRecord.nama : (tar.anggota[0] ? tar.anggota[0].nama : '');

      const totalAnggota     = tar.anggota.length;
      const hadirList        = tar.anggota.filter(a => a.hadir || a.titip);
      const tidakHadirList   = tar.anggota.filter(a => !a.hadir && !a.titip);
      const jumlahHadir      = hadirList.length;
      const jumlahTidakHadir = tidakHadirList.length;
      const totalBayarMasuk  = jumlahHadir * IURAN_TOTAL;
      const sbDiterima       = (totalAnggota - 1) * IURAN_SB;
      const kasDariHadir     = Math.max(0, (jumlahHadir - 1)) * IURAN_KAS;
      const talanganKeluar   = jumlahTidakHadir * IURAN_TOTAL;
      const kasDariTalangan  = jumlahTidakHadir * IURAN_KAS;
      const totalKasMasuk    = kasDariHadir + kasDariTalangan;
      const netKas           = totalKasMasuk - talanganKeluar;

      const noUrut = idx + 1;
      const tanggal = tar.tanggal;

      newRekap.push([
        noUrut, no, tanggal, sb, jumlahHadir, jumlahTidakHadir,
        totalAnggota, totalBayarMasuk, sbDiterima,
        totalKasMasuk, talanganKeluar, netKas
      ]);

      newKasHadiran.push([
        noUrut, no, tanggal, sb, totalAnggota,
        jumlahHadir, jumlahTidakHadir, totalBayarMasuk, sbDiterima,
        totalKasMasuk, talanganKeluar, netKas, '', ts
      ]);

      // KasRT iuran row
      newKasIuran.push([
        'KAS_IN_' + tanggal + '_T' + no,
        tanggal,
        'Iuran Mingguan',
        'Iuran ' + Math.max(0, jumlahHadir-1) + ' hadir (excl.SB) + ' + jumlahTidakHadir + ' talangan → SB: ' + sb,
        totalKasMasuk,
        0,
        0, // saldo: akan dihitung di step terakhir
        ts
      ]);

      // KasRT talangan row (kalau ada)
      if (talanganKeluar > 0) {
        newKasIuran.push([
          'KAS_OUT_' + tanggal + '_T' + no,
          tanggal,
          'Talangan Sohibul Bait',
          'Talangan ' + jumlahTidakHadir + ' anggota (50k/org) → ' + sb,
          0,
          talanganKeluar,
          0, // saldo: dihitung nanti
          ts
        ]);
      }

      stats.rebuilt++;
      log.push('    Tarikan ' + no + ': ' + totalAnggota + ' anggota, ' +
               jumlahHadir + '/' + jumlahTidakHadir + ' (H/T), Kas Rp ' + _fmt(totalKasMasuk) +
               (talanganKeluar > 0 ? ', Talangan Rp ' + _fmt(talanganKeluar) : ''));
    });

    log.push('\n[3] Rebuild ringkasan: ' + stats.rebuilt + ' tarikan dihitung ulang');

    /* ── STEP 4: Rebuild TalanganAnggota ── */
    const shT = ss.getSheetByName('TalanganAnggota');
    if (!shT) throw new Error('Sheet TalanganAnggota tidak ditemukan');

    // Cek pembayaran talangan yang sudah ada di KasRT manual entries
    const pembayaranTalangan = manualEntries.filter(r =>
      String(r[2] || '').toLowerCase().indexOf('pembayaran talangan') >= 0 ||
      String(r[2] || '').toLowerCase().indexOf('koreksi talangan') >= 0
    );

    const newTalangan = [];
    nomorTarikan.forEach(no => {
      const tar = tarikanMap[no];
      const tidakHadir = tar.anggota.filter(a => !a.hadir && !a.titip);

      tidakHadir.forEach(a => {
        // Cek apakah sudah lunas di manual entries
        const lunas = pembayaranTalangan.find(p => {
          const ket = String(p[3] || '').toLowerCase();
          return ket.indexOf(a.nama.toLowerCase()) >= 0 &&
                 (ket.indexOf('tarikan #' + no) >= 0 ||
                  ket.indexOf('tarikan-' + no) >= 0 ||
                  ket.indexOf('ke-' + no) >= 0);
        });

        newTalangan.push([
          'TAL_' + no + '_' + a.nama,
          no,
          tar.tanggal,
          a.nama,
          IURAN_TOTAL,
          lunas ? 'Lunas' : 'Belum Lunas',
          lunas ? _toTglStr(lunas[1]) : '',
          ts
        ]);
      });
    });
    log.push('[4] Rebuild TalanganAnggota: ' + newTalangan.length + ' baris');

    /* ── STEP 5: Gabung newKasIuran + manualEntries, sort by tanggal, rekalk saldo ── */
    const allKasRows = newKasIuran.concat(manualEntries);

    // Sort by tanggal (kolom B), kemudian timestamp (kolom H)
    allKasRows.sort((a, b) => {
      const tA = _toTglStr(a[1]);
      const tB = _toTglStr(b[1]);
      if (tA !== tB) return tA.localeCompare(tB);
      const tsA = a[7] ? new Date(a[7]).getTime() : 0;
      const tsB = b[7] ? new Date(b[7]).getTime() : 0;
      return tsA - tsB;
    });

    // Rekalkulasi saldo berjalan
    let saldo = 0;
    allKasRows.forEach(r => {
      saldo += (Number(r[4]) || 0) - (Number(r[5]) || 0);
      r[6] = saldo;
    });
    log.push('[5] Rekalkulasi saldo berjalan KasRT: ' + allKasRows.length + ' baris');
    log.push('    Saldo akhir: Rp ' + _fmt(saldo));

    /* ── EKSEKUSI (kalau bukan dryRun) ── */
    if (!dryRun) {
      log.push('\n[6] WRITE ke spreadsheet:');

      // Clear RekapPertemuan (kecuali header)
      if (shR.getLastRow() > 1) {
        shR.getRange(2, 1, shR.getLastRow() - 1, shR.getLastColumn()).clear();
      }
      if (newRekap.length > 0) {
        shR.getRange(2, 1, newRekap.length, newRekap[0].length).setValues(newRekap);
      }
      log.push('    ✓ RekapPertemuan ditulis ulang (' + newRekap.length + ' baris)');

      // Clear KasHadiran
      if (shKH.getLastRow() > 1) {
        shKH.getRange(2, 1, shKH.getLastRow() - 1, shKH.getLastColumn()).clear();
      }
      if (newKasHadiran.length > 0) {
        shKH.getRange(2, 1, newKasHadiran.length, newKasHadiran[0].length).setValues(newKasHadiran);
      }
      log.push('    ✓ KasHadiran ditulis ulang (' + newKasHadiran.length + ' baris)');

      // Clear TalanganAnggota
      if (shT.getLastRow() > 1) {
        shT.getRange(2, 1, shT.getLastRow() - 1, shT.getLastColumn()).clear();
      }
      if (newTalangan.length > 0) {
        shT.getRange(2, 1, newTalangan.length, newTalangan[0].length).setValues(newTalangan);
      }
      log.push('    ✓ TalanganAnggota ditulis ulang (' + newTalangan.length + ' baris)');

      // Clear KasRT (kecuali header) lalu tulis allKasRows
      if (shKR.getLastRow() > 1) {
        shKR.getRange(2, 1, shKR.getLastRow() - 1, shKR.getLastColumn()).clear();
      }
      if (allKasRows.length > 0) {
        shKR.getRange(2, 1, allKasRows.length, allKasRows[0].length).setValues(allKasRows);
      }
      log.push('    ✓ KasRT diurutkan + saldo direkalkulasi (' + allKasRows.length + ' baris)');

      // Invalidate cache
      try {
        const cache = CacheService.getScriptCache();
        cache.removeAll([
          'DATA_KAS_HISTORY', 'DATA_KAS_HADIRAN', 'DATA_REKAP',
          'DATA_TALANGAN', 'DATA_PERTEMUAN', 'DATA_JADWAL', 'DATA_ANGGOTA',
          'PAGE_dash', 'PAGE_absensi', 'PAGE_kas', 'PAGE_talangan',
          'PAGE_kasrt', 'PAGE_jadwal', 'PAGE_rekap'
        ]);
        log.push('    ✓ Cache di-invalidate');
      } catch(e) {}

      SpreadsheetApp.flush();
      log.push('\n✅ SELESAI — silakan refresh aplikasi Hadiran');
    } else {
      log.push('\n INI HANYA SIMULASI — tidak ada yang diubah');
      log.push('   Untuk eksekusi nyata, jalankan: rapikanData()');
    }

    const output = log.join('\n');
    Logger.log(output);

    return {
      success: true,
      dryRun: dryRun,
      stats: stats,
      saldoAkhir: saldo,
      report: output
    };

  } catch(e) {
    log.push('\n❌ ERROR: ' + e.message);
    Logger.log(log.join('\n'));
    return { success:false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

/* ═══════════════════════════════════════════════════════════════════════
   HELPER
   ═══════════════════════════════════════════════════════════════════════ */

function _toTglStr(v) {
  if (!v) return '';
  if (typeof v === 'string') return v.substring(0, 10);
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v).substring(0, 10);
}

function _fmt(n) {
  n = Number(n) || 0;
  return n.toLocaleString('id-ID');
}

/* ══════════════════════════════════════════════════════════════════════════
   BONUS PATCH: Auto-rekalk saldo setiap kali simpanAbsensi dipanggil
   ──────────────────────────────────────────────────────────────────────────
   OPSIONAL. Kalau Anda mau MENCEGAH masalah ini terjadi lagi di masa depan,
   cari function `simpanAbsensi` di Code.gs Anda, lalu tambahkan baris ini
   DI ATAS baris `return { success:true, ...` di akhir function:

       _rekalkSaldoOtomatis();

   Lalu tambahkan helper berikut di akhir Code.gs:
   ══════════════════════════════════════════════════════════════════════════ */

function _rekalkSaldoOtomatis() {
  const shKR = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KasRT');
  if (!shKR || shKR.getLastRow() < 2) return;

  const data = shKR.getRange(2, 1, shKR.getLastRow() - 1, 8).getValues();

  // Sort by tanggal kemudian timestamp
  data.sort((a, b) => {
    const tA = _toTglStr(a[1]);
    const tB = _toTglStr(b[1]);
    if (tA !== tB) return tA.localeCompare(tB);
    const tsA = a[7] ? new Date(a[7]).getTime() : 0;
    const tsB = b[7] ? new Date(b[7]).getTime() : 0;
    return tsA - tsB;
  });

  // Rekalkulasi saldo
  let saldo = 0;
  data.forEach(r => {
    saldo += (Number(r[4]) || 0) - (Number(r[5]) || 0);
    r[6] = saldo;
  });

  shKR.getRange(2, 1, data.length, 8).setValues(data);
}

/* ══════════════════════════════════════════════════════════════════════════
   ALUR LENGKAP YANG DISARANKAN
   ──────────────────────────────────────────────────────────────────────────

    LANGKAH 1: Cek dulu (read-only, aman)
      → Apps Script Editor → pilih function `cekData` → ▶ Run
      → Lihat output di View → Executions → run terakhir → Logs
      → Kalau ada warning ⚠️ "DATA NGACO" lanjut ke step 2

    LANGKAH 2: Simulasi (read-only juga)
      → Pilih function `previewRapikan` → ▶ Run
      → Lihat di Logs: per-tarikan dihitung berapa, saldo akhir berapa
      → Bandingkan dengan ekspektasi Anda

    LANGKAH 3: Eksekusi (WRITE — hati-hati)
      → Pilih function `rapikanData` → ▶ Run
      → Tunggu sampai selesai (5-30 detik tergantung jumlah data)
      → Refresh aplikasi Hadiran → saldo & rekap akan benar

    LANGKAH 4 (pencegahan masa depan):
      → Edit function `simpanAbsensi` di Code.gs
      → Tambahkan `_rekalkSaldoOtomatis();` sebelum return statement
      → Save & Deploy ulang
      → Sekarang setiap save absensi akan auto-rekalk saldo

   ═══════════════════════════════════════════════════════════════════════ */

   /* ══════════════════════════════════════════════════════════════════════════
    PASTE #1 — Code.gs (paste di paling bawah, jangan timpa yang ada)
   ══════════════════════════════════════════════════════════════════════════ */
 
function getAllAnggota() {
  const cache = CacheService.getScriptCache();
  const key = 'DATA_ANGGOTA_ALL';
  const cached = cache.get(key);
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }
 
  const sh = getSheet('MasterAnggota');
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
 
  const result = data.slice(1)
    .filter(r => r[1])  // hanya yang punya nama
    .map((r, i) => ({
      no:     i + 1,
      nama:   String(r[1] || '').trim(),
      noHp:   String(r[2] || '').trim(),
      status: String(r[3] || 'Aktif').trim()
    }));
 
  try { cache.put(key, JSON.stringify(result), CACHE_TTL_LONG); } catch(e) {}
  return result;
}
 
 
function editAnggotaStatus(payload) {
  if (!payload || !payload.nama) {
    return { success: false, message: 'Nama anggota tidak valid' };
  }
 
  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); }
  catch(e) { return { success: false, message: 'Server sibuk, coba lagi' }; }
 
  try {
    const sh = getSheet('MasterAnggota');
    if (!sh) return { success: false, message: 'Sheet MasterAnggota tidak ditemukan' };
 
    const data = sh.getDataRange().getValues();
    const targetNama = String(payload.nama).trim().toLowerCase();
    let foundRow = -1;
 
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1] || '').trim().toLowerCase() === targetNama) {
        foundRow = i + 1;  // 1-indexed
        break;
      }
    }
 
    if (foundRow < 0) {
      return { success: false, message: 'Anggota "' + payload.nama + '" tidak ditemukan' };
    }
 
    const updates = [];
 
    // Update No HP (kolom C / index 3)
    if (payload.noHp !== undefined) {
      sh.getRange(foundRow, 3).setValue(String(payload.noHp || '').trim());
      updates.push('No HP');
    }
 
    // Update Status (kolom D / index 4)
    if (payload.status !== undefined) {
      const validStatus = ['Aktif','Nonaktif','Mengundurkan Diri','Pindah','Almarhum'];
      const newStatus = String(payload.status || 'Aktif').trim();
      if (validStatus.indexOf(newStatus) < 0) {
        return { success: false, message: 'Status tidak valid: ' + newStatus };
      }
      sh.getRange(foundRow, 4).setValue(newStatus);
      updates.push('Status → ' + newStatus);
    }
 
    SpreadsheetApp.flush();
 
    // Invalidate cache
    try {
      const cache = CacheService.getScriptCache();
      cache.removeAll([
        'DATA_ANGGOTA', 'DATA_ANGGOTA_ALL',
        'PAGE_dash', 'PAGE_absensi'
      ]);
    } catch(e) {}
 
    return {
      success: true,
      message: updates.join(' & ') + ' berhasil diperbarui',
      nama: payload.nama
    };
 
  } catch(e) {
    Logger.log('editAnggotaStatus error: ' + e.message);
    return { success: false, message: 'Error: ' + e.message };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

/* ══════════════════════════════════════════════════════════════════════════
    PASTE #1 — Code.gs (paste di paling bawah Code.gs)
   ══════════════════════════════════════════════════════════════════════════ */

function hapusTalangan(payload) {
  if (!payload || !payload.id) {
    return { success: false, message: 'ID talangan tidak valid' };
  }

  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); }
  catch(e) { return { success: false, message: 'Server sibuk, coba lagi' }; }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('TalanganAnggota');
    if (!sh) return { success: false, message: 'Sheet TalanganAnggota tidak ada' };

    const last = sh.getLastRow();
    if (last < 2) return { success: false, message: 'Sheet kosong' };

    const data = sh.getRange(1, 1, last, sh.getLastColumn()).getValues();
    const targetId = String(payload.id).trim();
    let foundRow = -1;
    let info = null;

    // Cari berdasarkan ID (kolom A) — paling akurat
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === targetId) {
        foundRow = i + 1;
        info = {
          noTarikan: data[i][1],
          tanggal: data[i][2],
          nama: data[i][3],
          jumlah: data[i][4],
          status: data[i][5]
        };
        break;
      }
    }

    if (foundRow < 0) {
      return { success: false, message: 'Talangan dengan ID ' + targetId + ' tidak ditemukan' };
    }

    // Cegah hapus talangan yang sudah lunas (kalau sudah lunas, hapus = duit hilang dari kas)
    if (String(info.status).toLowerCase() === 'lunas' && !payload.forceLunas) {
      return {
        success: false,
        message: 'Talangan ini sudah LUNAS. Jika dihapus, pembayaran tsb akan tampak masih masuk di kas. Yakin? Set forceLunas:true untuk override.'
      };
    }

    // Hapus baris
    sh.deleteRow(foundRow);

    // OPSIONAL: catat di KasRT sebagai keterangan write-off (audit trail)
    if (payload.writeOff) {
      const shK = ss.getSheetByName('KasRT');
      if (shK) {
        const ts = new Date().toISOString();
        const tglStr = (info.tanggal instanceof Date)
          ? Utilities.formatDate(info.tanggal, Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : String(info.tanggal).substring(0, 10);

        // Catatan write-off: pemasukan = pengeluaran = 0, hanya untuk audit
        shK.appendRow([
          'WO_TAL_' + targetId,
          new Date(),
          'Write-off Talangan',
          'Penghapusan talangan ' + info.nama + ' (Tarikan #' + info.noTarikan + ') — ' + (payload.alasan || 'Resign/Tidak terbayar'),
          0,
          0,
          0, // saldo: akan dihitung ulang
          ts
        ]);

        // Rekalk saldo (kalau function ada)
        if (typeof _rekalkSaldoKasRT === 'function') {
          _rekalkSaldoKasRT(shK);
        }
      }
    }

    SpreadsheetApp.flush();

    // Invalidate cache
    try {
      const cache = CacheService.getScriptCache();
      cache.removeAll([
        'DATA_TALANGAN', 'DATA_KAS_HISTORY', 'DATA_KAS_HADIRAN',
        'PAGE_dash', 'PAGE_talangan', 'PAGE_kas', 'PAGE_kasrt'
      ]);
    } catch(e) {}

    return {
      success: true,
      message: 'Talangan ' + info.nama + ' (Tarikan #' + info.noTarikan + ') sebesar Rp ' +
               Number(info.jumlah).toLocaleString('id-ID') + ' berhasil dihapus',
      info: info
    };

  } catch(e) {
    Logger.log('hapusTalangan error: ' + e.message);
    return { success: false, message: 'Error: ' + e.message };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}
// ── fixTalanganMusa — koreksi talangan tarikan 1 + hapus write-off ──
function fixTalanganMusa() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shK = ss.getSheetByName('KasRT');
  const data = shK.getDataRange().getValues();

  for (let i = data.length - 1; i >= 1; i--) {
    const id    = String(data[i][0] || '');
    const jenis = String(data[i][2] || '');

    // Hapus Write-off Talangan (tidak perlu, 0 pemasukan 0 pengeluaran)
    if (jenis === 'Write-off Talangan') {
      shK.deleteRow(i + 1);
      continue;
    }

    // Koreksi talangan tarikan 1: dari 650000 → 600000
    if (id === 'KAS_OUT_2026-05-02_T1' ||
        (jenis === 'Talangan Sohibul Bait' &&
         String(data[i][1]).indexOf('2026-05-02') >= 0)) {
      shK.getRange(i + 1, 6).setValue(600000); // kolom F = Pengeluaran
      Logger.log('Koreksi talangan tarikan 1: 650000 → 600000');
    }
  }

  // Rekalkulasi semua saldo berjalan
  _rekalkSaldoOtomatis();

  // Clear cache
  CacheService.getScriptCache().removeAll([
    'DATA_KAS_HISTORY', 'DATA_KAS_HADIRAN', 'DATA_REKAP',
    'DATA_TALANGAN', 'PAGE_dash', 'PAGE_kas', 'PAGE_talangan'
  ]);

  SpreadsheetApp.flush();
  Logger.log('Selesai. Saldo akhir: Rp ' + getSaldoKas().toLocaleString('id-ID'));
}
function diagnoseTalangan() {
  const shT = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TalanganAnggota');
  const data = shT.getDataRange().getValues();
  let belumLunas = 0, lunas = 0, total = 0;
  
  data.slice(1).forEach(r => {
    const status = String(r[5] || '').trim();
    const jumlah = Number(r[4]) || 0;
    Logger.log('Tarikan ' + r[1] + ' | ' + r[3] + ' | ' + status + ' | Rp' + jumlah.toLocaleString('id-ID'));
    if (status === 'Belum Lunas') { belumLunas++; total += jumlah; }
    else lunas++;
  });
  
  Logger.log('──────────────────────────────');
  Logger.log('Belum Lunas: ' + belumLunas + ' orang = Rp' + total.toLocaleString('id-ID'));
  Logger.log('Lunas: ' + lunas + ' orang');
}

function hapusMusaDariTalangan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shT = ss.getSheetByName('TalanganAnggota');
  const data = shT.getDataRange().getValues();
  
  const toDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    const nama = String(data[i][3] || '').trim().toLowerCase();
    if (nama === 'musa') {
      toDelete.push(i + 1);
      Logger.log('Hapus baris ' + (i+1) + ': Musa Tarikan ' + data[i][1]);
    }
  }
  
  toDelete.sort((a,b) => b-a).forEach(r => shT.deleteRow(r));
  Logger.log('Total dihapus: ' + toDelete.length + ' baris');

  // Clear cache
  CacheService.getScriptCache().removeAll([
    'DATA_TALANGAN', 'DATA_KAS_HISTORY', 'DATA_KAS_HADIRAN',
    'PAGE_dash', 'PAGE_kas', 'PAGE_talangan'
  ]);

  SpreadsheetApp.flush();
  Logger.log('Selesai. Saldo: Rp' + getSaldoKas().toLocaleString('id-ID'));
  Logger.log('Sisa talangan belum lunas: Rp' + 
    getTalanganList().filter(t => t.status === 'Belum Lunas')
    .reduce((s,t) => s+t.jumlah, 0).toLocaleString('id-ID'));
}

function fixTalanganTarikan2Musa() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shK = ss.getSheetByName('KasRT');
  const data = shK.getDataRange().getValues();

  for (let i = data.length - 1; i >= 1; i--) {
    const id    = String(data[i][0] || '');
    const jenis = String(data[i][2] || '');
    const tgl   = String(data[i][1] || '');

    // Koreksi talangan tarikan 2: dari 500000 → 450000
    if (id === 'KAS_OUT_2026-05-09_T2' ||
        (jenis === 'Talangan Sohibul Bait' &&
         (tgl.indexOf('2026-05-09') >= 0 || String(data[i][1]) === '2026-05-09'))) {
      shK.getRange(i + 1, 6).setValue(450000); // kolom F = Pengeluaran
      Logger.log('Koreksi talangan tarikan 2: 500000 → 450000');
    }
  }

  // Rekalkulasi semua saldo berjalan
  _rekalkSaldoOtomatis();

  // Clear cache
  CacheService.getScriptCache().removeAll([
    'DATA_KAS_HISTORY', 'DATA_KAS_HADIRAN', 'DATA_REKAP',
    'DATA_TALANGAN', 'PAGE_dash', 'PAGE_kas', 'PAGE_talangan'
  ]);

  SpreadsheetApp.flush();
  Logger.log('Selesai. Saldo akhir: Rp' + getSaldoKas().toLocaleString('id-ID'));
}
function fixSaldoMusaTarikan1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shK = ss.getSheetByName('KasRT');
  const ts = new Date().toISOString();

  // Tambah pemasukan 50.000 sebagai koreksi hapus talangan Musa Tarikan 1
  shK.appendRow([
    'KAS_KOREKSI_MUSA_T1',
    '2026-05-02',
    'Koreksi Talangan',
    'Koreksi hapus talangan Musa Tarikan 1 (mengundurkan diri)',
    50000, // pemasukan
    0,
    0,    // saldo dihitung ulang
    ts
  ]);

  // Rekalkulasi saldo
  _rekalkSaldoOtomatis();

  // Clear cache
  CacheService.getScriptCache().removeAll([
    'DATA_KAS_HISTORY', 'DATA_KAS_HADIRAN', 'DATA_REKAP',
    'DATA_TALANGAN', 'PAGE_dash', 'PAGE_kas', 'PAGE_talangan'
  ]);

  SpreadsheetApp.flush();
  Logger.log('Selesai. Saldo akhir: Rp' + getSaldoKas().toLocaleString('id-ID'));
}

function batalKoreksiMusaT1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shK = ss.getSheetByName('KasRT');
  const data = shK.getDataRange().getValues();

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === 'KAS_KOREKSI_MUSA_T1') {
      shK.deleteRow(i + 1);
      Logger.log('Baris koreksi Musa T1 dihapus');
      break;
    }
  }

  _rekalkSaldoOtomatis();

  CacheService.getScriptCache().removeAll([
    'DATA_KAS_HISTORY', 'DATA_KAS_HADIRAN', 'DATA_REKAP',
    'DATA_TALANGAN', 'PAGE_dash', 'PAGE_kas', 'PAGE_talangan'
  ]);

  SpreadsheetApp.flush();
  Logger.log('Selesai. Saldo akhir: Rp' + getSaldoKas().toLocaleString('id-ID'));
}
function fixTalanganFinal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shT = ss.getSheetByName('TalanganAnggota');
  const data = shT.getDataRange().getValues();

  // Hapus SEMUA entri Musa yang masih ada
  const toDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    const nama = String(data[i][3] || '').trim().toLowerCase();
    if (nama === 'musa') {
      toDelete.push(i + 1);
      Logger.log('Hapus: Musa Tarikan ' + data[i][1] + ' | ' + data[i][5]);
    }
  }
  toDelete.sort((a,b) => b-a).forEach(r => shT.deleteRow(r));
  Logger.log('Total dihapus: ' + toDelete.length + ' baris Musa');

  // Cek sisa talangan
  const sisa = shT.getDataRange().getValues().slice(1);
  let totalBelumLunas = 0;
  sisa.forEach(r => {
    if (String(r[5]).trim() === 'Belum Lunas') {
      totalBelumLunas += Number(r[4]) || 0;
    }
  });

  CacheService.getScriptCache().removeAll([
    'DATA_TALANGAN', 'DATA_KAS_HISTORY', 'DATA_KAS_HADIRAN',
    'PAGE_dash', 'PAGE_kas', 'PAGE_talangan'
  ]);

  SpreadsheetApp.flush();
  Logger.log('Saldo KasRT: Rp' + getSaldoKas().toLocaleString('id-ID'));
  Logger.log('Talangan Belum Lunas: Rp' + totalBelumLunas.toLocaleString('id-ID'));
}
function cekSemuaTalangan() {
  const shT = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TalanganAnggota');
  const data = shT.getDataRange().getValues().slice(1);
  
  let totalBelumLunas = 0;
  Logger.log('=== SEMUA TALANGAN ===');
  data.forEach((r, i) => {
    const status = String(r[5]).trim();
    const jumlah = Number(r[4]) || 0;
    Logger.log((i+1) + '. Tarikan ' + r[1] + ' | ' + r[3] + ' | ' + status + ' | Rp' + jumlah.toLocaleString('id-ID'));
    if (status === 'Belum Lunas') totalBelumLunas += jumlah;
  });
  Logger.log('──────────────────────');
  Logger.log('Total Belum Lunas: Rp' + totalBelumLunas.toLocaleString('id-ID'));
  Logger.log('Total baris: ' + data.length);
}
function cekPembayaranTalangan() {
  const shK = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KasRT');
  const data = shK.getDataRange().getValues().slice(1);
  
  Logger.log('=== Pembayaran Talangan di KasRT ===');
  data.forEach(r => {
    const jenis = String(r[2] || '').trim();
    if (jenis === 'Pembayaran Talangan') {
      Logger.log(r[0] + ' | ' + r[1] + ' | ' + r[3] + ' | Rp' + Number(r[4]).toLocaleString('id-ID'));
    }
  });
}

function fixStatusLunasTarikan1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shT = ss.getSheetByName('TalanganAnggota');
  const data = shT.getDataRange().getValues();

  // 9 orang yang sudah bayar beserta tanggal bayarnya
  const sudahLunas = {
    'rahmat deni'      : '2026-05-03',
    'ahmad fauzi'      : '2026-05-09',
    'yulianto'         : '2026-05-09',
    'h.kadimun'        : '2026-05-09',
    'komarudin'        : '2026-05-09',
    'h.deny fredianto' : '2026-05-09',
    'tasman h.sainan'  : '2026-05-09',
    'ach.bayu octamar' : '2026-05-09',
    'idham firmansyah' : '2026-05-09',
  };

  let fixed = 0;
  for (let i = 1; i < data.length; i++) {
    const noTarikan = Number(data[i][1]);
    const nama      = String(data[i][3] || '').trim().toLowerCase();
    const status    = String(data[i][5] || '').trim();

    // Hanya update Tarikan 1 yang masih Belum Lunas
    if (noTarikan === 1 && status === 'Belum Lunas' && sudahLunas[nama]) {
      shT.getRange(i + 1, 6).setValue('Lunas');
      shT.getRange(i + 1, 7).setValue(sudahLunas[nama]);
      Logger.log('✓ Lunas: ' + data[i][3] + ' | ' + sudahLunas[nama]);
      fixed++;
    }
  }

  CacheService.getScriptCache().removeAll([
    'DATA_TALANGAN', 'DATA_KAS_HISTORY', 'DATA_KAS_HADIRAN',
    'PAGE_dash', 'PAGE_kas', 'PAGE_talangan'
  ]);

  SpreadsheetApp.flush();
  Logger.log('Total diupdate: ' + fixed + ' orang');
  Logger.log('Saldo KasRT: Rp' + getSaldoKas().toLocaleString('id-ID'));

  // Cek sisa belum lunas
  const sisa = shT.getDataRange().getValues().slice(1);
  let totalBL = 0;
  sisa.forEach(r => {
    if (String(r[5]).trim() === 'Belum Lunas') totalBL += Number(r[4]) || 0;
  });
  Logger.log('Talangan Belum Lunas: Rp' + totalBL.toLocaleString('id-ID'));
}

function clearAllCache() {
  CacheService.getScriptCache().removeAll([
    'DATA_KAS_HISTORY','DATA_KAS_HADIRAN','DATA_REKAP',
    'DATA_TALANGAN','DATA_ANGGOTA','DATA_JADWAL',
    'PAGE_dash','PAGE_absensi','PAGE_kas','PAGE_talangan',
    'PAGE_kasrt','HTML_INDEX'
  ]);
  Logger.log('Cache cleared — siap pertemuan!');
}
function cekMusaDiPertemuan() {
  const shP = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pertemuan');
  const data = shP.getDataRange().getValues().slice(1);
  const musa = data.filter(r => String(r[3]||'').trim().toLowerCase() === 'musa');
  
  if (musa.length === 0) {
    Logger.log('✅ Musa TIDAK ADA di sheet Pertemuan — PDF aman!');
  } else {
    musa.forEach(r => {
      Logger.log('⚠️ Musa masih ada: Tarikan ' + r[1] + ' | Tgl ' + r[2] + ' | Hadir: ' + r[5]);
    });
  }
}

function hapusMusaDariPertemuan() {
  const shP = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pertemuan');
  const data = shP.getDataRange().getValues();
  const toDelete = [];
  
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][3]||'').trim().toLowerCase() === 'musa') {
      toDelete.push(i + 1);
      Logger.log('Hapus: Musa Tarikan ' + data[i][1]);
    }
  }
  
  toDelete.sort((a,b) => b-a).forEach(r => shP.deleteRow(r));
  
  CacheService.getScriptCache().removeAll([
    'DATA_REKAP','DATA_KAS_HADIRAN','DATA_KAS_HISTORY',
    'PAGE_dash','PAGE_absensi','PAGE_kas','PAGE_talangan'
  ]);
  
  SpreadsheetApp.flush();
  Logger.log('Selesai. ' + toDelete.length + ' baris Musa dihapus dari Pertemuan');
  
  // Verifikasi ulang
  const cek = shP.getDataRange().getValues().slice(1)
    .filter(r => String(r[3]||'').trim().toLowerCase() === 'musa');
  Logger.log(cek.length === 0 ? '✅ Musa sudah bersih dari Pertemuan' : '⚠️ Masih ada ' + cek.length + ' baris Musa');
}
