const DB_WILBIN_ID = '1JLQGTHvpInDIonogXSic6752lZjrQWvILily-xhfjXk';
const DB_POKTAN_ID = '1rfwl8IAhnxyYzNti1__IcZEdJ5iWUKuIJdVvC4aVYxs';

// Jumlah baris header per sheet Poktan.
// BKB punya 2 baris header (baris 1 judul kolom, baris 2 kosong) -> data mulai baris 3.
// BKR, BKL, UPPKA, PIKR hanya punya 1 baris header -> data mulai baris 2.
// PENTING: ini akar dari bug lama (index data-loop di-hardcode "2" untuk semua sheet,
// sehingga baris data pertama BKR/BKL/UPPKA/PIKR selalu terlewat).
const HEADER_ROW_COUNT = {
  BKB: 2,
  BKR: 1,
  BKL: 1,
  UPPKA: 1,
  PIKR: 1
};

function getHeaderRowCount_(namaSheet) {
  return HEADER_ROW_COUNT[namaSheet] !== undefined ? HEADER_ROW_COUNT[namaSheet] : 1;
}

// --- MANAJEMEN ROUTING (Ganti doGet lama dengan ini) ---
function doGet(e) {
  const page = e.parameter.page || 'dashboard'; // Default ke dashboard
  const scriptUrl = ScriptApp.getService().getUrl();
  
  let template;
  if (page === 'input') {
    template = HtmlService.createTemplateFromFile('Index');
    template.title = 'Input Konfirmasi Poktan NTT';
  } else {
    template = HtmlService.createTemplateFromFile('Dashboard');
    template.title = 'Dasbor Pemantauan Poktan NTT';
  }
  
  // Menyuntikkan URL dinamis ke frontend
  template.scriptUrl = scriptUrl; 
  
  return template.evaluate()
    .setTitle(template.title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- FUNGSI DATA DASBOR PUBLIK ---

// Dashboard publik ini bisa dibuka siapa saja kapan saja tanpa login, dan tiap
// pemanggilan tadinya selalu membaca ulang 5 sheet dari nol (~8.700 baris total).
// Supaya tidak berat & tidak boros kuota baca Spreadsheet, hasil agregasinya
// di-cache di server selama REKAP_CACHE_TTL_SECONDS. Karena ukurannya jauh di atas
// limit 100KB per key CacheService, datanya dipecah jadi beberapa "chunk" (lihat
// cachePutChunked_/cacheGetChunked_ di bawah).
const REKAP_CACHE_KEY = 'rekapDashboardPublik_v1';
const REKAP_CACHE_TTL_SECONDS = 300; // 5 menit

/**
 * Mengambil data teragregasi untuk dasbor publik (Berbasis Jenis Poktan & Wilayah).
 * @param {boolean} [forceRefresh] - true untuk melewati cache dan hitung ulang dari sheet
 *   (dipakai oleh tombol "Segarkan Data" di Dashboard.html).
 * @return {{data: Array, generatedAt: string}} data mentah + waktu data ini dihasilkan (ISO string).
 */
function getRekapDashboardPublik(forceRefresh) {
  const cache = CacheService.getScriptCache();

  if (!forceRefresh) {
    const cachedJson = cacheGetChunked_(cache, REKAP_CACHE_KEY);
    if (cachedJson) {
      try {
        return JSON.parse(cachedJson);
      } catch (e) {
        // Cache korup/tidak lengkap - lanjut hitung ulang dari sheet di bawah.
      }
    }
  }

  const result = {
    data: hitungRekapDashboardPublik_(),
    generatedAt: new Date().toISOString()
  };

  try {
    cachePutChunked_(cache, REKAP_CACHE_KEY, JSON.stringify(result), REKAP_CACHE_TTL_SECONDS);
  } catch (e) {
    // Kalau gagal menyimpan cache (misal ukuran data membengkak), tetap kembalikan
    // datanya - cache cuma optimisasi, bukan syarat supaya fitur ini berjalan.
  }

  return result;
}

function hitungRekapDashboardPublik_() {
  const sheets = ['BKB', 'BKR', 'BKL', 'UPPKA', 'PIKR'];
  const ssPoktan = SpreadsheetApp.openById(DB_POKTAN_ID);
  let allData = [];

  sheets.forEach(namaSheet => {
    const sheet = ssPoktan.getSheetByName(namaSheet);
    if (!sheet) return;
    
    const headerRows = getHeaderRowCount_(namaSheet);
    const colMap = getKolomIndexDinamis_(sheet, headerRows);
    const data = getRawSheetValues_(sheet, namaSheet);
    
    for (let i = headerRows; i < data.length; i++) {
      let row = data[i];
      let kab = String(row[3]).trim().toUpperCase(); // Kolom D: Nama Kabupaten
      let kec = String(row[colMap.kecamatan]).trim().toUpperCase();
      let desa = String(row[colMap.desa]).trim().toUpperCase();
      
      let rekomSistem = '-';
      if (row[colMap.pertahankan] == 1) rekomSistem = 'Dipertahankan';
      else if (row[colMap.penonaktifan] == 1) rekomSistem = 'Dinonaktifkan';

      let putusan = String(row[colMap.rtl] || '').trim();

      allData.push({
        jenisPoktan: namaSheet, // BKB, BKR, BKL, UPPKA, PIKR
        kabupaten: kab,
        kecamatan: kec,
        desa: desa, // dipakai tabel wilayah untuk drill-down sampai level desa
        rekomendasi: rekomSistem,
        putusan: putusan === '' ? 'Belum Input' : putusan
      });
    }
  });
  
  return allData;
}

// --- Util cache "chunked" (memecah string besar jadi beberapa key kecil) ---
// CacheService membatasi tiap value maksimal 100KB, sementara data dasbor publik
// bisa >1MB kalau di-JSON.stringify utuh. CHUNK_SIZE_ dijaga jauh di bawah limit itu.
const CACHE_CHUNK_SIZE_ = 90000;

function cachePutChunked_(cache, baseKey, str, ttlSeconds) {
  const totalChunks = Math.max(1, Math.ceil(str.length / CACHE_CHUNK_SIZE_));
  const payload = {};
  for (let i = 0; i < totalChunks; i++) {
    payload[baseKey + '_' + i] = str.substring(i * CACHE_CHUNK_SIZE_, (i + 1) * CACHE_CHUNK_SIZE_);
  }
  payload[baseKey + '_count'] = String(totalChunks);
  cache.putAll(payload, ttlSeconds);
}

function cacheGetChunked_(cache, baseKey) {
  const countStr = cache.get(baseKey + '_count');
  if (!countStr) return null;
  const count = parseInt(countStr, 10);
  if (!count || count <= 0) return null;

  const keys = [];
  for (let i = 0; i < count; i++) keys.push(baseKey + '_' + i);
  const parts = cache.getAll(keys);

  let result = '';
  for (let i = 0; i < count; i++) {
    const part = parts[baseKey + '_' + i];
    if (part === undefined || part === null) return null; // salah satu chunk sudah kedaluwarsa -> anggap cache miss semua
    result += part;
  }
  return result;
}

function cacheRemoveChunked_(cache, baseKey) {
  const countStr = cache.get(baseKey + '_count');
  if (!countStr) return;
  const count = parseInt(countStr, 10) || 0;
  const keys = [baseKey + '_count'];
  for (let i = 0; i < count; i++) keys.push(baseKey + '_' + i);
  cache.removeAll(keys);
}

// --- CACHE DATA MENTAH PER-SHEET POKTAN ---
// Bedanya dengan cache di getRekapDashboardPublik(): ini men-cache HASIL BACA MENTAH
// sheet.getDataRange().getValues() per sheet (BKB/BKR/BKL/UPPKA/PIKR), dipakai bersama
// oleh getPoktanData(), hitungRekapDashboardPublik_(), dan getProgresWilayahSaya().
// Tujuannya: kalau banyak PKB membuka tab yang sama dalam rentang beberapa detik
// (misalnya semua orang login serentak menjelang tenggat), Spreadsheet cuma benar-benar
// dibaca SEKALI - sisanya ambil dari cache. Ini mengurangi waktu eksekusi per
// panggilan (penting karena kuota "30 eksekusi bersamaan" Apps Script dibagi rata ke
// SEMUA pengunjung web app ini) sekaligus mengurangi total kuota baca harian.
//
// TTL sengaja pendek (90 detik) - cukup untuk meredam ledakan akses bersamaan, tapi
// tidak membuat data terasa basi. Begitu ada yang menyimpan perubahan ke satu sheet,
// cache sheet itu langsung dihapus paksa (lihat invalidateRawSheetCache_), jadi
// perubahan sendiri tetap langsung terlihat tanpa perlu menunggu TTL habis.
const RAW_SHEET_CACHE_PREFIX_ = 'rawSheet_v1_';
const RAW_SHEET_CACHE_TTL_SECONDS = 90;

function getRawSheetValues_(sheet, namaSheet) {
  const cache = CacheService.getScriptCache();
  const cacheKey = RAW_SHEET_CACHE_PREFIX_ + namaSheet;

  const cachedJson = cacheGetChunked_(cache, cacheKey);
  if (cachedJson) {
    try {
      return JSON.parse(cachedJson);
    } catch (e) {
      // Cache korup/tidak lengkap - lanjut baca ulang dari sheet di bawah.
    }
  }

  const data = sheet.getDataRange().getValues();
  try {
    cachePutChunked_(cache, cacheKey, JSON.stringify(data), RAW_SHEET_CACHE_TTL_SECONDS);
  } catch (e) {
    // Gagal simpan cache (misal data sedang sangat besar) bukan masalah besar -
    // datanya tetap dikembalikan, cuma tidak sempat di-cache kali ini.
  }
  return data;
}

function invalidateRawSheetCache_(namaSheet) {
  const cache = CacheService.getScriptCache();
  cacheRemoveChunked_(cache, RAW_SHEET_CACHE_PREFIX_ + namaSheet);
}

// --- PROTEKSI PERCOBAAN LOGIN BERULANG (BRUTE-FORCE) ---
// Bukan pengganti otentikasi yang kuat (itu di luar cakupan proyek ini), tapi menambah
// friksi supaya menebak-nebak NIP/passcode tidak bisa dilakukan tanpa batas. Dilacak
// per "kunci" via CacheService - nama PKB yang dicoba, kabupaten Dinas yang dicoba,
// atau satu kunci global untuk Coach & Mentor (karena passcode-nya memang tunggal) -
// supaya kegagalan pada satu identitas tidak ikut mengunci identitas lain yang tidak
// terkait.
const MAX_PERCOBAAN_GAGAL_LOGIN = 5;
const JENDELA_KUNCI_LOGIN_DETIK = 300; // 5 menit terkunci setelah gagal berulang

function cekBatasPercobaanGagal_(kunci) {
  const cache = CacheService.getScriptCache();
  const countStr = cache.get('gagalLogin_' + kunci);
  const count = countStr ? parseInt(countStr, 10) : 0;
  if (count >= MAX_PERCOBAAN_GAGAL_LOGIN) {
    throw new Error("Terlalu banyak percobaan gagal. Coba lagi dalam beberapa menit.");
  }
}

function catatPercobaanGagal_(kunci) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'gagalLogin_' + kunci;
  const countStr = cache.get(cacheKey);
  const count = countStr ? parseInt(countStr, 10) : 0;
  cache.put(cacheKey, String(count + 1), JENDELA_KUNCI_LOGIN_DETIK);
}

function resetPercobaanGagal_(kunci) {
  CacheService.getScriptCache().remove('gagalLogin_' + kunci);
}



// --- 1. OTENTIKASI & REFERENSI WILAYAH ---

function getDaftarKabupaten() {
  const ss = SpreadsheetApp.openById(DB_WILBIN_ID);
  const data = ss.getSheetByName('datapkb').getDataRange().getValues();
  data.shift();
  const kab = [...new Set(data.map(r => r[0]).filter(Boolean))];
  return kab.sort();
}

function getDaftarNama(kabupaten) {
  const ss = SpreadsheetApp.openById(DB_WILBIN_ID);
  const data = ss.getSheetByName('datapkb').getDataRange().getValues();
  data.shift();
  return data.filter(r => r[0] === kabupaten && r[1]).map(r => r[1]).sort();
}

/**
 * Validasi login PKB/PLKB berdasarkan NIP dan wilayah binaannya.
 */
function loginDanGetWilayah(nama, nip) {
  // Dikunci per NAMA yang dicoba, bukan global - supaya orang lain yang login dengan
  // nama berbeda tidak ikut terkena imbas kalau satu nama sedang ditebak-tebak NIP-nya.
  const kunciPercobaan = 'pkb_' + nama.toString().trim().toUpperCase();
  cekBatasPercobaanGagal_(kunciPercobaan);

  const ssWilbin = SpreadsheetApp.openById(DB_WILBIN_ID);
  const targetNIP = nip.toString().trim();
  
  const dataPkb = ssWilbin.getSheetByName('datapkb').getDataRange().getValues();
  const validUser = dataPkb.find(r => r[1] === nama && r[2] && r[2].toString().trim() === targetNIP);
  
  if (!validUser) {
    catatPercobaanGagal_(kunciPercobaan);
    throw new Error("Nama atau NIP PKB/PLKB tidak valid.");
  }
  resetPercobaanGagal_(kunciPercobaan);
  
  let wilayahBinaan = [];
  let status = 'BELUM_INPUT';
  
  const sheetHasil = ssWilbin.getSheetByName('HasilPendataan');
  if (sheetHasil && sheetHasil.getLastRow() > 1) {
    const dataHasil = sheetHasil.getRange(2, 4, sheetHasil.getLastRow() - 1, 3).getValues();
    const filterHasil = dataHasil.filter(r => r[0] && r[0].toString().trim() === targetNIP);
    
    if (filterHasil.length > 0) {
      status = 'SUDAH_INPUT';
      filterHasil.forEach(r => {
        let kec = r[1] ? r[1].toString().trim().toUpperCase() : '';
        let desaStr = r[2] ? r[2].toString().trim() : '';
        let arrDesa = desaStr.split(',').map(d => d.trim().toUpperCase()).filter(Boolean);
        wilayahBinaan.push({ kecamatan: kec, desa: arrDesa });
      });
      return { tipe: 'PKB', nip: targetNIP, nama: nama, kabupaten: validUser[0], status: status, wilayah: wilayahBinaan };
    }
  }
  
  let kecDefault = validUser[3] ? validUser[3].toString().trim().toUpperCase() : '';
  wilayahBinaan.push({ kecamatan: kecDefault, desa: [] });
  return { tipe: 'PKB', nip: targetNIP, nama: nama, kabupaten: validUser[0], status: status, wilayah: wilayahBinaan };
}

/**
 * Inisialisasi akses untuk Dinas OPD KB (mencakup seluruh kecamatan di Kabupaten tersebut).
 */
/**
 * Cari passcode Dinas OPD untuk satu kabupaten di sheet "dataDinasPasscode" (DB_WILBIN_ID).
 * Mengembalikan null kalau sheet belum dibuat, kabupaten belum terdaftar, atau kolom
 * passcode-nya masih kosong - supaya AMAN SECARA DEFAULT (kabupaten yang belum diatur
 * pasti ditolak, bukan malah kebuka).
 */
function getPasscodeDinasUntukKabupaten_(kabupaten) {
  const ssWilbin = SpreadsheetApp.openById(DB_WILBIN_ID);
  const sheet = ssWilbin.getSheetByName('dataDinasPasscode');
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const targetKab = kabupaten.toString().trim().toUpperCase();

  for (let i = 1; i < data.length; i++) {
    const kab = String(data[i][0] || '').trim().toUpperCase();
    if (kab === targetKab) {
      const passcode = data[i][1] !== undefined && data[i][1] !== null ? String(data[i][1]).trim() : '';
      return passcode || null;
    }
  }
  return null;
}

function loginDinasOPD(kabupaten, dataDinas, passcode) {
  if (!kabupaten || !dataDinas.nama || !dataDinas.nip) {
    throw new Error("Mohon lengkapi Nama, NIP, Jabatan, dan Nomor WA.");
  }

  // Passcode KHUSUS kabupaten yang dipilih - satu kabupaten bocor tidak berarti semua
  // kabupaten ikut kebobolan (beda dengan satu passcode tunggal untuk semua kabupaten).
  const passcodeValid = getPasscodeDinasUntukKabupaten_(kabupaten);
  if (!passcodeValid) {
    throw new Error(
      "Akses Dinas OPD untuk kabupaten " + kabupaten + " belum diaktifkan oleh admin. " +
      "Hubungi admin untuk mengatur passcode kabupaten ini di sheet dataDinasPasscode."
    );
  }

  // Dikunci per KABUPATEN yang dicoba - ditaruh setelah cek di atas supaya kabupaten
  // yang belum diatur admin tidak ikut kena hitungan "percobaan gagal".
  const kunciPercobaan = 'dinas_' + kabupaten.toString().trim().toUpperCase();
  cekBatasPercobaanGagal_(kunciPercobaan);

  if (!passcode || passcode.toString().trim() !== passcodeValid) {
    catatPercobaanGagal_(kunciPercobaan);
    throw new Error("Passcode kabupaten salah.");
  }
  resetPercobaanGagal_(kunciPercobaan);
  
  // Dinas OPD KB dapat mengakses seluruh kecamatan di kabupaten pilihannya
  const ssWilbin = SpreadsheetApp.openById(DB_WILBIN_ID);
  const sheetMaster = ssWilbin.getSheetByName('masterWilayah');
  const dataMaster = sheetMaster.getDataRange().getValues();
  
  const setKec = new Set();
  for (let i = 1; i < dataMaster.length; i++) {
    if (dataMaster[i][0] === kabupaten && dataMaster[i][1]) {
      setKec.add(dataMaster[i][1].toString().trim().toUpperCase());
    }
  }
  
  let wilayahBinaan = [{ kecamatan: 'SEMUA', desa: [] }];
  if (setKec.size > 0) {
    wilayahBinaan = Array.from(setKec).map(k => ({ kecamatan: k, desa: [] }));
  }
  
  return {
    tipe: 'DINAS',
    nama: dataDinas.nama,
    nip: dataDinas.nip,
    jabatan: dataDinas.jabatan,
    wa: dataDinas.wa,
    kabupaten: kabupaten,
    status: 'DINAS_OPD',
    wilayah: wilayahBinaan
  };
}


// --- 2. PEMROSESAN DATA POKTAN (Dynamic Mapping) ---

/**
 * Mendeteksi index kolom berdasarkan teks header.
 * headerRows menentukan berapa baris paling atas yang benar-benar berisi judul kolom
 * (BUKAN data) - lihat konstanta HEADER_ROW_COUNT di atas. Sebelumnya fungsi ini selalu
 * membaca 2 baris (getRange(1,1,2,...)), sehingga untuk sheet yang cuma punya 1 baris
 * header, baris data pertama ikut "dianggap" bagian dari teks header.
 */
function getKolomIndexDinamis_(sheet, headerRows) {
  if (headerRows === undefined) headerRows = 2; // fallback ke perilaku lama jika tidak dikirim
  const data = sheet.getRange(1, 1, headerRows, sheet.getLastColumn()).getValues();
  const colMap = {};
  
  for (let i = 0; i < data[0].length; i++) {
    let headerStr = '';
    for (let r = 0; r < headerRows; r++) {
      headerStr += String(data[r][i]) + ' ';
    }
    headerStr = headerStr.toLowerCase();
    
    if (headerStr.includes('nama kecamatan')) colMap.kecamatan = i;
    else if (headerStr.includes('nama desa')) colMap.desa = i;
    else if (headerStr.includes('id poktan')) colMap.id = i;
    else if (headerStr.includes('nama poktan')) colMap.nama = i;
    else if (headerStr.includes('rencana tindak lanjut') || headerStr.includes('tindak lanjut')) colMap.rtl = i;
    else if (headerStr.includes('alasan') || headerStr.includes('justifikasi')) colMap.alasan = i;
    else if (headerStr.includes('pertahankan')) colMap.pertahankan = i;
    else if (headerStr.includes('perbaikan')) colMap.perbaikan = i;
    else if (headerStr.includes('penonaktifan')) colMap.penonaktifan = i;
  }
  return colMap;
}

/**
 * Menentukan apakah satu baris data poktan (kabupaten/kecamatan/desa) termasuk wilayah
 * binaan authData yang sedang login - dipakai bersama oleh getPoktanData() dan
 * getProgresWilayahSaya(), supaya logika "milik siapa wilayah ini" hanya ada di satu tempat.
 */
function cocokWilayah_(authData, recKab, recKec, recDesa) {
  if (authData.tipe === 'DINAS') {
    // Dinas OPD KB memfilter berdasarkan Kabupaten di baris data Poktan
    return recKab.toUpperCase() === authData.kabupaten.toUpperCase();
  }
  // PKB/PLKB mencocokkan kecamatan/desa binaan
  for (let w of authData.wilayah) {
    if (w.kecamatan === recKec) {
      if (authData.status === 'BELUM_INPUT' || w.desa.length === 0 || w.desa.includes(recDesa)) {
        return true;
      }
    }
  }
  return false;
}

function getPoktanData(authData, jenisPoktan) {
  const ssPoktan = SpreadsheetApp.openById(DB_POKTAN_ID);
  const sheet = ssPoktan.getSheetByName(jenisPoktan);
  if (!sheet) throw new Error("Sheet " + jenisPoktan + " tidak ditemukan.");
  
  const headerRows = getHeaderRowCount_(jenisPoktan);
  const colMap = getKolomIndexDinamis_(sheet, headerRows);
  const data = getRawSheetValues_(sheet, jenisPoktan);
  let result = [];
  
  for (let i = headerRows; i < data.length; i++) {
    let row = data[i];
    let recKec = String(row[colMap.kecamatan]).trim().toUpperCase();
    let recDesa = String(row[colMap.desa]).trim().toUpperCase();
    let recKab = String(row[3]).trim(); // Kolom ke-4 diasumsikan Nama Kabupaten di file Konfirmasi

    let isMatch = cocokWilayah_(authData, recKab, recKec, recDesa);
    
    if (isMatch) {
      let rekom = '-';
      if (row[colMap.pertahankan] == 1) rekom = 'Dipertahankan';
      else if (row[colMap.penonaktifan] == 1) rekom = 'Dinonaktifkan';
      
      result.push({
        rowIdx: i + 1,
        idPoktan: row[colMap.id],
        namaPoktan: row[colMap.nama],
        desa: String(row[colMap.desa]).trim(),
        kecamatan: recKec, // ditambahkan agar frontend bisa filter/paginasi per kecamatan (berguna untuk Dinas OPD)
        rekomendasi: rekom,
        rtl: row[colMap.rtl] || '',
        alasan: row[colMap.alasan] || ''
      });
    }
  }
  
  return { kolomRTL: colMap.rtl + 1, kolomAlasan: colMap.alasan + 1, data: result };
}

function simpanFeedbackPoktan(authData, jenisPoktan, payload) {
  // Validasi bisnis: setiap baris yang RTL-nya sudah diisi WAJIB disertai alasan
  // (mayoritas rekomendasi sistem adalah "Dinonaktifkan", jadi keputusan apapun -
  // termasuk "Dipertahankan" - harus punya justifikasi tertulis).
  // Divalidasi ulang di server (bukan cuma di frontend) supaya tidak bisa dilewati.
  const invalid = (payload.updates || []).filter(u =>
    u.rtl && u.rtl.toString().trim() !== '' &&
    (!u.alasan || u.alasan.toString().trim() === '')
  );
  if (invalid.length > 0) {
    throw new Error(
      "Ada " + invalid.length + " baris dengan Rencana Tindak Lanjut terisi namun alasan masih kosong. " +
      "Alasan wajib diisi sebelum data bisa disimpan."
    );
  }

  // Serialisasi bagian TULIS dengan LockService - supaya kalau dua orang (misal PKB dan
  // Dinas OPD) klik "Simpan" di waktu yang nyaris sama, mereka dieksekusi satu-satu
  // (antre), bukan tabrakan. Ini juga memperbaiki bug nyata di catatLogInput_(): tanpa
  // lock, dua eksekusi bisa sama-sama membaca sheet.getLastRow() sebelum salah satu
  // sempat menulis, sehingga log yang satu MENIMPA log yang lain di baris yang sama.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // tunggu maksimal 15 detik giliran
  } catch (e) {
    throw new Error("Sistem sedang memproses simpanan lain. Coba simpan lagi dalam beberapa detik.");
  }

  try {
    const ssPoktan = SpreadsheetApp.openById(DB_POKTAN_ID);
    const sheet = ssPoktan.getSheetByName(jenisPoktan);
    if (!sheet) throw new Error("Sheet " + jenisPoktan + " tidak ditemukan.");

    // Tulis perubahan secara batch (menghindari timeout).
    // Catatan: payload.updates sekarang hanya berisi baris yang benar-benar diubah
    // di sisi client (lihat perubahan simpanPerubahan() di Index.html), sehingga
    // baris yang tidak disentuh oleh sesi ini tidak ikut ditimpa.
    //
    // Sekalian kumpulkan data untuk LOG RIWAYAT INPUT (siapa mengubah apa, kapan) -
    // dicatat di sheet terpisah "LogInput", BUKAN dengan menambah kolom di sheet
    // BKB/BKR/BKL/UPPKA/PIKR. Alasannya: struktur kolom tiap sheet berbeda-beda dan
    // beberapa formula rekap kemungkinan bergantung pada posisi kolom yang ada
    // sekarang, jadi menyisipkan kolom baru berisiko merusak itu. Sheet log terpisah
    // dengan skema kita sendiri jauh lebih aman dan sekaligus jadi riwayat perubahan
    // (bukan cuma "siapa terakhir edit").
    const logEntries = [];

    payload.updates.forEach(u => {
      let rtlSebelum = '';
      if (u.rtl !== undefined) {
        rtlSebelum = sheet.getRange(u.rowIdx, payload.kolomRTL).getValue();
        sheet.getRange(u.rowIdx, payload.kolomRTL).setValue(u.rtl);
      }
      if (u.alasan !== undefined) {
        sheet.getRange(u.rowIdx, payload.kolomAlasan).setValue(u.alasan);
      }
      logEntries.push({
        idPoktan: u.idPoktan,
        namaPoktan: u.namaPoktan,
        kecamatan: u.kecamatan,
        rtlSebelum: rtlSebelum,
        rtlSesudah: u.rtl,
        alasanSesudah: u.alasan
      });
    });

    try {
      catatLogInput_(ssPoktan, jenisPoktan, authData, logEntries);
    } catch (e) {
      // Kalau pencatatan log gagal, jangan sampai menggagalkan penyimpanan data utama -
      // ini cuma jejak audit, bukan syarat inti fitur simpan.
      console.error('Gagal menulis LogInput: ' + e.message);
    }

    // Data sheet ini baru saja berubah - hapus cache mentahnya supaya pembacaan
    // berikutnya (termasuk refresh KPI "Progres Wilayah Anda" milik orang yang baru
    // saja simpan) tidak melihat angka basi sampai TTL cache habis.
    invalidateRawSheetCache_(jenisPoktan);

    return "Data " + jenisPoktan + " berhasil diperbarui (" + payload.updates.length + " baris).";
  } finally {
    lock.releaseLock();
  }
}

// --- LOG RIWAYAT INPUT (sheet "LogInput", dibuat otomatis kalau belum ada) ---

const LOG_SHEET_NAME = 'LogInput';
const LOG_HEADERS_ = [
  'Waktu', 'Jenis Poktan', 'ID Poktan', 'Nama Poktan', 'Kabupaten', 'Kecamatan',
  'Tipe Penginput', 'Nama Penginput', 'NIP Penginput',
  'RTL Sebelum', 'RTL Sesudah', 'Alasan Sesudah'
];

function getOrCreateLogSheet_(ss) {
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.getRange(1, 1, 1, LOG_HEADERS_.length).setValues([LOG_HEADERS_]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Menulis satu baris log untuk tiap baris poktan yang diubah pada satu kali "Simpan".
 * @param {Spreadsheet} ss - spreadsheet DB_POKTAN_ID (log ditaruh di file yang sama).
 * @param {string} jenisPoktan - BKB/BKR/BKL/UPPKA/PIKR
 * @param {Object} authData - hasil login (punya .tipe, .nama, .nip, .kabupaten)
 * @param {Array} logEntries - [{idPoktan, namaPoktan, kecamatan, rtlSebelum, rtlSesudah, alasanSesudah}]
 */
function catatLogInput_(ss, jenisPoktan, authData, logEntries) {
  if (!logEntries || logEntries.length === 0) return;

  const sheet = getOrCreateLogSheet_(ss);
  const now = new Date();
  const namaPenginput = authData && authData.nama ? authData.nama : '';
  const nipPenginput = authData && authData.nip ? authData.nip : '';
  const tipePenginput = authData && authData.tipe ? authData.tipe : '';
  const kabupaten = authData && authData.kabupaten ? authData.kabupaten : '';

  const rows = logEntries.map(e => [
    now,
    jenisPoktan,
    e.idPoktan || '',
    e.namaPoktan || '',
    kabupaten,
    e.kecamatan || '',
    tipePenginput,
    namaPenginput,
    nipPenginput,
    e.rtlSebelum || '',
    e.rtlSesudah || '',
    e.alasanSesudah || ''
  ]);

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, LOG_HEADERS_.length).setValues(rows);
}

// --- TAB "PROGRES COACH & MENTOR" (Dashboard.html) ---

// Nama Script Property tempat passcode disimpan. Admin mengaturnya lewat:
// Apps Script Editor > Project Settings (ikon gerigi) > Script Properties > Add script property
// key = COACH_MENTOR_PASSCODE, value = passcode yang diinginkan.
// Passcode bisa diganti kapan saja lewat sana TANPA perlu edit kode / deploy ulang.
const COACH_MENTOR_PASSCODE_PROPERTY = 'COACH_MENTOR_PASSCODE';

function cekPasscodeCoachMentor_(passcode) {
  const stored = PropertiesService.getScriptProperties().getProperty(COACH_MENTOR_PASSCODE_PROPERTY);
  if (!stored) {
    throw new Error(
      'Passcode untuk halaman Coach & Mentor belum diatur oleh admin. ' +
      'Atur dulu di Apps Script > Project Settings > Script Properties, key: ' + COACH_MENTOR_PASSCODE_PROPERTY
    );
  }

  // Satu kunci global karena passcode-nya memang satu untuk semua Mentor/Coach -
  // ditaruh setelah cek di atas supaya kondisi "belum diatur admin" tidak ikut
  // dihitung sebagai percobaan gagal.
  const kunciPercobaan = 'coachMentor';
  cekBatasPercobaanGagal_(kunciPercobaan);

  if (!passcode || passcode.toString() !== stored.toString()) {
    catatPercobaanGagal_(kunciPercobaan);
    throw new Error('Passcode salah.');
  }
  resetPercobaanGagal_(kunciPercobaan);
}

/**
 * Validasi passcode, lalu kembalikan peta penugasan Mentor & Coach per Kabupaten/Kecamatan
 * dari sheet "dataCoach" (DB_WILBIN_ID) - spreadsheet yang berbeda dari data poktan.
 * Frontend (Dashboard.html) menggabungkan (join) hasil ini dengan rawData yang sudah
 * dimuat publik lewat getRekapDashboardPublik(), jadi tidak perlu ambil ulang seluruh
 * data poktan di sini.
 *
 * Catatan: kolom "MENTOR 2" di sheet dataCoach sengaja diabaikan karena saat ini
 * selalu identik dengan kolom "MENTOR" - mudah ditambahkan lagi kalau nanti dipakai.
 */
function getPetaCoachMentor(passcode) {
  cekPasscodeCoachMentor_(passcode);

  const ssWilbin = SpreadsheetApp.openById(DB_WILBIN_ID);
  const sheet = ssWilbin.getSheetByName('dataCoach');
  if (!sheet) throw new Error('Sheet dataCoach tidak ditemukan.');

  const data = sheet.getDataRange().getValues();
  // Header baris 1: KAB/KOTA, KECAMATAN, MENTOR, MENTOR 2, COACH
  const assignments = [];
  const mentorSet = new Set();
  const coachSet = new Set();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const kab = String(row[0] || '').trim().toUpperCase();
    const kec = String(row[1] || '').trim().toUpperCase();
    const mentor = String(row[2] || '').trim();
    const coach = String(row[4] || '').trim();
    if (!kab || !kec) continue;

    assignments.push({ kabupaten: kab, kecamatan: kec, mentor: mentor, coach: coach });
    if (mentor) mentorSet.add(mentor);
    if (coach) coachSet.add(coach);
  }

  return {
    assignments: assignments,
    mentors: Array.from(mentorSet).sort(),
    coaches: Array.from(coachSet).sort()
  };
}

// --- "PROGRES WILAYAH ANDA" (Index.html) ---

/**
 * Menjumlahkan progres SELURUH jenis poktan (BKB, BKR, BKL, UPPKA, PIKR) yang termasuk
 * wilayah binaan authData yang sedang login - dipakai untuk KPI "Progres Wilayah Anda"
 * di Panel Input. Beda dengan getPoktanData() yang cuma satu sektor & mengembalikan
 * seluruh baris; di sini cukup hitung angkanya saja (jauh lebih ringan), dan mencakup
 * semua sektor sekaligus karena satu PKB/PLKB bertanggung jawab atas kelimanya di
 * wilayahnya, bukan cuma satu sektor yang sedang dibuka.
 */
function getProgresWilayahSaya(authData) {
  const sheets = ['BKB', 'BKR', 'BKL', 'UPPKA', 'PIKR'];
  const ssPoktan = SpreadsheetApp.openById(DB_POKTAN_ID);

  const perSektor = [];
  let totalTarget = 0, totalDipertahankan = 0, totalDinonaktifkan = 0, totalBedaRekom = 0;

  sheets.forEach(namaSheet => {
    const sheet = ssPoktan.getSheetByName(namaSheet);
    if (!sheet) { perSektor.push({ jenisPoktan: namaSheet, target: 0, selesai: 0 }); return; }

    const headerRows = getHeaderRowCount_(namaSheet);
    const colMap = getKolomIndexDinamis_(sheet, headerRows);
    const data = getRawSheetValues_(sheet, namaSheet);

    let target = 0, selesai = 0;
    for (let i = headerRows; i < data.length; i++) {
      const row = data[i];
      const recKab = String(row[3]).trim();
      const recKec = String(row[colMap.kecamatan]).trim().toUpperCase();
      const recDesa = String(row[colMap.desa]).trim().toUpperCase();

      if (!cocokWilayah_(authData, recKab, recKec, recDesa)) continue;

      target++;
      totalTarget++;

      const rtl = String(row[colMap.rtl] || '').trim();
      if (rtl !== '') {
        selesai++;
        if (rtl === 'Dipertahankan') totalDipertahankan++;
        else if (rtl === 'Dinonaktifkan') totalDinonaktifkan++;

        let rekomSistem = '-';
        if (row[colMap.pertahankan] == 1) rekomSistem = 'Dipertahankan';
        else if (row[colMap.penonaktifan] == 1) rekomSistem = 'Dinonaktifkan';
        if (rekomSistem === 'Dinonaktifkan' && rtl === 'Dipertahankan') totalBedaRekom++;
      }
    }
    perSektor.push({ jenisPoktan: namaSheet, target: target, selesai: selesai });
  });

  return {
    totalTarget: totalTarget,
    totalSelesai: totalDipertahankan + totalDinonaktifkan,
    totalDipertahankan: totalDipertahankan,
    totalDinonaktifkan: totalDinonaktifkan,
    totalBedaRekom: totalBedaRekom,
    perSektor: perSektor
  };
}
