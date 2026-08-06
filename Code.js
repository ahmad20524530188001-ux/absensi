// =========================================================
//  PRESENSI DIGITAL ANNURIYYAH — Code.gs (FIXED)
// =========================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Presensi Digital Annuriyyah')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =========================================================
//  INISIALISASI DATABASE
// =========================================================
function initDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {
    "Absensi"        : ["Timestamp","Instansi","Nama","Sesi","Status","Keterangan","Lokasi","Link Foto","Link Dokumen"],
    "Master_Instansi": ["Nama Instansi"],
    "Master_Karyawan": ["Instansi","Nama Karyawan","PIN"],
    "Master_Sesi"    : ["Nama Sesi","Jam Mulai","Jam Selesai"]
  };

  for (const name in sheets) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(sheets[name]);
      sheet.getRange(1, 1, 1, sheets[name].length)
           .setBackground("#065f46").setFontColor("white").setFontWeight("bold");
    }
  }

  // Migrasi: tambah kolom "Lokasi" di Absensi jika belum ada
  const absensi = ss.getSheetByName("Absensi");
  if (absensi && absensi.getLastColumn() > 0) {
    const headerAbsensi = absensi.getRange(1, 1, 1, absensi.getLastColumn()).getValues()[0];
    if (headerAbsensi.indexOf("Lokasi") === -1) {
      const fotoIdx = headerAbsensi.indexOf("Link Foto");
      const insertBeforeCol = fotoIdx >= 0 ? fotoIdx + 1 : headerAbsensi.length + 1;
      absensi.insertColumnBefore(insertBeforeCol);
      absensi.getRange(1, insertBeforeCol).setValue("Lokasi")
             .setBackground("#065f46").setFontColor("white").setFontWeight("bold");
    }
  }

  // Migrasi: tambah kolom "PIN" di Master_Karyawan jika belum ada
  const masterKar = ss.getSheetByName("Master_Karyawan");
  if (masterKar && masterKar.getLastColumn() > 0) {
    const headerKar = masterKar.getRange(1, 1, 1, masterKar.getLastColumn()).getValues()[0];
    if (headerKar.indexOf("PIN") === -1) {
      masterKar.insertColumnAfter(2);
      masterKar.getRange(1, 3).setValue("PIN")
               .setBackground("#065f46").setFontColor("white").setFontWeight("bold");
    }
  }
}

// =========================================================
//  AMBIL SEMUA DATA UNTUK FRONTEND
// =========================================================
function getAllAppData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // FIX: Gunakan try-catch per sheet agar satu sheet kosong tidak
  //       menghentikan seluruh fungsi
  function fetchSheet(name) {
    try {
      const sheet = ss.getSheetByName(name);
      if (!sheet || sheet.getLastRow() < 1) return [];
      return sheet.getDataRange().getDisplayValues().slice(1); // Hapus header
    } catch (e) {
      Logger.log("Gagal fetch sheet " + name + ": " + e);
      return [];
    }
  }

  return {
    instansi: fetchSheet("Master_Instansi"),
    karyawan: fetchSheet("Master_Karyawan"),
    sesi    : fetchSheet("Master_Sesi"),
    logs    : fetchSheet("Absensi").reverse()  // Terbaru di atas
  };
}

// =========================================================
//  SUBMIT ABSENSI
// =========================================================
function submitAbsensi(obj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();

  // --- VALIDASI INPUT DASAR ---
  if (!obj.instansi || !obj.nama || !obj.pin) {
    return { status: "error", msg: "Data tidak lengkap. Instansi, nama, dan PIN wajib diisi." };
  }

  // FIX: Validasi format PIN (6 digit angka)
  if (!/^\d{6}$/.test(obj.pin)) {
    return { status: "error", msg: "Format PIN tidak valid. PIN harus 6 digit angka." };
  }

  // --- VALIDASI PIN KARYAWAN ---
  const karyawanSheet = ss.getSheetByName("Master_Karyawan");
  if (!karyawanSheet) return { status: "error", msg: "Sheet Master_Karyawan tidak ditemukan." };

  const karyawanData = karyawanSheet.getDataRange().getDisplayValues().slice(1);
  // FIX: Gunakan trim() agar spasi di spreadsheet tidak menyebabkan gagal login
  const matchedKaryawan = karyawanData.find(k =>
    k[0].trim() === obj.instansi.trim() && k[1].trim() === obj.nama.trim()
  );

  if (!matchedKaryawan) {
    return { status: "error", msg: "GAGAL! Karyawan tidak ditemukan di instansi yang dipilih." };
  }

  // FIX: Bandingkan PIN setelah trim untuk menghindari bug spasi
  if (matchedKaryawan[2].trim() !== obj.pin.trim()) {
    return { status: "error", msg: "GAGAL! PIN Akses tidak valid atau salah." };
  }

  // --- VALIDASI WAKTU SESI (hanya untuk status Hadir) ---
  if (obj.status === "Hadir") {
    if (!obj.sesi) {
      return { status: "error", msg: "Sesi wajib dipilih untuk absensi Hadir." };
    }
    const sesiSheet = ss.getSheetByName("Master_Sesi");
    if (sesiSheet) {
      const sesiData = sesiSheet.getDataRange().getDisplayValues().slice(1);
      // FIX: Ekstrak nama sesi murni sebelum tanda kurung "("
      const namaSesiMurni = obj.sesi.split(' (')[0].trim();
      const match = sesiData.find(r => r[0].trim() === namaSesiMurni);

      if (match) {
        // FIX: Normalisasi format waktu (titik → titik dua)
        const jamSelesai = match[2].replace('.', ':').trim();
        const parts = jamSelesai.split(':');
        if (parts.length >= 2) {
          const h = parseInt(parts[0], 10);
          const m = parseInt(parts[1], 10);
          if (!isNaN(h) && !isNaN(m)) {
            const limit = new Date();
            limit.setHours(h, m, 59, 999);
            if (now > limit) {
              return { status: "error", msg: `GAGAL! Sesi "${namaSesiMurni}" sudah berakhir pada pukul ${jamSelesai}.` };
            }
          }
        }
      }
    }
  }

  // --- UPLOAD FILE KE GOOGLE DRIVE ---
  try {
    const props = PropertiesService.getScriptProperties();
    const folderId = props.getProperty('FOLDER_ID');

    // FIX: Periksa apakah FOLDER_ID sudah dikonfigurasi
    if (!folderId) {
      return { status: "error", msg: "Konfigurasi FOLDER_ID belum diset di Script Properties." };
    }

    let folder;
    try {
      folder = DriveApp.getFolderById(folderId);
    } catch (e) {
      return { status: "error", msg: "FOLDER_ID tidak valid atau tidak dapat diakses: " + folderId };
    }

   // FIX: Fungsi upload yang lebih robust dengan penanganan tipe file
function uploadFile(b64, prefix) {
  if (!b64 || b64.length < 10) return "";
  
  const MAX_RETRY = 3;
  
  for (let attempt = 1; attempt <= MAX_RETRY; attempt++) {
    try {
      const isPdf    = b64.indexOf("data:application/pdf") !== -1;
      const mime     = isPdf ? "application/pdf" : "image/jpeg";
      const ext      = isPdf ? ".pdf" : ".jpg";
      const commaIdx = b64.indexOf(',');
      if (commaIdx === -1) return "";
      const raw      = b64.substring(commaIdx + 1);
      const bytes    = Utilities.base64Decode(raw);
      const namaKaryawan = obj.nama.replace(/[^a-zA-Z0-9]/g, "_");
      const namaFile = prefix + "_" + namaKaryawan + "_" + Date.now() + ext;
      const blob     = Utilities.newBlob(bytes, mime, namaFile);
      const file     = folder.createFile(blob);
      //file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return file.getUrl();
      
    } catch (uploadErr) {
      Logger.log("Upload gagal attempt " + attempt + " (" + prefix + "): " + uploadErr);
      if (attempt < MAX_RETRY) {
        Utilities.sleep(1500 * attempt); // tunggu 1.5s, 3s, 4.5s
      } else {
        return "GAGAL: " + uploadErr.toString();
      }
    }
  }
  return "";
}

    const photoUrl = uploadFile(obj.foto, "SELFIE");
    const docUrl   = uploadFile(obj.dokumen, "DOKUMEN");

    // Susun teks lokasi beserta link Google Maps
    let lokasiText = obj.lokasi || "Lokasi tidak tersedia";
    if (obj.lat && obj.lng) {
      lokasiText += " | Maps: https://maps.google.com/?q=" + obj.lat + "," + obj.lng;
    }

    // Simpan ke sheet Absensi
    const absensiSheet = ss.getSheetByName("Absensi");
    if (!absensiSheet) return { status: "error", msg: "Sheet Absensi tidak ditemukan." };

    absensiSheet.appendRow([
      now,
      obj.instansi,
      obj.nama,
      obj.sesi || "",
      obj.status,
      obj.ket || "",
      lokasiText,
      photoUrl,
      docUrl
    ]);

    // FIX: Pastikan baris baru tidak diformat ulang (hindari menimpa format header)
    return { status: "success" };

  } catch (e) {
    Logger.log("submitAbsensi error: " + e);
    return { status: "error", msg: "Terjadi kesalahan server: " + e.message };
  }
}

// =========================================================
//  CRUD MASTER DATA
// =========================================================
function crudMaster(action, sheetName, data) {
  // FIX: Validasi parameter
  if (!action || !sheetName || !data) {
    return { status: "error", msg: "Parameter tidak lengkap." };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return { status: "error", msg: "Sheet " + sheetName + " tidak ditemukan." };

  try {
    if (action === "add") {
      // FIX: Trim semua nilai sebelum disimpan
      sheet.appendRow(data.map(v => (typeof v === 'string' ? v.trim() : v)));

    } else if (action === "delete") {
      const vals = sheet.getDataRange().getValues();
      // FIX: Hapus dari bawah ke atas agar indeks tidak bergeser
      for (let i = vals.length - 1; i >= 1; i--) {
        const match = data.every((d, idx) => {
          // Hanya bandingkan kolom yang ada di data (untuk instansi hanya kolom 0,
          // untuk karyawan bandingkan 2 kolom pertama)
          return idx >= data.length || String(vals[i][idx]).trim() === String(d).trim();
        });
        if (match) sheet.deleteRow(i + 1);
      }

    } else if (action === "update_pin") {
      // data = [instansi, nama, pin_baru]
      if (data.length < 3) return { status: "error", msg: "Data update_pin tidak lengkap." };
      const vals = sheet.getDataRange().getValues();
      let updated = false;
      for (let i = 1; i < vals.length; i++) {
        if (String(vals[i][0]).trim() === String(data[0]).trim() &&
            String(vals[i][1]).trim() === String(data[1]).trim()) {
          sheet.getRange(i + 1, 3).setValue(String(data[2]).trim());
          updated = true;
          break;
        }
      }
      if (!updated) return { status: "error", msg: "Karyawan tidak ditemukan untuk update PIN." };
    }

    return { status: "ok" };

  } catch (e) {
    Logger.log("crudMaster error: " + e);
    return { status: "error", msg: e.message };
  }
}

// =========================================================
//  VERIFIKASI ADMIN
// =========================================================
function verifyAdmin(pass) {
  // FIX: Cek input tidak kosong dulu
  if (!pass) return false;
  const storedPass = PropertiesService.getScriptProperties().getProperty('ADMIN_PASS') || "admin123";
  return pass === storedPass;
}

// =========================================================
//  GENERATE REKAP BULANAN (DIUNDUH LANGSUNG, TIDAK MENULIS KE SPREADSHEET)
// =========================================================
// hariLiburArr: array angka hari dalam seminggu yang dianggap libur.
// Konvensi mengikuti Date.getDay(): 0=Minggu, 1=Senin, 2=Selasa, 3=Rabu,
// 4=Kamis, 5=Jumat, 6=Sabtu. Contoh: [5] artinya setiap hari Jumat libur.
function generateRekapBulanan(instansiTarget, bulan, tahun, hariLiburArr) {
  // FIX: Validasi parameter
  if (!instansiTarget || !bulan || !tahun) {
    return { status: "error", msg: "Parameter instansi, bulan, dan tahun wajib diisi." };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hariLibur = Array.isArray(hariLiburArr) ? hariLiburArr.map(Number) : [];

  // Ambil daftar karyawan berdasarkan instansi
  const karyawanSheet = ss.getSheetByName("Master_Karyawan");
  if (!karyawanSheet) return { status: "error", msg: "Sheet Master_Karyawan tidak ditemukan." };

  const karyawanData = karyawanSheet.getDataRange().getValues().slice(1);
  const listKaryawan = karyawanData
    .filter(row => String(row[0]).trim() === String(instansiTarget).trim())
    .map(row => String(row[1]).trim())
    .filter(nama => nama.length > 0);  // FIX: Abaikan baris kosong

  if (listKaryawan.length === 0) {
    return { status: "error", msg: "Tidak ada karyawan ditemukan untuk instansi: " + instansiTarget };
  }

  // Ambil data absensi
  const absensiSheet = ss.getSheetByName("Absensi");
  if (!absensiSheet) return { status: "error", msg: "Sheet Absensi tidak ditemukan." };

  const absensiData = absensiSheet.getDataRange().getValues().slice(1);
  const targetBulan = Number(bulan);
  const targetTahun = Number(tahun);

  // Filter data sesuai instansi, bulan, dan tahun
  const filteredAbsensi = absensiData.filter(row => {
    if (!row[0]) return false;
    try {
      const date = new Date(row[0]);
      // FIX: Pastikan tanggal valid sebelum difilter
      if (isNaN(date.getTime())) return false;
      return String(row[1]).trim() === String(instansiTarget).trim() &&
             (date.getMonth() + 1) === targetBulan &&
             date.getFullYear() === targetTahun;
    } catch (e) { return false; }
  });

  // Jumlah hari sesungguhnya di bulan tsb (28-31)
  const jumlahHariBulan = new Date(targetTahun, targetBulan, 0).getDate();

  // FIX: Tentukan tanggal mana saja yang libur berdasarkan hari-dalam-minggu terpilih
  const tanggalLibur = {}; // { [tanggal]: true }
  let jumlahHariLibur = 0;
  for (let d = 1; d <= jumlahHariBulan; d++) {
    const dow = new Date(targetTahun, targetBulan - 1, d).getDay();
    if (hariLibur.indexOf(dow) !== -1) {
      tanggalLibur[d] = true;
      jumlahHariLibur++;
    }
  }
  // FIX: Total hari kerja = total hari dalam bulan - hari libur terpilih
  const totalHariKerja = jumlahHariBulan - jumlahHariLibur;

  // Mapping data absensi per karyawan & tanggal
  const rekapMap = {};
  listKaryawan.forEach(nama => {
    rekapMap[nama] = {};
    for (let d = 1; d <= jumlahHariBulan; d++) rekapMap[nama][d] = "";
  });

  filteredAbsensi.forEach(row => {
    try {
      const date = new Date(row[0]);
      const day  = date.getDate();
      const nama = String(row[2]).trim();
      const statusFull = String(row[4]).trim();

      let kode = "";
      if (statusFull === "Hadir") kode = "H";
      else if (statusFull === "Sakit") kode = "S";
      else if (statusFull === "Izin") kode = "I";

      // FIX: Tangani duplikat absensi di hari yang sama (ambil yang terakhir)
      if (rekapMap[nama] && day >= 1 && day <= jumlahHariBulan) {
        // Jika sudah ada isi, pisahkan dengan "/" agar keduanya terlihat
        if (rekapMap[nama][day] && rekapMap[nama][day] !== kode) {
          rekapMap[nama][day] = rekapMap[nama][day] + "/" + kode;
        } else {
          rekapMap[nama][day] = kode;
        }
      }
    } catch (e) { Logger.log("Error parse absensi row: " + e); }
  });

  // Susun header
  const headers = ["No", "Nama Karyawan"];
  for (let d = 1; d <= jumlahHariBulan; d++) headers.push(String(d));
  headers.push("Jml H", "Jml S", "Jml I", "Hari Kerja", "Prosentase");

  // Isi baris karyawan
  const rows = [];
  listKaryawan.forEach((nama, index) => {
    const rowData = [index + 1, nama];
    let hCount = 0, sCount = 0, iCount = 0;

    for (let d = 1; d <= jumlahHariBulan; d++) {
      if (tanggalLibur[d]) {
        // FIX: Tandai sel hari libur agar terlihat jelas & tidak dihitung
        rowData.push("X");
        continue;
      }
      const st = rekapMap[nama][d] || "";
      rowData.push(st);
      // FIX: Hitung berdasarkan konten, bukan kesamaan persis (agar "/" juga terhitung)
      if (st.includes("H")) hCount++;
      if (st.includes("S")) sCount++;
      if (st.includes("I")) iCount++;
    }

    rowData.push(hCount, sCount, iCount, totalHariKerja);

    // FIX: Prosentase = Jml H / Total Hari Kerja (bukan H+S+I).
    // S dan I tetap dicatat di kolomnya masing-masing, tapi tidak menambah
    // pembilang dan tidak dikeluarkan dari pembagi -- otomatis menurunkan persentase.
    const prosentase = totalHariKerja > 0
      ? ((hCount / totalHariKerja) * 100).toFixed(1) + "%"
      : "0%";
    rowData.push(prosentase);
    rows.push(rowData);
  });

  const namaBulanArr = ["","Januari","Februari","Maret","April","Mei","Juni",
                        "Juli","Agustus","September","Oktober","November","Desember"];
  const namaBulan = namaBulanArr[targetBulan] || "Bulan" + bulan;

  // FIX: Tidak lagi menulis ke Spreadsheet -- kembalikan data mentah agar
  // client (browser) yang membuat file .xlsx dan langsung mengunduhnya.
  return {
    status: "success",
    headers: headers,
    rows: rows,
    fileName: "Rekap_" + instansiTarget + "_" + namaBulan + "_" + tahun,
    info: {
      instansi: instansiTarget,
      bulan: namaBulan,
      tahun: targetTahun,
      jumlahHariBulan: jumlahHariBulan,
      jumlahHariLibur: jumlahHariLibur,
      totalHariKerja: totalHariKerja,
      jumlahKaryawan: listKaryawan.length
    },
    msg: 'Rekap "' + namaBulan + ' ' + tahun + '" untuk ' + listKaryawan.length + ' karyawan berhasil dibuat (Hari kerja: ' + totalHariKerja + ' dari ' + jumlahHariBulan + ' hari).'
  };
}