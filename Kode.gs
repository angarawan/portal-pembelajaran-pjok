// Ganti dengan ID Spreadsheet Anda
const SPREADSHEET_ID = '1K86t0nXVj6gmbacQl2vOfvg2APbnj_Amn_4BwH7eyFU';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('LMS PJOK')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Fungsi untuk mengambil semua data awal saat web dimuat
function getInitialData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 1. Ambil Judul LMS & Background
  let judul = "LMS PJOK";
  let background = "";
  try {
    const sheetPengaturan = ss.getSheetByName('Pengaturan');
    if(sheetPengaturan) {
      judul = sheetPengaturan.getRange("B1").getValue() || judul;
      background = sheetPengaturan.getRange("B2").getValue() || ""; // Mengambil background dari B2
    }
  } catch(e) {}

  // 2. Ambil Data Siswa (Kelas, Nama, Absen)
  let dataSiswa = [];
  try {
    const sheetSiswa = ss.getSheetByName('DataSiswa');
    if(sheetSiswa) {
      const vals = sheetSiswa.getDataRange().getDisplayValues();
      // Lewati baris pertama jika itu header
      for(let i = 1; i < vals.length; i++) {
        if(vals[i][0] && vals[i][1]) {
          dataSiswa.push({
            kelas: vals[i][0].toString().trim(),
            nama: vals[i][1].toString().trim(),
            absen: vals[i][2].toString().trim()
          });
        }
      }
    }
  } catch(e) {}
  // 3. Ambil Materi (6 Kolom: Nama Utama, Sub Materi, Isi, Gambar, Video, Lampiran)
  let dataMateri = [];
  try {
    const sheetMateri = ss.getSheetByName('Materi');
    if(sheetMateri) {
      const vals = sheetMateri.getDataRange().getDisplayValues();
      for(let i = 1; i < vals.length; i++) {
        if(vals[i][0] || vals[i][1] || vals[i][2]) {
          dataMateri.push({
            id: i,
            namaMateri: vals[i][0] ? vals[i][0].toString().trim() : "Materi Umum", // Kolom A: Nama Materi Utama
            subMateri: vals[i][1] ? vals[i][1].toString().trim() : "",            // Kolom B: Sub Materi
            isiMateri: vals[i][2] ? vals[i][2].toString().trim() : "",            // Kolom C: Isi Teks
            gambarMateri: vals[i][3] ? vals[i][3].toString().trim() : "",         // Kolom D: URL Gambar
            videoMateri: vals[i][4] ? vals[i][4].toString().trim() : "",          // Kolom E: URL YouTube
            lampiranMateri: vals[i][5] ? vals[i][5].toString().trim() : ""       // Kolom F: URL Lampiran
          });
        }
      }
    }
  } catch(e) {}

  return {
    judul: judul,
    background: background,
    siswa: dataSiswa,
    materi: dataMateri
  };
}

// Fungsi untuk menyimpan rekam jejak siswa setelah selesai
function saveRecord(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheetRekam = ss.getSheetByName('RekamJejak');
  
  // Jika sheet belum ada, buat baru dan beri header
  if (!sheetRekam) {
    sheetRekam = ss.insertSheet('RekamJejak');
    sheetRekam.appendRow(['Waktu Mulai', 'Waktu Selesai', 'Kelas', 'Nama', 'No Absen', 'Materi Dipilih']);
    sheetRekam.getRange("A1:F1").setFontWeight("bold");
  } else {
    // Pastikan header kolom F (Materi Dipilih) ada
    const headers = sheetRekam.getRange("A1:F1").getValues()[0];
    if (!headers[5]) {
      sheetRekam.getRange("F1").setValue("Materi Dipilih").setFontWeight("bold");
    }
  }

  // Format tanggal untuk kerapian
  const mulai = new Date(data.startTime);
  const selesai = new Date(data.endTime);
  const formattedMulai = Utilities.formatDate(mulai, "GMT+8", "dd/MM/yyyy HH:mm:ss");
  const formattedSelesai = Utilities.formatDate(selesai, "GMT+8", "dd/MM/yyyy HH:mm:ss");

  sheetRekam.appendRow([
    formattedMulai,
    formattedSelesai,
    data.kelas,
    data.nama,
    data.absen,
    data.materiDipilih || "Semua Materi"
  ]);

  return "Berhasil disimpan";
}