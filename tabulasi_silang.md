# Proses Pembuatan tabulasi Silang 

## 1. Bagian I

Langkah pertama yang saya lakukan adalah mengimpor data ke dalam Spreadsheet. Setelah data berhasil diimpor, saya beralih ke bagian Ekstensi, kemudian navigasi ke Apps Script. Di sana, saya memberikan nama file yang diperlukan, lalu memasukkan kode berikut : 

function causeAndYear() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Penyebab Kematian di Indonesia yang Dilaporkan - Clean');
  let numberRows = sheet.getLastRow();
  let numberCols = sheet.getLastColumn();

  let data = sheet.getRange(2, 1, numberRows - 1, numberCols).getValues();

  // Mendapatkan penyebab dan tahun yang unik
  let causes = [];
  let years = [];
  for (let i = 0; i < data.length; i++) {
    let cause = data[i][0];
    let year = data[i][2];
    if (causes.indexOf(cause) == -1) causes.push(cause);
    if (years.indexOf(year) == -1) years.push(year);
  }

  // Mengurutkan tahun secara menaik
  years.sort((a, b) => a - b);

  // Membuat lembar kerja baru
  let newSheet = ss.insertSheet('Penyebab Kematian di Indonesia');
  newSheet.getRange('A1').setValue('Penyebab');
  for (let i = 0; i < years.length; i++) {
    newSheet.getRange(1, i + 2).setValue(years[i]);
  }

  // Menghitung total dan menambahkannya ke lembar kerja
  for (let i = 0; i < causes.length; i++) {
    let cause = causes[i];
    newSheet.getRange(i + 2, 1).setValue(cause);
    for (let j = 0; j < years.length; j++) {
      let year = years[j];
      let totalDeaths = 0;
      for (let k = 0; k < data.length; k++) {
        if (data[k][0] == cause && data[k][2] == year) {
          totalDeaths += Number(data[k][4]);
        }
      }
      newSheet.getRange(i + 2, j + 2).setValue(totalDeaths);
    }
  }
}

Dalam langkah ini, proses yang dilakukan adalah sebagai berikut:

1. Mendapatkan akses ke lembar kerja (spreadsheet) aktif dengan menggunakan SpreadsheetApp.getActiveSpreadsheet().
2. Mengambil lembar kerja dengan nama 'Penyebab Kematian di Indonesia yang Dilaporkan - Clean' menggunakan ss.getSheetByName().
3. Mendapatkan jumlah baris dan kolom dari lembar kerja tersebut dengan getLastRow() dan getLastColumn().
4. Mengambil data dari lembar kerja, mengabaikan baris header, dan menyimpannya dalam bentuk array dua dimensi data menggunakan getRange().getValues(). Data ini berisi informasi tentang penyebab kematian, tahun, dan lain-lain.
5. Mendapatkan daftar penyebab unik dan daftar tahun unik dari data, dan menyimpannya dalam array causes dan years.
6. Mengurutkan tahun-tahun dalam array years secara berurutan dari kecil ke besar.
7. Membuat lembar kerja baru dengan nama 'Penyebab Kematian di Indonesia' menggunakan ss.insertSheet(). Lembar kerja baru ini akan digunakan untuk menampilkan data yang telah diproses.
8. Menempatkan label "Cause" pada sel 'A1' di lembar kerja baru menggunakan newSheet.getRange('A1').setValue('Cause').
9. Memasukkan tahun-tahun unik ke dalam baris pertama lembar kerja baru menggunakan loop for.
10. Menghitung jumlah total kematian untuk setiap penyebab pada setiap tahun dan memasukkan hasilnya ke dalam sel-sel yang sesuai di lembar kerja baru.

# 1. Bagian II

Langkah selanjutnya adalah membuat tabulasi silang untuk tipe kematian dan tahun. Berikut kode yang digunakan:

function typeAndYear() {
  let ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName('Penyebab Kematian di Indonesia yang Dilaporkan - Clean');
  let numberRows = sheet.getDataRange().getNumRows();
  let numberCols = sheet.getLastColumn();

  let data = sheet.getRange(2, 1, numberRows - 1, numberCols).getValues();

  // Mendapatkan tahun dan tipe yang unik
  let years = [];
  let types = [];
  for (let i = 0; i < data.length; i++) {
    let year = data[i][2];
    let type = data[i][1];
    if (years.indexOf(year) == -1) years.push(year);
    if (types.indexOf(type) == -1) types.push(type);
  }

  // Mengurutkan tahun secara menaik
  years.sort((a, b) => a - b);

  // Membuat lembar kerja baru
  let newSheet = ss.insertSheet('Kematian Berdasarkan Tahun dan Tipe');
  newSheet.getRange('A1').setValue('Tipe');
  for (let i = 0; i < years.length; i++) {
    newSheet.getRange(1, i + 2).setValue(years[i]);
  }

  // Menghitung total dan menambahkannya ke lembar kerja
  for (let i = 0; i < types.length; i++) {
    let type = types[i];
    newSheet.getRange(i + 2, 1).setValue(type);
    for (let j = 0; j < years.length; j++) {
      let year = years[j];
      let totalDeaths = 0;
      for (let k = 0; k < data.length; k++) {
        if (data[k][2] == year && data[k][1] == type) {
          totalDeaths += Number(data[k][4]);
        }
      }
      newSheet.getRange(i + 2, j + 2).setValue(totalDeaths);
    }
  }
}

Dalam langkah ini, proses yang dilakukan serupa dengan bagian a, namun kali ini data diolah berdasarkan tipe kematian dan tahun.