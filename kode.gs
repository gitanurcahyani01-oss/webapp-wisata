const SPREADSHEET_ID = "1p4_mQYpu0CvJVjpBvytVD-q09KGkkvJhf1EW8PF6ZO0";

const FOLDER_ID = "1NvdZpGUbvstfB39WZZWxgrecpKwbdoUu";



// =========================
// GET DATA
// =========================

function doGet() {

  const sheet = SpreadsheetApp
    .openById(SPREADSHEET_ID)
    .getSheetByName("Sheet1");

  const values = sheet.getDataRange().getValues();

  let result = [];

  for (let i = 1; i < values.length; i++) {

    result.push({
      id: values[i][0],
      nama_wisata: values[i][1],
      kategori: values[i][2],
      lokasi_kota: values[i][3],
      harga_tiket: values[i][4],
      deskripsi: values[i][5],
      gambar: values[i][6]
    });

  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}



// =========================
// POST DATA
// =========================

function doPost(e) {

  try {

    const sheet = SpreadsheetApp
      .openById(SPREADSHEET_ID)
      .getSheetByName("Sheet1");

    const folder = DriveApp.getFolderById(FOLDER_ID);

    const data = JSON.parse(e.postData.contents);



    // VALIDASI DATA
    if (!data.gambar) {

      return ContentService
        .createTextOutput(JSON.stringify({
          status: "error",
          message: "Gambar tidak ada"
        }))
        .setMimeType(ContentService.MimeType.JSON);

    }



    // =========================
    // PROSES BASE64
    // =========================

    const matches = data.gambar.match(/^data:(.+);base64,(.+)$/);

    if (!matches) {

      return ContentService
        .createTextOutput(JSON.stringify({
          status: "error",
          message: "Format gambar salah"
        }))
        .setMimeType(ContentService.MimeType.JSON);

    }

    const contentType = matches[1];
    const base64Data = matches[2];



    // EXTENSION FILE
    let extension = "png";

    if (contentType.includes("jpeg")) {
      extension = "jpg";
    }

    if (contentType.includes("png")) {
      extension = "png";
    }

    if (contentType.includes("webp")) {
      extension = "webp";
    }



    // =========================
    // BUAT FILE
    // =========================

    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      contentType,
      "wisata_" + new Date().getTime() + "." + extension
    );



    // =========================
    // UPLOAD KE DRIVE
    // =========================

    const file = folder.createFile(blob);



    // SHARE PUBLIC
    file.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );



    // URL GAMBAR
    const imageUrl =
      "https://drive.google.com/uc?export=view&id=" + file.getId();



    // =========================
    // SIMPAN KE SHEETS
    // =========================

    sheet.appendRow([
      data.id,
      data.nama_wisata,
      data.kategori,
      data.lokasi_kota,
      data.harga_tiket,
      data.deskripsi,
      imageUrl
    ]);



    return ContentService
      .createTextOutput(JSON.stringify({
        status: "success",
        imageUrl: imageUrl
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {

    return ContentService
      .createTextOutput(JSON.stringify({
        status: "error",
        message: err.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);

  }

}
