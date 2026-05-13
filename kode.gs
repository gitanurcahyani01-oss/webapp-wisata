const SPREADSHEET_ID = "1p4_mQYpu0CvJVjpBvytVD-q09KGkkvJhf1EW8PF6ZO0";

const FOLDER_ID = "1NvdZpGUbvstfB39WZZWxgrecpKwbdoUu";


// =======================
// GET DATA
// =======================

function doGet() {

  const sheet = SpreadsheetApp
    .openById(SPREADSHEET_ID)
    .getSheetByName("Sheet1");

  const data = sheet.getDataRange().getValues();

  let result = [];

  for(let i = 1; i < data.length; i++){

    result.push({
      id: data[i][0],
      nama_wisata: data[i][1],
      kategori: data[i][2],
      lokasi_kota: data[i][3],
      harga_tiket: data[i][4],
      deskripsi: data[i][5],
      gambar: data[i][6]
    });

  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}



// =======================
// POST DATA
// =======================

function doPost(e) {

  try {

    const sheet = SpreadsheetApp
      .openById(SPREADSHEET_ID)
      .getSheetByName("Sheet1");

    const folder = DriveApp.getFolderById(FOLDER_ID);

    const data = JSON.parse(e.postData.contents);

    // VALIDASI GAMBAR
    if (!data.gambar || data.gambar === "") {

      return ContentService
        .createTextOutput(
          JSON.stringify({
            status: "error",
            message: "Gambar kosong"
          })
        )
        .setMimeType(ContentService.MimeType.JSON);

    }

    // AMBIL BASE64
    const base64Data = data.gambar.split(",")[1];

    // AMBIL CONTENT TYPE
    const contentType =
      data.gambar.match(/^data:(image\/\w+);base64,/)[1];

    // EXTENSION
    const extension = contentType.split("/")[1];

    // BUAT BLOB
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      contentType,
      "wisata_" + new Date().getTime() + "." + extension
    );

    // SIMPAN FILE KE DRIVE
    const file = folder.createFile(blob);

    // SHARE PUBLIC
    file.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );

    // LINK GAMBAR
    const imageUrl =
      "https://drive.google.com/uc?export=view&id=" + file.getId();

    // SIMPAN KE SHEETS
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
      .createTextOutput(
        JSON.stringify({
          status: "success"
        })
      )
      .setMimeType(ContentService.MimeType.JSON);

  } catch(error) {

    return ContentService
      .createTextOutput(
        JSON.stringify({
          status: "error",
          message: error.toString()
        })
      )
      .setMimeType(ContentService.MimeType.JSON);

  }

}
