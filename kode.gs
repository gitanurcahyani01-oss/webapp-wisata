function doGet(e) {

  var sheet = SpreadsheetApp
    .openById("1p4_mQYpu0CvJVjpBvytVD-q09KGkkvJhf1EW8PF6ZO0")
    .getSheetByName("Sheet1");

  var data = sheet.getDataRange().getValues();

  var result = [];

  for (var i = 1; i < data.length; i++) {

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



function doPost(e) {

  var sheet = SpreadsheetApp
    .openById("1p4_mQYpu0CvJVjpBvytVD-q09KGkkvJhf1EW8PF6ZO0")
    .getSheetByName("Sheet1");

  var data = JSON.parse(e.postData.contents);

  sheet.appendRow([
    data.id,
    data.nama_wisata,
    data.kategori,
    data.lokasi_kota,
    data.harga_tiket,
    data.deskripsi,
    data.gambar
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({
      status: "success"
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
