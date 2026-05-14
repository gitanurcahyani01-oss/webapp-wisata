var SHEET_ID = '1vHwgHIdpE2cGlax9LBrP0EKdwAmlq-f-nwi3A-CEscY'; 
var FOLDER_ID = '14rV8lA0qO1vW7P7LLTOsyCIa54pFrEhz';

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle('Web App Wisata')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Fungsi Simpan Data
function uploadDataWisata(formObject) {
  try {
    var folder = DriveApp.getFolderById(FOLDER_ID);
    var sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();

    var fileData = formObject.fileData;
    var fileName = formObject.fileName;
    var mimeType = formObject.mimeType;

    var blob = Utilities.newBlob(Utilities.base64Decode(fileData), mimeType, fileName);
    var file = folder.createFile(blob);
    var fileUrl = file.getUrl();

    sheet.appendRow([
      formObject.id,
      formObject.nama,
      formObject.kategori,
      formObject.lokasi,
      formObject.harga,
      fileUrl
    ]);

    return "Sukses menyimpan data!";
  } catch (error) {
    return "Error: " + error.toString();
  }
}

// Fungsi Ambil Data untuk ditampilkan di Box 2
function getDataWisata() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
  return sheet.getDataRange().getValues(); 
}
