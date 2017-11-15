function convertExceltoGoogleSpreadsheet(fileName) {
  
  try {
    
    // Written by Sarmad Gardezi
    // sarmadgardezi.com

    fileName = fileName || "microsoft-excel.xlsx";
    
    var excelFile = DriveApp.getFilesByName(fileName).next();
    var fileId = excelFile.getId();
    var folderId = Drive.Files.get(fileId).parents[0].id;  
    var blob = excelFile.getBlob();
    var resource = {
      title: excelFile.getName(),
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [{id: folderId}],
    };
    
    Drive.Files.insert(resource, blob);
    
  } catch (f) {
    Logger.log(f.toString());
  }
  
}