# Convert Microsoft Excel to Google Spreadsheet
How to Convert Microsoft Excel to Google Spreadsheet Format with Apps Script

If your colleagues have been emailing you Microsoft Excel spreadsheets in xls or xlsx format, here’s a little snippet that will help you convert those Excel sheets into native Google Spreadsheet format using the Advanced Drive API service of Google Apps Script.

```ruby
function convertExceltoGoogleSpreadsheet(fileName) {
  
  try {
    
    // Written by Sarmad GArdezi

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
```

The script finds the existing Excel workbook by name in your Google Drive, gets the blob of the file and creates a new file of Google Sheets mimetype (application/vnd.google-apps.spreadsheet) with the blob.

You do need to enable the Google Drive API under Resources > Advanced Google Services and also enable the Drive API inside the Google Cloud Platform project associated with your Google Apps Script.

The other option, instead of specifying the mimetype, is to set the argument convert to true and it will automatically convert the source file into corresponding native Google Format at the time of insert it into Google drive.
```ruby
unction convertExceltoGoogleSpreadsheet2(fileName) {
  
  try {
    
    fileName = fileName || "microsoft-excel.xlsx";
    
    var excelFile = DriveApp.getFilesByName(fileName).next();
    var fileId = excelFile.getId();
    var folderId = Drive.Files.get(fileId).parents[0].id;  
    var blob = excelFile.getBlob();
    var resource = {
      title: excelFile.getName().replace(/.xlsx?/, ""),
      key: fileId
    };
    Drive.Files.insert(resource, blob, {
      convert: true
    });
    
  } catch (f) {
    Logger.log(f.toString());
  }
  
}
```

Written by [Sarmad Gardezi](https://sarmadgardezi.com) for [Developers](http://thedevelopers.pk) & Beginners.
