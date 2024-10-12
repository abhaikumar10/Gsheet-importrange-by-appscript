function getdata(sourceID,sourceRange) {
  var gsheet = SpreadsheetApp.openById(sourceID);
  let sourceRng = gsheet.getRange(sourceRange)
  let data  = sourceRng.getValues();
        return(data);
}

