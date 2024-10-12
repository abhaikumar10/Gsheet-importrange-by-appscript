
 //---Logger--///
function logData(dd,ee,logg){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("log");
  ss.appendRow([new Date(),dd.destinationID+"-"+dd.destinationRangeStart,dd.sourceID+"-"+dd.sourceRange,ee,logg]);
  SpreadsheetApp.flush();
 // Logger.log(yy);
};