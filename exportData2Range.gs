function importRangee(data,how, destinationID, destinationRangeStart){  
  const destinationSS = SpreadsheetApp.openById(destinationID);
  const destStartRange = destinationSS.getRange(destinationRangeStart);
  const destSheet = destStartRange.getSheet();
  SpreadsheetApp.flush();
    if(how=="Replace") {
    destStartRange.clear();  
    const destRange = destSheet.getRange(
                      destStartRange.getRow(),
                      destStartRange.getColumn(),
                      data.length,
                      data[0].length
                    );
    destRange.setValues(data); 
    }
  if(how=="Append") { 
    destSheet.getRange(destSheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
    //destSheet.appendRow(data);
     
  }
  else {
    return "how is how incorrect. Use (how=Replace or Append)";
  }
  SpreadsheetApp.flush();
}
