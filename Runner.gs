function newrunner() {
const dash= SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard')
const ssd= SpreadsheetApp.getActiveSpreadsheet().getSheetByName('my_imports');
   const ln=5 || ssd.getRange('A1:A').getLastRow();
   for(i=2;i<=ln;i++)
   {                                                                Logger.log("for row i= "+i)
      let filterKey= ssd.getRange(i,5).getValue()|| '';
      let copytype=ssd.getRange(i,1).getValue() || "Gsheet";
      let how = ssd.getRange(i,2).getValue() || "Append";
      let fname = ssd.getRange(i,3).getValue();
      let filterBy= ssd.getRange(i,4).getValue();
      let ytorall=ssd.getRange(i,6).getValue();
      let sourceID = ssd.getRange(i,3).getValue();  
      let sourceRange = ssd.getRange(i,9).getValue();
      let filterColumnNum= null || 1;
      let destinationID = ssd.getRange(i,7).getValue();
      let destinationRange = ssd.getRange(i,8).getValue();
                                                                  //  Logger.log(filterColumnNum,filterBy,filterKey);
  switch(copytype) {
    case "Gsheet ID":
                importSheet2Sheet(sourceID,sourceRange, destinationID, destinationRange,how ,filterColumnNum,filterBy,filterKey)
                break;
    case "Gsheet Name":
                importMail2Sheet(sourceID,sourceRange, destinationID, destinationRange,how ,filterColumnNum,filterBy,filterKey)
                break;
    case "mailexcel":
                importemailExcel2Sheet(sourceID,sourceRange, destinationID, destinationRange,how ,filterColumnNum,filterBy,filterKey)
                break;
    case "mailcsv":
                importMail2Sheet(sourceID,sourceRange, destinationID, destinationRange,how ,filterColumnNum,filterBy,filterKey)
                break;
    default: Logger.log(`${copytype} not a valid Copy type choose from Gsheet, Mail etc`);
    Logger.log(`${sourceID} on ${sourceRange} Copy Done`);
  }
}
}
