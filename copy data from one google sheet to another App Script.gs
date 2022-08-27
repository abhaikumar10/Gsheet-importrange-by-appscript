/////////////---Main   runner--/////////////
function runn()  {
     const alast=dataexe();
     for(var i=2; i<=alast;i++) {
    Logger.log(" i ="+i+" \n last row = "+alast);
    
    //skip for flag
    var ssd = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ssdraw");
    var qqq=ssd.getRange(i,8).getValue();
       if(qqq===0){
        continue;
          } 

//data collection
    var ssd = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ssdraw");
        var dd={};
          dd.sourceID = ssd.getRange(i,3).getValue();
          dd.sourceRange = ssd.getRange(i,4).getValue();
          dd.destinationID = ssd.getRange(i,5).getValue();
          dd.destinationRangeStart = ssd.getRange(i,6).getValue();
    //Logger.log("i= "+i+" "+dd.sourceID+"-"+dd.sourceRange+"-"+dd.destinationID+"-"+dd.destinationRangeStart);
//sending to write
    
    try {  
           importRangee(dd.sourceID, dd.sourceRange, dd.destinationID, dd.destinationRangeStart);

           var logg= "Copy done: "+[dd.sourceRange]+" - " +[dd.destinationRangeStart];   
            logData(dd,null,logg);
           Logger.log("copy done");
           
         }
      
      catch(errr) 
          {
             dd.sheetname = ssd.getRange(i,2).getValue();
             var ee= "Error is:>> " +[errr]+" <<On Sheet: "+ dd.sheetname+"-"+dd.sourceRange;
             //Logger.log("Error is:>> " +errr+" <<On Sheet: "+ dd.sheetname+"-"+dd.sourceRange);
             logData(dd,ee,null);
            if(errr!=null)
            {
              continue;
            }
         
          }

  }
  

}
 

/////////////---data writer--/////////////
//thanks to :https://stackoverflow.com/questions/71312886/trying-to-copy-template-sheet-into-newly-created-sheet-via-apps-script

function importRangee(sourceID, sourceRange, destinationID, destinationRangeStart){
 
  // Gather the source range values
  const sourceSS = SpreadsheetApp.openById(sourceID);
  const sourceRng = sourceSS.getRange(sourceRange)
  const sourceVals = sourceRng.getValues();
 
  // Get the destiation sheet and cell location.
  const destinationSS = SpreadsheetApp.openById(destinationID);
  const destStartRange = destinationSS.getRange(destinationRangeStart);
  const destSheet = destStartRange.getSheet();
 
  // Clear previous entries.
  destStartRange.clear();
 
  // Get the full data range to paste from start range.
  const destRange = destSheet.getRange(
      destStartRange.getRow(),
      destStartRange.getColumn(),
      sourceVals.length,
      sourceVals[0].length
    );
  
  // Paste in the values.
  destRange.setValues(sourceVals);
 
  SpreadsheetApp.flush();
}



/////////////---last row check--/////////////
//limiting to fetch range of data to copy paste
function dataexe() {
 var ssd = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ssdraw");
    for(var j=1; j<50;j++) 
        {
         
          const lrow = ssd.getLastRow();
          const avals = ssd.getRange("A1:A"+lrow).getValues();
          var alast  = lrow - avals.reverse().findIndex(c=>c[0]!='');
        }
        return(alast);
}

/////////////---Logger--/////////////
//create a sheet named Log.
function logData(dd,ee,logg){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("log");
  ss.appendRow([new Date(),dd.destinationID,dd.sourceID+"-"+dd.destinationRangeStart,ee,logg]);
 // Logger.log(yy);
};


