function importmail2sheet() {
        if(kessa!="Mail") {
                return; 
              }  
        var sheet = gsheet.getSheetByName(fileName);     //Logger.log(sheet+" Loading...");
        if (sheet) {
            var data = sheet.getDataRange().getValues(); 
            data=getfile(dd) 
        }
        //Filter Data for YT or Overall Dates
        if(ssd.getRange(i,6).getValue()=="YT") {
            data = data.filter(function(row) {
                    var dateStr = row[0];
                    var date = new Date(dateStr);
                    return date.toDateString() === yesterday.toDateString();
                  });  Logger.log("fltr done: "+ data);
            
        }  

        ///Filter by date
        for (var j = 0; j< data.length; j++) {
              if (data[j][0] instanceof Date) {
                data[j][0] = Utilities.formatDate(new Date(data[j][0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
              }
        }  Logger.log("copy started");
        importRangee(data,io,destinationID, destinationRangeStart);
     
} 
