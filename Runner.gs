function sheetrun()  {
     const alast=mygetlast();
     for(var i=2; i<=alast;i++) {
try{   Logger.log(" i ="+i+" \n last row = "+alast);
     SpreadsheetApp.flush();
    var ssd = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("my_imports");
    var qqq=ssd.getRange(i,10).getValue();
       if(qqq===0){
        continue;
          } 
    var ssd = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("my_imports");
     var yesterday = new Date();
      yesterday.setDate(yesterday.getDate() - 1);
        var dd={};
         dd.fname = ssd.getRange(i,3).getValue();
         dd.fmate= ssd.getRange(i,4).getValue();
          dd.kessa=ssd.getRange(i,1).getValue()
          dd.io = ssd.getRange(i,2).getValue();
          dd.xlsxsheet = ssd.getRange(i,5).getValue();
         if(dd.kessa=="Mail") {
          dd.data=getfile(dd)
           if(dd.data.length==0){
              var ee="no core data";
                logData(dd,ee,null);
                continue;
            }
         }
          if(dd.kessa=="Gsheet")  {
          dd.sourceID = dd.fname
          dd.sourceRange = ssd.getRange(i,9).getValue();
           dd.sourceSS = SpreadsheetApp.openById(dd.sourceID);
            dd.sourceRng = dd.sourceSS.getRange(dd.sourceRange)
            dd.data  = dd.sourceRng.getValues();
            if(dd.data.length==0){
              var ee="no core data";
                logData(dd,ee,null);
                continue;
            }
          }
        if(dd.kessa=="mailexcel")  {
         //var dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), dd.fmate);
         var yesterday = new Date();
            yesterday.setDate(yesterday.getDate() -1 ); 
            z= yesterday.toDateString()
            y=Utilities.formatDate(yesterday, Session.getScriptTimeZone(), dd.fmate);       
           const emailSubject = dd.fname+y;
          dd.fid=e2s(emailSubject) 
 var todayDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), dd.fmate);
     Logger.log(dd)
     var yesterday = new Date();
      yesterday.setDate(yesterday.getDate() - 1); 
      z= yesterday.toDateString()
      y=Utilities.formatDate(yesterday, Session.getScriptTimeZone(), dd.fmate)
    if(dd.kessa=="mailexcel"){
      var post=z
       var fid=dd.fid;
    } else{
       var post=todayDate+".csv"
    }
     fileName=dd.fname+post;
           dd.data=getfile(dd) 
           if(dd.data.length==0){
              var ee="no core data";
                logData(dd,ee,null);
                continue;
              }
          }
Logger.log("mil gya data "+dd.data.length)
           dd.destinationID = ssd.getRange(i,7).getValue();
          dd.destinationRangeStart = ssd.getRange(i,8).getValue();
    try {     if(ssd.getRange(i,6).getValue()=="YT") {
        dd.data = dd.data.filter(function(row) {
                    var dateStr = row[0];
                    var date = new Date(dateStr);
                    return date.toDateString() === yesterday.toDateString();
                  });  Logger.log("fltr done: "+ dd.data);
                  if(dd.data.length==0){
                    var ee="no YT data";
                     logData(dd,ee,null);
                     continue;
                  }
            } 
             for (var j = 0; j< dd.data.length; j++) {
              if (dd.data[j][0] instanceof Date) {
                dd.data[j][0] = Utilities.formatDate(new Date(dd.data[j][0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
              }
            }  Logger.log("copy started");
           importRangee(dd.data,dd.io,dd.destinationID, dd.destinationRangeStart);
           var logg= "Copy done: "+[dd.fname] ; 
            logData(dd,null,logg);
           Logger.log("copy done"); 
                  }
    catch(errr)
          {
             dd.sheetname = ssd.getRange(i,2).getValue();
             var ee= "Error is:>> " +[errr]+" <<On Sheet copy: "+ dd.sheetname+"-"+dd.sourceRange||dd.destinationRangeStart;
             logData(dd,ee,null);
             Logger.log(errr);
            if(errr!=null)
            {
              continue;
            }
          }
}
catch(errr)
          {
             dd.sheetname = ssd.getRange(i,2).getValue();
             var ee= "Error is:>> " +[errr]+" <<On Sheet: "+ dd.sheetname+"-"+dd.sourceRange||dd.destinationRangeStart;
             logData(dd,ee,null);
            if(errr!=null)
            {
              continue;
            }
          }
  }
}
