function myFunction() {
            if(dd.kessa=="mailexcel")  {
         //var dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), dd.fmate);
         var yesterday = new Date();
            yesterday.setDate(yesterday.getDate() -1 );  
            z= yesterday.toDateString()
            y=Utilities.formatDate(yesterday, Session.getScriptTimeZone(), dd.fmate);        
           const emailSubject = dd.fname+y;
             dd.fid=e2s(emailSubject)  

    //            if(dd.kessa=="mailexcel"){

    //   var fileName=dd.xlsxsheet
    // }
           dd.data=getfile(dd)  
          //  if(dd.data.length==0){
          //     var ee="no core data";
          //       logData(dd,ee,null);
          //       continue;
          //     }
           
          }
}
