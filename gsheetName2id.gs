//input takes file name and return sheet id.
 
function filename2id(fileName) {  
    var files = DriveApp.getFilesByName(fileName); 
  if (files.hasNext()) {
    var file = files.next();
    var fid=file.getId();
    } 
    return fid;
}  
