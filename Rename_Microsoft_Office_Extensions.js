function RenameMicrosoftDriver() {

    let files = DriveApp.getFiles();
    let anyChangeCnt = 0; 
  
    while (files.hasNext()) {
      let file = files.next();
      let fileName = file.getName();
      
      if (fileName.endsWith(".xlsx")) 
        RenameFile(file, fileName, ".xlsx", ".gsheet");
      else if (fileName.endsWith(".docx")) 
        RenameFile(file, fileName, ".docx", ".gdoc");
      else if (fileName.endsWith(".pptx")) 
        RenameFile(file, fileName, ".pptx", ".gslides");
      else if (fileName.endsWith("ppt")) 
        RenameFile(file, fileName, ".ppt", ".gslides");
    }
  }
  
  
  function RenameFile(file, fileName, begExt, endExt) {
      let begExtLen = begExt.length;
      let fileNameLen = fileName.length;
      let newFileName = fileName.slice(0, fileNameLen - begExtLen);
      newFileName = newFileName + endExt;
      Logger.log("The old file name was: " + fileName);
      file.setName(newFileName);
      fileName = file.getName();
      Logger.log("The new file name is: " + fileName);
  }
  