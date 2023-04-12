// Special thanks: https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70


function sendExportedSheetAsPDFAttachment() {

  /* Update the following variables */
  const ssID = "aaaaaaa"; // sheetID
  const ssGID = "bbbbbbb"; //tabID for the report tab... the first tab has gid=0... the rest have their own next to gid=
  const emailTitle = "Title"; //currently not used
  const dir = DriveApp.getFolderById("ccccccc"); //GDrive folder's unique ID where you want reports to export to
  const driveLinkRowAdj = 3; //which row the "Report Link" title is at in "aggregate" tab
  const driveLinkCol = 16; //which column the "Report Link" title is at
  // you can also change the file name format in the loop below

  const ss2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Report");
  const ss3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Aggregate");
  const lastRow = ss3.getLastRow() - 6;
  const grabName = ss3.getRange('A4:A' + lastRow).getValues();
  const grabEmail = ss3.getRange('B4:B' + lastRow).getValues();
  let estTime = 0;

  for (let i = 0; i < grabName.length; i++){ 
    estTime += i % 20 + 1;
  }
  console.log("Total number of users: ", grabName.length);
  console.log("Estimated time to finish: ", Math.ceil(estTime/60+1)," minutes.");
//loop through the user list
  for (let i = 0; i < grabName.length; i++){ 
    let fullName = grabName[i];
    ss2.getRange('B2').setValue(fullName);
    SpreadsheetApp.flush(); //if you don't flush it, the export will be stuck at the first user
    let refreshName = ss2.getRange('B2').getDisplayValue();
    console.log("Printing for user No.", i+1, ", ", refreshName);
    let printURL = "https://docs.google.com/spreadsheets/d/" + ssID + "/export?exportFormat=pdf&format=pdf&size=0&portrait=true&gridlines=false&gid=" + ssGID;
    let fName = fullName; //Change export file name format here
    let timeOut = (i % 20 + 1) * 1000;
    Utilities.sleep(timeOut); //Exponential timeout; Google reported 429 when exporting too fast
    let blob = getFileAsBlob(printURL, fName); //get PDF file
    let exportedFile = dir.createFile(blob); //put PDF file in the folder dir
    let exportedURL = exportedFile.getUrl();
    ss3.getRange(i+driveLinkRowAdj+1,driveLinkCol).setValue(exportedURL); //update the cell with GDrive PDF file link
    
  
    /* when we want to send email directly
    var message = {
      cc: grabEmail[i].join(","),
      subject: emailTitle,
      body: "Dear " + fullName +", \n\nGreetings. \n\nThank you",
      attachments: [blob]
    };
    MailApp.sendEmail(message);
  */
  }
}

function getFileAsBlob(fileURL, fileName) {
  let token = ScriptApp.getOAuthToken();
  let blobs;
  let response = UrlFetchApp.fetch(fileURL, {
    headers: {
      'Authorization': 'Bearer ' +  token
    }
  });
  blobs = response.getBlob().setName(fileName + '.pdf');
  console.log("Grabbing PDF for", fileName,"...");
  console.log("Storage Space used: " + DriveApp.getStorageUsed()); //Use it to confirm if we got a non-zero-size file
  return blobs;
}