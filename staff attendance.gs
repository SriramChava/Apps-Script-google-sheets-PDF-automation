var ID = "https://docs.google.com/spreadsheets/d/1MLJeGUdLFPgnDuirYW8Uhkd-fAyReqlA5dgeURgqK2c/edit#gid=0";  
var lock = 'admin'                                         

/*======================================================== */

/* ==================== DO NOT CHANGE ANYTHING BELOW THIS LINE  ======================== */

var conf = 'config'
var ss = SpreadsheetApp.openByUrl(ID)

function doGet(e) {
  if (Object.keys(e.parameter).length === 0) {
    var htmlFile
    var sheetName = conf
    var activeSheet = ss.getSheetByName(sheetName)
    if (activeSheet !== null) {
      var values = activeSheet.getDataRange().getValues();
      for(var i=0, iLen=values.length; i<iLen; i++) {
        if(values[i][0] == 'Passcode') {
          var passCheck = activeSheet.getRange(i+1, 2).getValues()
          if(passCheck == lock) {
            htmlFile = 'Main'
            activeSheet.getRange(i+1, 2).clearContent()
          } else {
            htmlFile = 'Login'
          }
        }
      }
    } else {
      config()
      htmlFile = 'Login'
    }
   // return HtmlService.createHtmlOutputFromFile(htmlFile);
   return HtmlService.createHtmlOutputFromFile(htmlFile).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function removeEmptyColumns(sheetName) {
  var activeSheet = ss.getSheetByName(sheetName)
  var maxColumns = activeSheet.getMaxColumns(); 
  var lastColumn = activeSheet.getLastColumn();
  if (maxColumns-lastColumn != 0){
    activeSheet.deleteColumns(lastColumn+1, maxColumns-lastColumn);
  }
}

function validateUser(passcode) {
  if (passcode == lock) {
    var successMessage = 'Logging you in!';
    config(passcode)
    return successMessage
  } else {
    var errorMessage = 'Incorrect passcode :(';
    return errorMessage
  }
}

function config(passcode) {
  var sheetName = conf
  var activeSheet = ss.getSheetByName(sheetName)
  if (activeSheet == null) {
    activeSheet = ss.insertSheet().setName(sheetName);
    activeSheet.appendRow (["Config"])
    activeSheet.appendRow (["Lock"])
    activeSheet.appendRow (["Passcode"])
    removeEmptyColumns(sheetName);
    activeSheet.setFrozenRows(1)
    if (passcode !== undefined) {
      var values = activeSheet.getDataRange().getValues();
      var sheetRow;
      for(var i=0, iLen=values.length; i<iLen; i++) {
        if(values[i][0] == 'Passcode') {
          sheetRow = i+1
          activeSheet.getRange(sheetRow, 2).setValue(passcode)
        }
      }
    }
  } else {
    var values = activeSheet.getDataRange().getValues();
    var sheetRow;
    for(var i=0, iLen=values.length; i<iLen; i++) {
      if(values[i][0] == 'Passcode') {
        sheetRow = i+1
        activeSheet.getRange(sheetRow, 2).setValue(passcode)
      }
    }
  }
}

function webAppURL(linkAddr) {
  var linkAddr = ScriptApp.getService().getUrl()
  return linkAddr
}



// website related functions

var url = "https://docs.google.com/spreadsheets/d/1detjZoWB1zUu4e9SPDpMemK_nM0fKQE7285PhFb1Pwc/edit#gid=0";
var ss1 = SpreadsheetApp.openByUrl(url);
var ws = ss1.getSheetByName("data");

function entryNameAndTime(name)
{
  var dateAndTimeEntry = new Date();
  var currentTimeEntry = dateAndTimeEntry.toLocaleTimeString();
  var currentDateEntry = dateAndTimeEntry.toDateString();
  var endrow = ws.getLastRow();
  var exited = true;
  var data = ws.getDataRange().getValues();
   while(endrow-1)
    {   
      if(data[endrow-1][0]==name && data[endrow-1][3] == 0)
        exited = false;
        endrow--;
    }
    if(exited && name != "")
       ws.appendRow([name,currentDateEntry,currentTimeEntry,0]);
}
function addOffForEntry(name)
{
  if(name!=""){
   var dateAndTimeEntry = new Date();
   var currentDateEntry = dateAndTimeEntry.toDateString();
   ws.appendRow([name,currentDateEntry,"Off","Off"]);
  }
}
function addLeaveForEntry(name)
{
  if(name!="")
  {
    var dateAndTimeEntry = new Date();
    var currentDateEntry = dateAndTimeEntry.toDateString();
    ws.appendRow([name,currentDateEntry,"Leave","Leave"]);
  }
}
function exitNameAndTime(name)
{
  var dateAndTime = new Date();
  var currentTime = dateAndTime.toLocaleTimeString();
  data = ws.getDataRange().getValues();
    endrow = ws.getLastRow();
    while(endrow-1)
    {   if(data[endrow-1][0]==name && data[endrow-1][3] == 0)
        ws.getRange(endrow,4).setValue(currentTime);
        endrow--;
    }
}

function deleteALL ()
{ 
  endrow = ws.getLastRow();
  while(endrow-1)
    {
      ws.deleteRow(endrow);
      endrow--;
    }
}

function saveDataAndGeneratePDF()
 {  
    var date  = new Date();
    endrow = ws.getLastRow();
    data = ws.getDataRange().getValues();
    var classtaken = {};
    var offtaken = {};
    var leavetaken = {};
    for (var i = 1; i < endrow ; i++) 
      {        
        if (offtaken[data[i][0].toString().toUpperCase()] === undefined && data[i][2]=="Off") 
            offtaken[data[i][0].toString().toUpperCase()] = 1;
        else if(data[i][2]=="Off")
            offtaken[data[i][0].toString().toUpperCase()]++;
      }
    for (var i = 1; i < endrow ; i++) 
      {        
        if (leavetaken[data[i][0].toString().toUpperCase()] === undefined && data[i][2]=="Leave") 
            leavetaken[data[i][0].toString().toUpperCase()] = 1;
        else if(data[i][2]=="Leave")
            leavetaken[data[i][0].toString().toUpperCase()]++;
      }
    for (var i = 1; i < endrow ; i++) 
      {        
        if (classtaken[data[i][0].toString().toUpperCase()] === undefined && (data[i][2]!="Leave" && data[i][2]!="Off")) 
            classtaken[data[i][0].toString().toUpperCase()] = 1;
        else if(data[i][2]!="Off" && data[i][2]!="Leave")
            classtaken[data[i][0].toString().toUpperCase()]++;
      }

    
      saveDataInDrive();
      deleteALL();
      generateAndMailPDFs(classtaken,offtaken,leavetaken);

  }

function saveDataInDrive()
{
        // staff attendance sheet link : https://docs.google.com/spreadsheets/d/1detjZoWB1zUu4e9SPDpMemK_nM0fKQE7285PhFb1Pwc/edit?usp=sharing
      var staffAttendanceSheet = DriveApp.getFileById("1detjZoWB1zUu4e9SPDpMemK_nM0fKQE7285PhFb1Pwc"); 
      // Staff attendance data folder link : https://drive.google.com/drive/folders/1blEzTHOk8QdIbgQfJTtq0FTqxi_GSsdO?usp=sharing
      var staffAttendanceSheetFolder = DriveApp.getFolderById("1blEzTHOk8QdIbgQfJTtq0FTqxi_GSsdO");
      var date = new Date();
      var month = Utilities.formatDate(date, Session.getScriptTimeZone(), "MMM");
      var year = date.getFullYear();
      var filename = "Staff Attendance Report " + month + " " + year; 
      staffAttendanceSheet.makeCopy(filename,staffAttendanceSheetFolder);
}

function generateAndMailPDFs(classtaken,offtaken,leavetaken)
{}
