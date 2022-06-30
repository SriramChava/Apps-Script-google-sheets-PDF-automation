//**************************** DO NOT MODIFY CELL LOCATIONS IN EXCEL AS THE CODE WILL STOP WORKING CORRECTLY***********************/

How to copy a link for google drive folder/doc/sheets etc:  https://docs.google.com/spreadsheets/d/............../edit#gid=1111111111 
Copy the dotted portion of the link and this will be used in .getFile/FolderByID("") function.

// 1. DEFINING GLOBAL VARIABLES

// 1.1 Get google doc file link to edit
   var doc_file = DriveApp.getFileById("LINK FOR GOOGLE DOCS FILE"); 
// 1.2 Temporary folder which holds the above doc for each student this doc is copied, edited and deleted for each student. 
       This is to make sure that the original doc is not changed as then each report will be modified with the earlier changes. 
       After the doc is converted to pdf and mailed the temporary doc is generated again for each report: 
   var Temp_Folder =  DriveApp.getFolderById("LINK FOR TEMPORARY DRIVE FOLDER");
// 1.3 This folder is responsible for holding the reports of all the students. 
   var PDF_Folder =DriveApp.getFolderById("LINK FOR DRIVE FOLER TO HOLD FINAL REPORTS");
// 1.4 google sheet holding attendance,and assignment submission of all the students we get the sheet by name. 
   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NAME OF THE SHEET");
//1.4 Other global variables
  var lastr = sheet.getLastRow()-2;                     // Last occupied row in the "Attendance" excel sheet
  var lastc = sheet.getLastColumn();                    // Last occupied column in the "Attendance" excel sheet
  var tot_p = sheet.getRange("CE6").getValue();         // Total physics classes taken by the faculty in a month
  var tot_c = sheet.getRange("CF6").getValue();         // Total chemistry classes taken by the faculty in a month
  var tot_m = sheet.getRange("CG6").getValue();         // Total mathematics classes taken by the faculty in a month
  var tot_pasn = sheet.getRange("CB6").getValue();      // Total physics assignments given by the faculty in a month
  var tot_casn = sheet.getRange("CC6").getValue();      // Total chemistry assignments given by the faculty in a month
  var tot_masn = sheet.getRange("CD6").getValue();      // Total mathematics assignments given by the faculty in a month


// 2. FUNCTIONS INVOLVED

//2.1 This function takes the data range in the excel as its input and the data for attendance, and assignments are segregated
function getDatarange(firstr,firstc,lastr,lastc) 
{
  return sheet.getRange(firstr,firstc,lastr,lastc).getValues();
}

//2.2 This function filters each individual datarange as data is unclean with empty and null spaces
function filterData (data)
{
  for(var i = 0; i<lastr;i++)
   {
      if(i%4!=0)
        delete data[i];
   }
   return data.filter(element => element !== null);
}

//2.2 After the data ranges for attendance, late attendance, and assignments for PCM are filtered they are concatenated into a single 2D array "st_data"
function concat(parent_data,data1,data2,data3)
{
  for(var i =0; i<parent_data.length;i++)
  {
     parent_data[i] = parent_data[i].concat(data1[i],data2[i],data3[i]);
  }
  return parent_data
}

//2.3 Division of data for replacement of variables in docs
function div_data(st_data)
{
  st_data.forEach(function(row)
  {
      var roll = row[0];
      var name = row[1];
      var mail = row[2]; 
      var att_p = row[3];
      var att_c = row[4];
      var att_m = row[5];
      var late_p = row[6];
      var late_c = row[7];
      var late_m = row[8];
      var pasn = row[9];
      var pasi = row[10];
      var casn = row[11];
      var casi = row[12]; 
      var masn = row[13];
      var masi = row[14];
      
  
create_PDFnMail(tot_p,tot_c,tot_m,name,tot_pasn,tot_casn,tot_masn,roll,mail,att_p,att_c,att_m,late_p,late_c,late_m,pasn,pasi,casn,casi,masn,masi,doc_file,Temp_Folder,PDF_Folder)
  }
 )
}

// 3. AUTOMATED PDF CREATION AND EMAILING

// Temporary file is created and a copy of the original doc having the attendance report template is created in that copy all the variables are replaced a pdf is created and the copied doc is deleted.The pdf is mailed and after this a new copy is created every iteration for each student.
function create_PDFnMail(tot_p,tot_c,tot_m,name,tot_pasn,tot_casn,tot_masn,roll, mail,att_p,att_c,att_m,late_p,late_c,late_m,pasn,pasi,casn,casi,masn,masi,doc_file,Temp_Folder,PDF_Folder)
{
  var temp_File = doc_file.makeCopy(Temp_Folder);
  var temp_doc_file = DocumentApp.openById(temp_File.getId());
  var body = temp_doc_file.getBody();
  body.replaceText("{st_name}",name);
  body.replaceText("{rl_num}",roll);
  body.replaceText("{tot_p}",tot_p);
  body.replaceText("{att_p}",att_p);
  body.replaceText("{tot_c}",tot_c);
  body.replaceText("{att_c}",att_c);
  body.replaceText("{tot_m}",tot_m);
  body.replaceText("{att_m}",att_m);
  body.replaceText("{late_p}",late_p);
  body.replaceText("{late_c}",late_c);
  body.replaceText("{late_m}",late_m);
  body.replaceText("{tot_pasn}",tot_pasn);
  body.replaceText("{tot_casn}",tot_casn);
  body.replaceText("{tot_masn}",tot_masn);
  body.replaceText("{pasn}",pasn);
  body.replaceText("{pasi}",pasi);
  body.replaceText("{casn}",casn);
  body.replaceText("{casi}",casi);  
  body.replaceText("{masn}",masn);
  body.replaceText("{masi}",masi);

  temp_doc_file.saveAndClose();

  var PDF_content = temp_File.getAs(MimeType.PDF);
  var PDF_File = PDF_Folder.createFile(PDF_content).setName("Monthly Attendance" + name);
  Temp_Folder.removeFile(temp_File);

/* FUNCTION FOR MAILING UNCOMMENT WHENEVER WEBAPP IS READY*/
//   // MailApp.sendEmail(mail, 'Monthly Attendance Report', "The monthly attendance report for student " + name + " can be found as follows:",   //{               
//   //   attachments: [PDF_File.getAs(MimeType.PDF)],
//   //   name : 'Progressive minds automated emailer'
//   // });

  Logger.log(PDF_File)
 }

//Parent function this is the main function and runs all the other fuctions for data filtering, automated pdf generation and automated emailing.
function getData()
{
  var st_data = filterData(getDatarange(4,1,lastr,3));
  var st_att =  filterData(getDatarange(4,lastc-9,lastr,3));
  var late =    filterData(getDatarange(4,lastc-12,lastr,3));
  var asn =     filterData(getDatarange(4,lastc-18,lastr,6));
  var st_data = concat(st_data,st_att,late,asn);
  Logger.log(st_data);
  div_data(st_data);
}
