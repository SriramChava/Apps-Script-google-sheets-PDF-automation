//**************************** DO NOT MODIFY CELL LOCATIONS IN EXCEL AS THE CODE WILL STOP WORKING CORRECTLY***********************/

// 1. DEFINING GLOBAL VARIABLES

// 1.1 Google doc link for pdf template of Fee Structure report:https://docs.google.com/document/d/1xRvMFsm-TzDZQHm-6aLY2PRthWtb6e4LAcv5q7OdCIE/edit
var dc2_file = DriveApp.getFileById("1xRvMFsm-TzDZQHm-6aLY2PRthWtb6e4LAcv5q7OdCIE"); 
// 1.2 Temporary folder which holds the above doc for each student this doc is copied, edited and deleted for each student. This is to make sure that the original doc is not changed as then each report will be modified with the earlier changes. After the doc is converted to pdf and mailed the temporary doc is generated again for each report: https://drive.google.com/drive/u/0/folders/1MtuLu3NCbJYh5C5RsWpjDi1KYnxHv-Lf
  var Tmp2_Folder =  DriveApp.getFolderById("1MtuLu3NCbJYh5C5RsWpjDi1KYnxHv-Lf");
// 1.3 This folder is responsible for holding the reports of all the students. :  https://drive.google.com/drive/u/0/folders/1wQBF6bMangBnxFvgbJRI6ZKdgm-jitkv
  var PDF2_Foldr =DriveApp.getFolderById("1wQBF6bMangBnxFvgbJRI6ZKdgm-jitkv");
//1.4 google sheet holding attendance,and assignment submission of all the students we get the sheet by name and the name of the sheet is "Attendance" : https://docs.google.com/spreadsheets/d/1M6wvyTyRvZMWFkGV-TTRf9bjkyG53iE8QDPNDdtU894/edit?usp=sharing 
  var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Class 11th 2YR");

// upto 2 decimal value 
function typecast2(numb) {
  numb = numb.toFixed(1);
  return numb;
}
// Global variables
  var nam2 = sheet2.getRange("A8").getValue();
  var rol2 = sheet2.getRange("A9").getValue();
  var mai2 = sheet2.getRange("A10").getValue();
  var totl_fees2 = typecast2(sheet2.getRange("D16").getValue()); 
  var totl_dis2 = typecast2(sheet2.getRange("D24").getValue());
  var gros_fee2 = typecast2(sheet2.getRange("D25").getValue());
  var gs_fee2 = typecast2(sheet2.getRange("D26").getValue());
  var nt_fee2 = typecast2(sheet2.getRange("D27").getValue());
  
  var ist_12 = typecast2(sheet2.getRange("C31").getValue());
  var ist_22 = typecast2(sheet2.getRange("C32").getValue());
  var cur_date2 = sheet2.getRange("B31").getValue();
  var i2_date2 = sheet2.getRange("B32").getValue();
  var i3_date2 = sheet2.getRange("B33").getValue();
  var i4_date2 = sheet2.getRange("B34").getValue();
  var i5_date2 = sheet2.getRange("B35").getValue();
  var i6_date2 = sheet2.getRange("B36").getValue();
  var i7_date2 = sheet2.getRange("B37").getValue();

  var s_date2 = sheet2.getRange("D8").getValue();
  var e_date2 = sheet2.getRange("E8").getValue();
  var adm_date2 = sheet2.getRange("D9").getValue();
 

//2.1 This function takes the data range in the excel as its input and the data for attendance, and assignments are segregated
function getDatarange2(firstr,firstc,lastr,lastc) 
{
  return sheet2.getRange(firstr,firstc,lastr,lastc).getValues();
}

function div_data2(fee_data, disc_data){
   var add_fe2 = typecast2(fee_data[0][1]);
   var add_per2 = fee_data[0][0]*100;
  
   var tu_fe2 = typecast2(fee_data[1][1]);
   var tu_per2 = fee_data[1][0]*100;
  
   var bok_fe2 = typecast2(fee_data[2][1]);
   var bok_per2 = fee_data[2][0]*100;
 
   var if_fe2 = typecast2(fee_data[3][1]);
   var if_per2 = fee_data[3][0]*100;
   
   var th_fe2 = typecast2(disc_data[0][1]);
   var th_per2 = disc_data[0][0]*100;
   

   var rk_fe2 = typecast2(disc_data[1][1]);
   var rk_per2 = disc_data[1][0]*100;
  
   var pr_fe2 = typecast2(disc_data[2][1]);
   var pr_per2 = disc_data[2][0]*100;
 

   var rf_fe2 = typecast2(disc_data[3][1]);
   var rf_per2 = disc_data[3][0]*100;
  
   var ls_fe2 = typecast2(disc_data[4][1]);
   var ls_per2 = disc_data[4][0]*100;
   
   var sh_fe2 = typecast2(disc_data[5][1]);
   var sh_per2 = disc_data[5][0]*100;
 

   create_PDFnMail2(add_fe2,add_per2,tu_fe2,tu_per2,bok_fe2,bok_per2,if_fe2,if_per2,th_fe2,th_per2,rf_fe2,rf_per2,ls_fe2,ls_per2,sh_fe2,sh_per2,rk_fe2,rk_per2,pr_fe2,pr_per2,dc2_file,Tmp2_Folder,PDF2_Foldr)
}

// CREATE AND MAIL FUNCTION

function create_PDFnMail2(add_fe2,add_per2,tu_fe2,tu_per2,bok_fe2,bok_per2,if_fe2,if_per2,th_fe2,th_per2,rf_fe2,rf_per2,ls_fe2,ls_per2,sh_fe2,sh_per2,rk_fe2,rk_per2,pr_fe2,pr_per2,dc2_file,Tmp2_Folder,PDF2_Foldr)
{
  var temp_File = dc2_file.makeCopy(Tmp3_Folder);
  var temp_doc_file = DocumentApp.openById(temp_File.getId());
  var body = temp_doc_file.getBody();
  body.replaceText("{s_name}",nam2);
  body.replaceText("{r_num}",rol2);
  body.replaceText("{tt_fee}",totl_fees2);
  
  body.replaceText("{tt_dis}",totl_dis2);

  body.replaceText("{gros_fee}",gros_fee2);

  body.replaceText("{add_fee}",add_fe2);
 
  body.replaceText("{add_per}",add_per2);
  body.replaceText("{tut_fee}",tu_fe2);
 
  body.replaceText("{tut_per}",tu_per2);
  body.replaceText("{bok_fee}",bok_fe2);
 
  body.replaceText("{md_per}",bok_per2);
  body.replaceText("{ifr_fee}",if_fe2);

  body.replaceText("{tch_per}",if_per2);
 
  body.replaceText("{tc_dis}",th_fe2);
  body.replaceText("{lc_per}",th_per2);
 
  
  body.replaceText("{rnk_dis}",rk_fe2);
  body.replaceText("{rnk_per}",rk_per2);
  body.replaceText("{oth_dis}",pr_fe2);
  body.replaceText("{oth_per}",pr_per2);
 

  body.replaceText("{rf_dis}",rf_fe2);
  body.replaceText("{rf_per}",rf_per2);

  body.replaceText("{lns_fee}",ls_fe2);
  body.replaceText("{ls_per}",ls_per2);
 
  body.replaceText("{sh_dis}",sh_fe2);
  body.replaceText("{sh_per}",sh_per2);
 
  body.replaceText("{ist_1}",ist_12);
  body.replaceText("{ist_2}",ist_22);
  body.replaceText("{ist_3}",ist_22);
  body.replaceText("{ist_4}",ist_22);
  body.replaceText("{ist_5}",ist_22);
  body.replaceText("{ist_6}",ist_22);
  body.replaceText("{ist_7}",ist_22);

  body.replaceText("{nt_fee}",nt_fee2);
  
  body.replaceText("{cur_date}",cur_date2);
  body.replaceText("{ist_2_date}",i2_date2);
  body.replaceText("{ist_3_date}",i3_date2);
  body.replaceText("{ist_4_date}",i4_date2);
  body.replaceText("{ist_5_date}",i5_date2);
  body.replaceText("{ist_6_date}",i6_date2);
  body.replaceText("{ist_7_date}",i7_date2);

  body.replaceText("{gs}",gs_fee2);

  body.replaceText("{s_date}",s_date2);
  body.replaceText("{e_date}",e_date2);
  body.replaceText("{ad_test}",adm_date2);
  temp_doc_file.saveAndClose();

  var PDF_content = temp_File.getAs(MimeType.PDF);
  var PDF_File = PDF2_Foldr.createFile(PDF_content).setName("Admission Report " + nam2 + " " + rol2);
  Tmp2_Folder.removeFile(temp_File);

/* FUNCTION FOR MAILING UNCOMMENT WHENEVER WEBAPP IS READY*/
  // MailApp.sendEmail(mai2, 'Admission Report', "The admission report for student " + nam2 + " can be found as follows:",   {               
  //   attachments: [PDF_File.getAs(MimeType.PDF)],
  //   name : 'Progressive minds automated emailer'
  // });

  Logger.log(PDF_File)
 }


  function getdata2 () {
    var fee_data = getDatarange2(12, 3, 4, 2);
    var disc_data = getDatarange2(18, 3, 6, 2);
     div_data2(fee_data, disc_data);
    // Logger.log(fee_data);
      //Logger.log(disc_data);
  }


