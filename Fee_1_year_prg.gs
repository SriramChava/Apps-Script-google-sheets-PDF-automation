//**************************** DO NOT MODIFY CELL LOCATIONS IN EXCEL AS THE CODE WILL STOP WORKING CORRECTLY***********************/

// 1. DEFINING GLOBAL VARIABLES

// 1.1 Google doc link for pdf template of Fee Structure report :https://docs.google.com/document/d/1vUs9lt6Io0QODGZZjRqLVI_tl9F2xQMHJ7CMhPoLbwk/edit
var doc_file = DriveApp.getFileById("1vUs9lt6Io0QODGZZjRqLVI_tl9F2xQMHJ7CMhPoLbwk"); 
// 1.2 Temporary folder which holds the above doc for each student this doc is copied, edited and deleted for each student. This is to make sure that the original doc is not changed as then each report will be modified with the earlier changes. After the doc is converted to pdf and mailed the temporary doc is generated again for each report: https://drive.google.com/drive/u/0/folders/1MtuLu3NCbJYh5C5RsWpjDi1KYnxHv-Lf
  var Temp_Folder =  DriveApp.getFolderById("1MtuLu3NCbJYh5C5RsWpjDi1KYnxHv-Lf");
// 1.3 This folder is responsible for holding the reports of all the students. :  https://drive.google.com/drive/u/0/folders/1wQBF6bMangBnxFvgbJRI6ZKdgm-jitkv
  var PDF_Folder =DriveApp.getFolderById("1wQBF6bMangBnxFvgbJRI6ZKdgm-jitkv");
//1.4 google sheet holding attendance,and assignment submission of all the students we get the sheet by name and the name of the sheet is "Attendance" : https://docs.google.com/spreadsheets/d/1M6wvyTyRvZMWFkGV-TTRf9bjkyG53iE8QDPNDdtU894/edit?usp=sharing 
  var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Class 12th 1YR");

// upto 2 decimal value 
function typecast1(numb) {
  numb = numb.toFixed(1);
  return numb;
}
// Global variables
  var name1 = sheet1.getRange("B8").getValue();
  var roll1 = sheet1.getRange("B9").getValue();
  var mail1 = sheet1.getRange("B10").getValue();
  var total_fees1 = typecast1(sheet1.getRange("D16").getValue()); 
  var total_dis1 = typecast1(sheet1.getRange("D24").getValue());
  var gross_fee1 = typecast1(sheet1.getRange("D25").getValue());
  var gst_fee1 = typecast1(sheet1.getRange("D26").getValue());
  var net_fee1 = typecast1(sheet1.getRange("D27").getValue());
  var inst_11 = typecast1(sheet1.getRange("C31").getValue());
  var inst_21 = typecast1(sheet1.getRange("C32").getValue());
  var curr_date1 = sheet1.getRange("B31").getValue();
  var in2_date1 = sheet1.getRange("B32").getValue();
  var in3_date1 = sheet1.getRange("B33").getValue();
  var in4_date1 = sheet1.getRange("B34").getValue();
  var st_date1 = sheet1.getRange("E8").getValue();
  var ed_date1 = sheet1.getRange("F8").getValue();

//2.1 This function takes the data range in the excel as its input and the data for attendance, and assignments are segregated
function getDatarange1(firstr,firstc,lastr,lastc) 
{
  return sheet1.getRange(firstr,firstc,lastr,lastc).getValues();
}

function div_data1(fee_data, disc_data){
   var ad_fe1 = typecast1(fee_data[0][1]);
   //Logger.log(ad_fe);
   var ad_per1 = fee_data[0][0]*100;
  //  Logger.log(ad_per);
   var tut_fe1 = typecast1(fee_data[1][1]);
   var tut_per1 = fee_data[1][0]*100;
   var book_fe1 = typecast1(fee_data[2][1]);
   var book_per1 = fee_data[2][0]*100;
   var inf_fe1 = typecast1(fee_data[3][1]);
   var inf_per1 = fee_data[3][0]*100;
   var tch_fe1 = typecast1(disc_data[0][1]);
   var tch_per1 = disc_data[0][0]*100;
   var rnk_fe1 = typecast1(disc_data[1][1]);
   var rnk_per1 = disc_data[1][0]*100;
   var par_fe1 = typecast1(disc_data[2][1]);
   var par_per1 = disc_data[2][0]*100;
   var ref_fe1 = typecast1(disc_data[3][1]);
   var ref_per1 = disc_data[3][0]*100;
   var lns_fe1 = typecast1(disc_data[4][1]);
   var lns_per1 = disc_data[4][0]*100;
   var sch_fe1 = typecast1(disc_data[5][1]);
   var sch_per1 = disc_data[5][0]*100;

   create_PDFnMail(ad_fe1,ad_per1,tut_fe1,tut_per1,book_fe1,book_per1,inf_fe1,inf_per1,tch_fe1,tch_per1,rnk_fe1,rnk_per1,par_fe1,par_per1,ref_fe1,ref_per1,lns_fe1,lns_per1,sch_fe1,sch_per1,doc_file,Temp_Folder,PDF_Folder)
}

// CREATE AND MAIL FUNCTION

function create_PDFnMail(ad_fe1,ad_per1,tut_fe1,tut_per1,book_fe1,book_per1,inf_fe1,inf_per1,tch_fe1,tch_per1,rnk_fe1,rnk_per1,par_fe1,par_per1,ref_fe1,ref_per1,lns_fe1,lns_per1,sch_fe1,sch_per1,doc_file,Temp_Folder,PDF_Folder)
{
  var temp_File = doc_file.makeCopy(Temp_Folder);
  var temp_doc_file = DocumentApp.openById(temp_File.getId());
  var body = temp_doc_file.getBody();
  body.replaceText("{st_name}",name1);
  body.replaceText("{rl_num}",roll1);
  body.replaceText("{tot_fee}",total_fees1);
  body.replaceText("{tot_dis}",total_dis1);
  body.replaceText("{gross_fee}",gross_fee1);
  body.replaceText("{ad_fee}",ad_fe1);
  body.replaceText("{ad_per}",ad_per1);
  body.replaceText("{tut_fee}",tut_fe1);
  body.replaceText("{tut_per}",tut_per1);
  body.replaceText("{book_fee}",book_fe1);
  body.replaceText("{mod_per}",book_per1);
  body.replaceText("{infr_fee}",inf_fe1);
  body.replaceText("{tech_per}",inf_per1);
  body.replaceText("{tec_dis}",tch_fe1);
  body.replaceText("{lec_per}",tch_per1);
  body.replaceText("{rnk_dis}",rnk_fe1);
  body.replaceText("{rnk_per}",rnk_per1);
  body.replaceText("{oth_dis}",par_fe1);  
  body.replaceText("{oth_per}",par_per1);
  body.replaceText("{ref_dis}",ref_fe1);
  body.replaceText("{ref_per}",ref_per1);
  body.replaceText("{ls_fee}",lns_fe1);
  body.replaceText("{lns_per}",lns_per1);
  body.replaceText("{sch_dis}",sch_fe1);
  body.replaceText("{sch_per}",sch_per1);
  body.replaceText("{inst_1}",inst_11);
  body.replaceText("{inst_2}",inst_21);
  body.replaceText("{inst_3}",inst_21);
  body.replaceText("{inst_4}",inst_21);
  body.replaceText("{net_fee}",net_fee1);
  body.replaceText("{curr_date}",curr_date1);
  body.replaceText("{inst_2_date}",in2_date1);
  body.replaceText("{inst_3_date}",in3_date1);
  body.replaceText("{inst_4_date}",in4_date1);
  body.replaceText("{gst}",gst_fee1);
  body.replaceText("{st_date}",st_date1);
  body.replaceText("{ed_date}",ed_date1);
  temp_doc_file.saveAndClose();

  var PDF_content = temp_File.getAs(MimeType.PDF);
  var PDF_File = PDF_Folder.createFile(PDF_content).setName("Admission Report " + name1 + " " + roll1);
  Temp_Folder.removeFile(temp_File);

/* FUNCTION FOR MAILING UNCOMMENT WHENEVER WEBAPP IS READY*/
  // MailApp.sendEmail(mail1, 'Admission Report', "The admission report for student " + name + " can be found as follows:",   {               
  //   attachments: [PDF_File.getAs(MimeType.PDF)],
  //   name : 'Progressive minds automated emailer'
  // });

  Logger.log(PDF_File)
 }


  function getdata1 () {
    var fee_data1 = getDatarange1(12, 3, 4, 2);
    var disc_data1 = getDatarange1(18, 3, 6, 2);
    div_data1(fee_data1, disc_data1);
   // Logger.log(fee_data);
    //Logger.log(disc_data);
  }
