//**************************** DO NOT MODIFY CELL LOCATIONS IN EXCEL AS THE CODE WILL STOP WORKING CORRECTLY***********************/

// 1. DEFINING GLOBAL VARIABLES

// 1.1 Google doc link for pdf template of Fee Structure report:https://docs.google.com/document/d/1slI-aZzDdVyDIQ_k82OA2hqf634tLMr6Y9KqNxQ27hU/edit
var dc3_file = DriveApp.getFileById("1slI-aZzDdVyDIQ_k82OA2hqf634tLMr6Y9KqNxQ27hU"); 
// 1.2 Temporary folder which holds the above doc for each student this doc is copied, edited and deleted for each student. This is to make sure that the original doc is not changed as then each report will be modified with the earlier changes. After the doc is converted to pdf and mailed the temporary doc is generated again for each report: https://drive.google.com/drive/u/0/folders/1MtuLu3NCbJYh5C5RsWpjDi1KYnxHv-Lf
  var Tmp3_Folder =  DriveApp.getFolderById("1MtuLu3NCbJYh5C5RsWpjDi1KYnxHv-Lf");
// 1.3 This folder is responsible for holding the reports of all the students. :  https://drive.google.com/drive/u/0/folders/1wQBF6bMangBnxFvgbJRI6ZKdgm-jitkv
  var PDF3_Foldr =DriveApp.getFolderById("1wQBF6bMangBnxFvgbJRI6ZKdgm-jitkv");
//1.4 google sheet holding attendance,and assignment submission of all the students we get the sheet by name and the name of the sheet is "Attendance" : https://docs.google.com/spreadsheets/d/1M6wvyTyRvZMWFkGV-TTRf9bjkyG53iE8QDPNDdtU894/edit?usp=sharing 
  var sheet3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Class 10th");

// upto 2 decimal value 
function typecast3(numb) {
  numb = numb.toFixed(1);
  return numb;
}
// Global variables
  var nam3 = sheet3.getRange("A8").getValue();
  var rol3 = sheet3.getRange("A9").getValue();
  var mai3 = sheet3.getRange("A10").getValue();
  var totl_fees3 = typecast3(sheet3.getRange("C16").getValue()); 
  var totl1_fees3 = typecast3(sheet3.getRange("D16").getValue()); 
  var totl_dis3 = typecast3(sheet3.getRange("C24").getValue());
  var totl1_dis3 = typecast3(sheet3.getRange("D24").getValue());
  var gros_fee3 = typecast3(sheet3.getRange("C25").getValue());
  var gros1_fee3 = typecast3(sheet3.getRange("D25").getValue());
  var gs_fee3 = typecast3(sheet3.getRange("C26").getValue());
  var gs1_fee3 = typecast3(sheet3.getRange("D26").getValue());
  var nt_fee3 = typecast3(sheet3.getRange("C27").getValue());
  var nt1_fee3 = typecast3(sheet3.getRange("D27").getValue());
  var ist_13 = typecast3(sheet3.getRange("C31").getValue());
  var ist_23 = typecast3(sheet3.getRange("C32").getValue());
  var cur_date3 = sheet3.getRange("B31").getValue();
  var i2_date3 = sheet3.getRange("B32").getValue();
  var i3_date3 = sheet3.getRange("B33").getValue();
  var i4_date3 = sheet3.getRange("B34").getValue();
  var i5_date3 = sheet3.getRange("B39").getValue();
  var i6_date3 = sheet3.getRange("B40").getValue();
  var i7_date3 = sheet3.getRange("B41").getValue();
  var i8_date3 = sheet3.getRange("B42").getValue();
  var i9_date3 = sheet3.getRange("B43").getValue();
  var i10_date3 = sheet3.getRange("B44").getValue();
  var i11_date3 = sheet3.getRange("B45").getValue();
  var s_date3 = sheet3.getRange("D8").getValue();
  var e_date3 = sheet3.getRange("E8").getValue();
  var adm_date3 = sheet3.getRange("D9").getValue();
  var ist_53 = typecast3(sheet3.getRange("C39").getValue());
  var ist_63 = typecast3(sheet3.getRange("C40").getValue());

//2.1 This function takes the data range in the excel as its input and the data for attendance, and assignments are segregated
function getDatarange3(firstr,firstc,lastr,lastc) 
{
  return sheet3.getRange(firstr,firstc,lastr,lastc).getValues();
}

function div_data3(fee_data, disc_data){
   var add_fe3 = typecast3(fee_data[0][1]);
   var add_per3 = fee_data[0][0]*100;
   var ad1_fe3 = typecast3(fee_data[0][2]);
   var tu_fe3 = typecast3(fee_data[1][1]);
   var tu_per3 = fee_data[1][0]*100;
   var tut1_fe3 = typecast3(fee_data[1][2]);
   var bok_fe3 = typecast3(fee_data[2][1]);
   var bok_per3 = fee_data[2][0]*100;
   var book1_fe3 = typecast3(fee_data[2][2]);
   var if_fe3 = typecast3(fee_data[3][1]);
   var if_per3 = fee_data[3][0]*100;
   var inf1_fe3 = typecast3(fee_data[3][2]);
   var th_fe3 = typecast3(disc_data[0][1]);
   var th_per3 = disc_data[0][0]*100;
   var tch1_fe3 = typecast3(disc_data[0][2]);

   var rk_fe3 = typecast3(disc_data[1][1]);
   var rk_per3 = disc_data[1][0]*100;
   var rnk1_fe3 = typecast3(disc_data[1][2]);
   var pr_fe3 = typecast3(disc_data[2][1]);
   var pr_per3 = disc_data[2][0]*100;
   var par1_fe3 = typecast3(disc_data[2][2]);

   var rf_fe3 = typecast3(disc_data[3][1]);
   var rf_per3 = disc_data[3][0]*100;
   var ref1_fe3 = typecast3(disc_data[3][2]);
   var ls_fe3 = typecast3(disc_data[4][1]);
   var ls_per3 = disc_data[4][0]*100;
   var lns1_fe3 = typecast3(disc_data[4][2]);
   var sh_fe3 = typecast3(disc_data[5][1]);
   var sh_per3 = disc_data[5][0]*100;
   var sch1_fe3 = typecast3(disc_data[5][2]);

   create_PDFnMail3(add_fe3,add_per3,tu_fe3,tu_per3,bok_fe3,bok_per3,if_fe3,if_per3,th_fe3,th_per3,rf_fe3,rf_per3,ls_fe3,ls_per3,sh_fe3,sh_per3,ad1_fe3,tut1_fe3,book1_fe3,inf1_fe3,rk_fe3,rk_per3,rnk1_fe3,pr_fe3,pr_per3,par1_fe3,ref1_fe3,lns1_fe3,sch1_fe3,tch1_fe3,dc3_file,Tmp3_Folder,PDF3_Foldr)
}

// CREATE AND MAIL FUNCTION

function create_PDFnMail3(add_fe3,add_per3,tu_fe3,tu_per3,bok_fe3,bok_per3,if_fe3,if_per3,th_fe3,th_per3,rf_fe3,rf_per3,ls_fe3,ls_per3,sh_fe3,sh_per3,ad1_fe3,tut1_fe3,book1_fe3,inf1_fe3,rk_fe3,rk_per3,rnk1_fe3,pr_fe3,pr_per3,par1_fe3,ref1_fe3,lns1_fe3,sch1_fe3,tch1_fe3,dc3_file,Tmp3_Folder,PDF3_Foldr)
{
  var temp_File = dc3_file.makeCopy(Tmp3_Folder);
  var temp_doc_file = DocumentApp.openById(temp_File.getId());
  var body = temp_doc_file.getBody();
  body.replaceText("{s_name}",nam3);
  body.replaceText("{r_num}",rol3);
  body.replaceText("{tt_fee}",totl_fees3);
  body.replaceText("{tot1_fee}",totl1_fees3);
  body.replaceText("{tt_dis}",totl_dis3);
  body.replaceText("{tot1_dis}",totl1_dis3);
  body.replaceText("{gros_fee}",gros_fee3);
  body.replaceText("{gross1_fee}",gros1_fee3);
  body.replaceText("{add_fee}",add_fe3);
  body.replaceText("{ad1_fee}",ad1_fe3);
  body.replaceText("{add_per}",add_per3);
  body.replaceText("{tut_fee}",tu_fe3);
  body.replaceText("{tut1_fee}",tut1_fe3);
  body.replaceText("{tut_per}",tu_per3);
  body.replaceText("{bok_fee}",bok_fe3);
  body.replaceText("{book1_fee}",book1_fe3);
  body.replaceText("{md_per}",bok_per3);
  body.replaceText("{ifr_fee}",if_fe3);
  body.replaceText("{infr1_fee}",inf1_fe3);
  body.replaceText("{tch_per}",if_per3);
  body.replaceText("{tec1_dis}",tch1_fe3);
  body.replaceText("{tc_dis}",th_fe3);
  body.replaceText("{lc_per}",th_per3);
 
  body.replaceText("{rnk1_dis}",rnk1_fe3);
  body.replaceText("{rnk_dis}",rk_fe3);
  body.replaceText("{rnk_per}",rk_per3);
  body.replaceText("{oth_dis}",pr_fe3);
  body.replaceText("{oth_per}",pr_per3);
  body.replaceText("{oth1_dis}",par1_fe3);

  body.replaceText("{rf_dis}",rf_fe3);
  body.replaceText("{rf_per}",rf_per3);
  body.replaceText("{ref1_dis}",ref1_fe3);
  body.replaceText("{lns_fee}",ls_fe3);
  body.replaceText("{ls_per}",ls_per3);
  body.replaceText("{ls1_fee}",lns1_fe3);
  body.replaceText("{sh_dis}",sh_fe3);
  body.replaceText("{sh_per}",sh_per3);
  body.replaceText("{sch1_dis}",sch1_fe3);
  body.replaceText("{ist_1}",ist_13);
  body.replaceText("{ist_2}",ist_23);
  body.replaceText("{ist_3}",ist_23);
  body.replaceText("{ist_4}",ist_23);
  body.replaceText("{inst_5}",ist_53);
  body.replaceText("{inst_6}",ist_63);
  body.replaceText("{inst_7}",ist_63);
  body.replaceText("{inst_8}",ist_63);
  body.replaceText("{inst_9}",ist_63);
  body.replaceText("{inst_10}",ist_63);
  body.replaceText("{inst_11}",ist_63);
  body.replaceText("{nt_fee}",nt_fee3);
  body.replaceText("{net1_fee}",nt1_fee3);
  body.replaceText("{cur_date}",cur_date3);
  body.replaceText("{ist_2_date}",i2_date3);
  body.replaceText("{ist_3_date}",i3_date3);
  body.replaceText("{ist_4_date}",i4_date3);
  body.replaceText("{ist_5_date}",i5_date3);
  body.replaceText("{ist_6_date}",i6_date3);
  body.replaceText("{ist_7_date}",i7_date3);
  body.replaceText("{ist_8_date}",i8_date3);
  body.replaceText("{ist_9_date}",i9_date3);
  body.replaceText("{ist_10_date}",i10_date3);
   body.replaceText("{ist_11_date}",i11_date3);
  body.replaceText("{gs}",gs_fee3);
  body.replaceText("{gst1}",gs1_fee3);
  body.replaceText("{s_date}",s_date3);
  body.replaceText("{e_date}",e_date3);
  body.replaceText("{ad_test}",adm_date3);
  temp_doc_file.saveAndClose();

  var PDF_content = temp_File.getAs(MimeType.PDF);
  var PDF_File = PDF3_Foldr.createFile(PDF_content).setName("Admission Report " + nam3 + " " + rol3);
  Tmp3_Folder.removeFile(temp_File);

/* FUNCTION FOR MAILING UNCOMMENT WHENEVER WEBAPP IS READY*/
  // MailApp.sendEmail(mai3, 'Admission Report', "The admission report for student " + nam3 + " can be found as follows:",   {               
  //   attachments: [PDF_File.getAs(MimeType.PDF)],
  //   name : 'Progressive minds automated emailer'
  // });

  Logger.log(PDF_File)
 }


  function getdata3 () {
    var fee_data = getDatarange3(12, 2, 4, 3);
    var disc_data = getDatarange3(18, 2, 6, 3);
     div_data3(fee_data, disc_data);
    // Logger.log(fee_data);
      //Logger.log(disc_data);
  }

