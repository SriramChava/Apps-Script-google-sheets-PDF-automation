//**************************** DO NOT MODIFY CELL LOCATIONS IN EXCEL AS THE CODE WILL STOP WORKING CORRECTLY***********************/

// 1. DEFINING GLOBAL VARIABLES

// 1.1 Google doc link for pdf template of Fee Structure report :https://docs.google.com/document/d/xxxxxxxxxxxxxxxxxxxxxxx/edit
var dc_file = DriveApp.getFileById("Your Google Docs Template"); 
// 1.2 Temporary folder which holds the above doc for each student this doc is copied, edited and deleted for each student. This is to make sure that the original doc is not changed as then each report will be modified with the earlier changes. After the doc is converted to pdf and mailed the temporary doc is generated again for each report: https://drive.google.com/drive/u/0/folders/xxxxxxxxxxxxxxxxxxxxxxxxx
  var Tmp_Folder =  DriveApp.getFolderById("Drive Temp Folder");
// 1.3 This folder is responsible for holding the reports of all the students. :  https://drive.google.com/drive/u/0/folders/xxxxxxxxxxxxxxxxxx
  var PDF_Foldr =DriveApp.getFolderById("Drve Main Folder");
//1.4 google sheet holding attendance,and assignment submission of all the students we get the sheet by name and the name of the sheet is "Attendance" : https://docs.google.com/spreadsheets/d/xxxxxxxxxxxxxxxxxxxxxxxxxxxxx/edit?usp=sharing 
  var sheet9 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet Name");

// upto 2 decimal value 
function typecast4(numb) {
  numb = numb.toFixed(1);
  return numb;
}
// Global variables
  var nam = sheet9.getRange("A8").getValue();
  var rol = sheet9.getRange("A9").getValue();
  var mai = sheet9.getRange("A10").getValue();
  var totl_fees = typecast4(sheet9.getRange("C16").getValue()); 
  var totl1_fees = typecast4(sheet9.getRange("D16").getValue()); 
  var totl_dis = typecast4(sheet9.getRange("C24").getValue());
  var totl1_dis = typecast4(sheet9.getRange("D24").getValue());
  var gros_fee = typecast4(sheet9.getRange("C25").getValue());
  var gros1_fee = typecast4(sheet9.getRange("D25").getValue());
  var gs_fee = typecast4(sheet9.getRange("C26").getValue());
  var gs1_fee = typecast4(sheet9.getRange("D26").getValue());
  var nt_fee = typecast4(sheet9.getRange("C27").getValue());
  var nt1_fee = typecast4(sheet9.getRange("D27").getValue());
  var ist_1 = typecast4(sheet9.getRange("C31").getValue());
  var ist_2 = typecast4(sheet9.getRange("C32").getValue());
  var cur_date = sheet9.getRange("B31").getValue();
  var i2_date = sheet9.getRange("B32").getValue();
  var i3_date = sheet9.getRange("B33").getValue();
  var i4_date = sheet9.getRange("B34").getValue();
  var i5_date = sheet9.getRange("B39").getValue();
  var i6_date = sheet9.getRange("B40").getValue();
  var i7_date = sheet9.getRange("B41").getValue();
  var i8_date = sheet9.getRange("B42").getValue();
  var s_date = sheet9.getRange("D8").getValue();
  var e_date = sheet9.getRange("E8").getValue();
  var adm_date = sheet9.getRange("D9").getValue();
  var ist_5 = typecast4(sheet9.getRange("C39").getValue());
  var ist_6 = typecast4(sheet9.getRange("C40").getValue());

//2.1 This function takes the data range in the excel as its input and the data for attendance, and assignments are segregated
function getDatarange4(firstr,firstc,lastr,lastc) 
{
  return sheet9.getRange(firstr,firstc,lastr,lastc).getValues();
}

function div_data4(fee_data, disc_data){
   var add_fe = typecast4(fee_data[0][1]);
   var add_per = fee_data[0][0]*100;
   var ad1_fe = typecast4(fee_data[0][2]);
   var tu_fe = typecast4(fee_data[1][1]);
   var tu_per = fee_data[1][0]*100;
   var tut1_fe = typecast4(fee_data[1][2]);
   var bok_fe = typecast4(fee_data[2][1]);
   var bok_per = fee_data[2][0]*100;
   var book1_fe = typecast4(fee_data[2][2]);
   var if_fe = typecast4(fee_data[3][1]);
   var if_per = fee_data[3][0]*100;
   var inf1_fe = typecast4(fee_data[3][2]);
   var th_fe = typecast4(disc_data[0][1]);
   var th_per = disc_data[0][0]*100;
   var tch1_fe = typecast4(disc_data[0][2]);

   var rk_fe = typecast4(disc_data[1][1]);
   var rk_per = disc_data[1][0]*100;
   var rnk1_fe = typecast4(disc_data[1][2]);
   var pr_fe = typecast4(disc_data[2][1]);
   var pr_per = disc_data[2][0]*100;
   var par1_fe = typecast4(disc_data[2][2]);

   var other_fee = Number(rk_fe) + Number(pr_fe);
   var other_per = Number(rk_per) + Number(pr_per);
   var other1_fee = Number(rnk1_fe) + Number(par1_fe);

   var rf_fe = typecast4(disc_data[3][1]);
   var rf_per = disc_data[3][0]*100;
   var ref1_fe = typecast4(disc_data[3][2]);
   var ls_fe = typecast4(disc_data[4][1]);
   var ls_per = disc_data[4][0]*100;
   var lns1_fe = typecast4(disc_data[4][2]);
   var sh_fe = typecast4(disc_data[5][1]);
   var sh_per = disc_data[5][0]*100;
   var sch1_fe = typecast4(disc_data[5][2]);

   create_PDFnMail4(add_fe,add_per,tu_fe,tu_per,bok_fe,bok_per,if_fe,if_per,th_fe,th_per,rf_fe,rf_per,ls_fe,ls_per,sh_fe,sh_per,ad1_fe,tut1_fe,book1_fe,inf1_fe,other1_fee,other_fee,other_per,ref1_fe,lns1_fe,sch1_fe,tch1_fe,dc_file,Tmp_Folder,PDF_Foldr)
}

// CREATE AND MAIL FUNCTION

function create_PDFnMail4(add_fe,add_per,tu_fe,tu_per,bok_fe,bok_per,if_fe,if_per,th_fe,th_per,rf_fe,rf_per,ls_fe,ls_per,sh_fe,sh_per,ad1_fe,tut1_fe,book1_fe,inf1_fe,other1_fee,other_fee,other_per,ref1_fe,lns1_fe,sch1_fe,tch1_fe,dc_file,Tmp_Folder,PDF_Foldr)
{
  var temp_File = dc_file.makeCopy(Tmp_Folder);
  var temp_doc_file = DocumentApp.openById(temp_File.getId());
  var body = temp_doc_file.getBody();
  body.replaceText("{s_name}",nam);
  body.replaceText("{r_num}",rol);
  body.replaceText("{tt_fee}",totl_fees);
  body.replaceText("{tot1_fee}",totl1_fees);
  body.replaceText("{tt_dis}",totl_dis);
  body.replaceText("{tot1_dis}",totl1_dis);
  body.replaceText("{gros_fee}",gros_fee);
  body.replaceText("{gross1_fee}",gros1_fee);
  body.replaceText("{add_fee}",add_fe);
  body.replaceText("{ad1_fee}",ad1_fe);
  body.replaceText("{add_per}",add_per);
  body.replaceText("{tut_fee}",tu_fe);
  body.replaceText("{tut1_fee}",tut1_fe);
  body.replaceText("{tut_per}",tu_per);
  body.replaceText("{bok_fee}",bok_fe);
  body.replaceText("{book1_fee}",book1_fe);
  body.replaceText("{md_per}",bok_per);
  body.replaceText("{ifr_fee}",if_fe);
  body.replaceText("{infr1_fee}",inf1_fe);
  body.replaceText("{tch_per}",if_per);
  body.replaceText("{tec1_dis}",tch1_fe);
  body.replaceText("{tc_dis}",th_fe);
  body.replaceText("{lc_per}",th_per);
  body.replaceText("{ot_dis}",other_fee);
  body.replaceText("{ot_per}",other_per);
  body.replaceText("{oth1_dis}",other1_fee);
  body.replaceText("{rf_dis}",rf_fe);
  body.replaceText("{rf_per}",rf_per);
  body.replaceText("{ref1_dis}",ref1_fe);
  body.replaceText("{lns_fee}",ls_fe);
  body.replaceText("{ls_per}",ls_per);
  body.replaceText("{ls1_fee}",lns1_fe);
  body.replaceText("{sh_dis}",sh_fe);
  body.replaceText("{sh_per}",sh_per);
  body.replaceText("{sch1_dis}",sch1_fe);
  body.replaceText("{ist_1}",ist_1);
  body.replaceText("{ist_2}",ist_2);
  body.replaceText("{ist_3}",ist_2);
  body.replaceText("{ist_4}",ist_2);
  body.replaceText("{inst_5}",ist_5);
  body.replaceText("{inst_6}",ist_6);
  body.replaceText("{inst_7}",ist_6);
  body.replaceText("{inst_8}",ist_6);
  body.replaceText("{nt_fee}",nt_fee);
  body.replaceText("{net1_fee}",nt1_fee);
  body.replaceText("{cur_date}",cur_date);
  body.replaceText("{ist_2_date}",i2_date);
  body.replaceText("{ist_3_date}",i3_date);
  body.replaceText("{ist_4_date}",i4_date);
  body.replaceText("{ist_5_date}",i5_date);
  body.replaceText("{ist_6_date}",i6_date);
  body.replaceText("{ist_7_date}",i7_date);
   body.replaceText("{ist_8_date}",i8_date);
  body.replaceText("{gs}",gs_fee);
  body.replaceText("{gst1}",gs1_fee);
  body.replaceText("{s_date}",s_date);
  body.replaceText("{e_date}",e_date);
  body.replaceText("{ad_test}",adm_date);
  temp_doc_file.saveAndClose();

  var PDF_content = temp_File.getAs(MimeType.PDF);
  var PDF_File = PDF_Foldr.createFile(PDF_content).setName("Admission Report " + nam + " " + rol);
  Tmp_Folder.removeFile(temp_File);

 FUNCTION FOR MAILING UNCOMMENT WHENEVER WEBAPP IS READY*/
   MailApp.sendEmail(mai, 'Admission Report', "The admission report for student " + nam + " can be found as follows:",   {               
     attachments: [PDF_File.getAs(MimeType.PDF)],
     name : 'Progressive minds automated emailer'
   });

  Logger.log(PDF_File)
 }


  function getdata4 () {
    var fee_data = getDatarange4(12, 2, 4, 3);
    var disc_data = getDatarange4(18, 2, 6, 3);
     div_data4(fee_data, disc_data);
     //Logger.log(fee_data);
     // Logger.log(disc_data);
  }
