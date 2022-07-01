//**************************** DO NOT MODIFY CELL LOCATIONS IN EXCEL AS THE CODE WILL STOP WORKING CORRECTLY***********************/

// 1. DEFINING GLOBAL VARIABLES

// 1.1 Google doc link for pdf template of Fee Structure report: https://docs.google.com/document/d/1c7inSR1iasiyPaGs6AmqIBVV1pVoMn_myR_nPthr7wU/edit
var docd_file = DriveApp.getFileById("1c7inSR1iasiyPaGs6AmqIBVV1pVoMn_myR_nPthr7wU"); 
// 1.2 Temporary folder which holds the above doc for each student this doc is copied, edited and deleted for each student. This is to make sure that the original doc is not changed as then each report will be modified with the earlier changes. After the doc is converted to pdf and mailed the temporary doc is generated again for each report: https://drive.google.com/drive/u/0/folders/1MtuLu3NCbJYh5C5RsWpjDi1KYnxHv-Lf
  var Tempd_Folder =  DriveApp.getFolderById("1MtuLu3NCbJYh5C5RsWpjDi1KYnxHv-Lf");
// 1.3 This folder is responsible for holding the reports of all the students. :  https://drive.google.com/drive/u/0/folders/1wQBF6bMangBnxFvgbJRI6ZKdgm-jitkv
  var PDFd_Folder =DriveApp.getFolderById("1wQBF6bMangBnxFvgbJRI6ZKdgm-jitkv");
//1.4 google sheet holding attendance,and assignment submission of all the students we get the sheet by name and the name of the sheet is "Attendance" : https://docs.google.com/spreadsheets/d/1M6wvyTyRvZMWFkGV-TTRf9bjkyG53iE8QDPNDdtU894/edit?usp=sharing 
  var sheetd = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dropper");

// upto 2 decimal value 
function typecastd(numb) {
  numb = numb.toFixed(1);
  return numb;
}
// Global variables
  var named = sheetd.getRange("B8").getValue();
  var rolld = sheetd.getRange("B9").getValue();
  var maild = sheetd.getRange("B10").getValue();
  var totald_fees = typecastd(sheetd.getRange("D16").getValue()); 
  var totald_dis = typecastd(sheetd.getRange("D24").getValue());
  var grossd_fee = typecastd(sheetd.getRange("D25").getValue());
  var gstd_fee = typecastd(sheetd.getRange("D26").getValue());
  var netd_fee = typecastd(sheetd.getRange("D27").getValue());
  var instd_1 = typecastd(sheetd.getRange("C31").getValue());
  var instd_2 = typecastd(sheetd.getRange("C32").getValue());
  var currd_date = sheetd.getRange("B31").getValue();
  var in2d_date = sheetd.getRange("B32").getValue();
  var in3d_date = sheetd.getRange("B33").getValue();
 // var in4_date = sheet.getRange("B34").getValue();
  var std_date = sheetd.getRange("E8").getValue();
  var edd_date = sheetd.getRange("F8").getValue();

//2.1 This function takes the data range in the excel as its input and the data for attendance, and assignments are segregated
function getDataranged(firstr,firstc,lastr,lastc) 
{
  return sheetd.getRange(firstr,firstc,lastr,lastc).getValues();
}

function div_datad(fee_data, disc_data){
   var ad_fed = typecastd(fee_data[0][1]);
   //Logger.log(ad_fe);
   var ad_perd = fee_data[0][0]*100;
  //  Logger.log(ad_per);
   var tutd_fe = typecastd(fee_data[1][1]);
   var tutd_per = fee_data[1][0]*100;
   var bookd_fe = typecastd(fee_data[2][1]);
   var bookd_per = fee_data[2][0]*100;
   var infd_fe = typecastd(fee_data[3][1]);
   var infd_per = fee_data[3][0]*100;
   var tchd_fe = typecastd(disc_data[0][1]);
   var tchd_per = disc_data[0][0]*100;
   var rnkd_fe = typecastd(disc_data[1][1]);
   var rnkd_per = disc_data[1][0]*100;
   var pard_fe = typecastd(disc_data[2][1]);
   var pard_per = disc_data[2][0]*100;
   var refd_fe = typecastd(disc_data[3][1]);
   var refd_per = disc_data[3][0]*100;
   var lnsd_fe = typecastd(disc_data[4][1]);
   var lnsd_per = disc_data[4][0]*100;
   var schd_fe = typecastd(disc_data[5][1]);
   var schd_per = disc_data[5][0]*100;

   create_PDFnMaild(ad_fed,ad_perd,tutd_fe,tutd_per,bookd_fe,bookd_per,infd_fe,infd_per,tchd_fe,tchd_per,rnkd_fe,rnkd_per,pard_fe,pard_per,refd_fe,refd_per,lnsd_fe,lnsd_per,schd_fe,schd_per,docd_file,Tempd_Folder,PDFd_Folder)
}

// CREATE AND MAIL FUNCTION

function create_PDFnMaild(ad_fed,ad_perd,tutd_fe,tutd_per,bookd_fe,bookd_per,infd_fe,infd_per,tchd_fe,tchd_per,rnkd_fe,rnkd_per,pard_fe,pard_per,refd_fe,refd_per,lnsd_fe,lnsd_per,schd_fe,schd_per,docd_file,Tempd_Folder,PDFd_Folder)
{
  var temp_File = docd_file.makeCopy(Tempd_Folder);
  var temp_doc_file = DocumentApp.openById(temp_File.getId());
  var body = temp_doc_file.getBody();
  body.replaceText("{st_name}",named);
  body.replaceText("{rl_num}",rolld);
  body.replaceText("{tot_fee}",totald_fees);
  body.replaceText("{tot_dis}",totald_dis);
  body.replaceText("{gross_fee}",grossd_fee);
  body.replaceText("{ad_fee}",ad_fed);
  body.replaceText("{ad_per}",ad_perd);
  body.replaceText("{tut_fee}",tutd_fe);
  body.replaceText("{tut_per}",tutd_per);
  body.replaceText("{book_fee}",bookd_fe);
  body.replaceText("{mod_per}",bookd_per);
  body.replaceText("{infr_fee}",infd_fe);
  body.replaceText("{tech_per}",infd_per);
  body.replaceText("{tec_dis}",tchd_fe);
  body.replaceText("{lec_per}",tchd_per);
  body.replaceText("{rnk_dis}",rnkd_fe);
  body.replaceText("{rnk_per}",rnkd_per);
  body.replaceText("{oth_dis}",pard_fe);  
  body.replaceText("{oth_per}",pard_per);
  body.replaceText("{ref_dis}",refd_fe);
  body.replaceText("{ref_per}",refd_per);
  body.replaceText("{ls_fee}",lnsd_fe);
  body.replaceText("{lns_per}",lnsd_per);
  body.replaceText("{sch_dis}",schd_fe);
  body.replaceText("{sch_per}",schd_per);
  body.replaceText("{inst_1}",instd_1);
  body.replaceText("{inst_2}",instd_2);
  body.replaceText("{inst_3}",instd_2);
  body.replaceText("{inst_4}",instd_2);
  body.replaceText("{net_fee}",netd_fee);
  body.replaceText("{curr_date}",currd_date);
  body.replaceText("{inst_2_date}",in2d_date);
  body.replaceText("{inst_3_date}",in3d_date);
 // body.replaceText("{inst_4_date}",in4_date);
  body.replaceText("{gst}",gstd_fee);
  body.replaceText("{st_date}",std_date);
  body.replaceText("{ed_date}",edd_date);
  temp_doc_file.saveAndClose();

  var PDF_content = temp_File.getAs(MimeType.PDF);
  var PDF_File = PDFd_Folder.createFile(PDF_content).setName("Admission Report " + named + " " + rolld);
  Tempd_Folder.removeFile(temp_File);

/* FUNCTION FOR MAILING UNCOMMENT WHENEVER WEBAPP IS READY*/
  // MailApp.sendEmail(maild, 'Admission Report', "The admission report for student " + named + " can be found as follows:",   {               
  //   attachments: [PDF_File.getAs(MimeType.PDF)],
  //   name : 'Progressive minds automated emailer'
  // });

  Logger.log(PDF_File)
 }


  function getdatad () {
    var fee_data = getDataranged(12, 3, 4, 2);
    var disc_data = getDataranged(18, 3, 6, 2);
    div_datad(fee_data, disc_data);
   // Logger.log(fee_data);
    //Logger.log(disc_data);
  }
