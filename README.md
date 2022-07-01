# Google-Apps-Script-Attendance-Staff
Google apps script and Html files for developmentent of PDF automation softwares, Staff management software

# Attendance
# 1. Inputs
Attendance PDF automation and mass emailing script takes a google sheet file as an input which has the following template : https://docs.google.com/spreadsheets/d/1hnFpYsxkFcuvRNZ0WQU3yI83I-nnSF_r210mYlJ0Nok/edit?usp=sharing
The google app script runs on this template so in case any changes are made in the google sheets template similar changes have to be made in the google apps script     file to extract data from the correct cells.
Google docs link for PDF template : 
https://docs.google.com/document/d/1OaJwbb9l3ID9j9K2yGzrcTWGlTbtNCnwAEtltyPXIXo/edit?usp=sharing
Variables from  google apps script can be directly imported into the google docs file using but on changing the variable names in the google docs file similar changes have to made in the code for correct substitution of the data.
Please note that the sheet takes input using manual entry directly through the excel. A webapp can be deployed using google apps script to add data in excel through a  form.
  functions involved : 
  function Main() - Parent function running this runs all the functions and starts the process of data extraction and PDF generation and Emailing.
  function getDatarange(firstr,firstc,lastr,lastc)
  
# 2. Data cleaning   
After the data entries have been completed for a month. Then data from the excel sheet is extracted row-wise. This row-wise data is further segregated into various      variables (i.e Subject-wise attendance, mail, student name, subject-wise assignment etc) and the main array holding the row is accordingly split up into these          variables. Google apps script treats non filled cells as empty and these are extracted as empty arrays ([]). So if two cells are merged in google sheets then one cell  will be treated as having data and the other cell in the merged cell will be treated as empty([]). Hence filtering functions are required to remove NULL/empty          elements. These need to filtered so that they are not substituted in the PDF. 
functions involved :
function filterData (data)
function concat(parent_data,data1,data2,data3)
function div_data(st_data)
 
 # 3. Automated PDF creation and emailing
After all the data has been segregated into different variables. To maintain the original doc template, A copy of the origial doc is created in a temporary drive folder. In this temporary google doc file the variables are substituted. After the substitution the temporary doc file is converted into a PDF and is stored in a variable. To conserve space the temporary doc file is deleted and then the PDF is mailed to all the users in the mail array. The PDFs are also stored in a permanent drive folder.
functions involved : 
function create_PDFnMail(tot_p,tot_c,tot_m,name,tot_pasn,tot_casn,tot_masn,roll,mail,att_p,att_c,att_m,late_p,late_c,late_m,pasn,pasi,casn,casi,masn,masi,doc_file,Temp_Folder,PDF_Folder)


# Fee Structure
Fee Structure PDF automation and emailing script takes a google sheet file and a google docs file as an input. All the discounts are calculated within the sheet 
itself, then all changes are reflected in the google docs template which is stored in google drive & also PDF will be generated and then the PDF is mailed to the student's E-mail id provided in the google sheets. 
Different scripts are written for all the different classes like 9th,10th,11th,12th & droppers.
A button is provided in the sheet itself, on clicking the button it will run the corresponding script and mail the PDF to the mentioned student & stores the PDF in google drive folder as well.

# Staff Management
The Staff Management system is a webapp, where staff of the organization can mark their daily attendance. The webapp includes a login dashboard which is protected by a passcode so that only staff can open the webpage.
Staff will type his/her name and click on the ADD ENTRY button to record their entry time and at the time of leaving they can again type their name and click on ADD EXIT to record their exit time.
A DELETE ENTRY button is also provided, to delete a specific entry whereas, the DELETE ALL button will delete all the entries from the spreadsheet.
After the end of every month the data from the google sheet will be converted to PDF and will get stored in the google drive folder and the sheet will become empty again.
Two seperate buttons are also provided in the webapp so that, user can see the google sheet as well as the google drive folder.
