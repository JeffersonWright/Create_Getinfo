function CreateGetinfo() {
//  SpreadsheetApp.getActiveSpreadsheet().renameActiveSheet('Getinfo');
  var GIF = SpreadsheetApp.getActive().getSheetByName('Getinfo'); //sets the active sheet to the variable "GIF"
//  SpreadsheetApp.getActiveSpreadsheet().insertSheet('Formula'); //inserts the Formula sheet
  var FUA = SpreadsheetApp.getActive().getSheetByName('Formula'); //sets the active sheet to the variable "FUA"
  
  GIF.setColumnWidth(1, 200); //sets column widths
  GIF.setColumnWidth(2, 200);
  GIF.setRowHeight(1,85); //sets row 1 height
  GIF.setFrozenColumns(1); //freezes the first column
  GIF.getRange('A1:BG1').setFontWeight("bold"); //bolds the first row of text
  for (i=1;i<60;i++){
    GIF.getRange(1 , i+1).setWrap(true);
  } //sets the wrapping for row 1

 
  //Getinfo sheet populate-------------------------------
  var Headers = ["School Name","School URL","Phone","City","State","Gender","Religious Affiliation","Year Founded","Campus Size (km2)","Campus Size (Acres)","Type of Campus","Grade Levels Offered","Total Enrollment","Acceptance Rate (%)","Average Class Size","Student:Teacher Ratio (x:1)","Boarding Students Percentage (%)","Intl' Students Percentage (%)","Faculty with Advanced Degrees(%)","Average SAT Score","Average ACT Score","Number of AP Courses Offered","Number of IB Courses Offered","ESL Courses Offered","Summer Programs Offered","School Address","School State","Nearby Universities","Nearby Major Cities & Distance","Nearest Major International Airport","Climate","Local Area Information","Brief Intro","School Highlights","Accreditation & Memberships","Notable Alumni","Foreign Languages Offered","Art Courses","Honor Courses","AP Courses","IB Courses","Boys Fall Athletics","Boys Winter Athletics","Boys Spring Athletics","Girls Fall Athletics","Girls Winter Athletics","Girls Spring Athletics","Extracurricular Activities","Annual International Student Tuition and Fees (USD)","Average Financial Aid Grant","Are scholarships offered for intl students? ","If scholarships are offered, what range is available? ","Entry Terms","Essay/Personal Statement","Interview","Letters of Recommendation","Application Deadlines","Application Fee (USD)","College Matriculations"];
  var i = 0
  while ((59-i)>0) {  
    GIF.getRange(1 , i+1).setValue(Headers[i]);
    i++;
  } //sets the headers in row 1

  var rule = SpreadsheetApp.newDataValidation().requireValueInList(['Annie Wright Schools','Asheville School','Avon Old Farms School','Ben Lippen School','Berkshire School','Blair Academy','Salisbury School','San Domenico School','Santa Catalina School','Ross School','Portsmouth Abbey School','South Kent School','Squaw Valley Academy'], true).build();
  GIF.getRange('A2').setDataValidation(rule);
  
  var Info = ['Salisbury School','=IMPORTXML(CONCATENATE(\"https://www.boardingschoolreview.com/\",TEXTJOIN(\"-\",1,SUBSTITUTE($A$2,\" \",\"-\"),\"profile\")),\"//div\[@class=\'top_card_ctn top_website_ctn\'\]/a/@href\")','=VLOOKUP(\"Telephone\", Formula!G:H, 2, FALSE)','=IMPORTXML(CONCATENATE(\"https://www.boardingschoolreview.com/\",TEXTJOIN(\"-\",1,SUBSTITUTE($A$2,\" \",\"-\"),\"profile\")),\"//span[@itemprop=\'addressLocality\'\]\")','=IMPORTXML(CONCATENATE(\"https://www.boardingschoolreview.com/\",TEXTJOIN(\"-\",1,SUBSTITUTE($A$2,\" \",\"-\"),\"profile\")),\"//span[@itemprop=\'addressRegion\'\]\")','=VLOOKUP(\"School Type\", Formula!A:B, 2, FALSE)','=VLOOKUP(\"Religious Affiliation\", Formula!A:B, 2, FALSE)','=VLOOKUP(\"Year Founded\", Formula!A:B, 2, FALSE)','=J2*0.0040468564224','=SUBSTITUTE(VLOOKUP(\"Campus Size\", Formula!A:B, 2, FALSE),\" acres\",)','','=SUBSTITUTE(IF(ISNA(VLOOKUP(\"Grades Offered\", Formula!A:B, 2, FALSE)),VLOOKUP(\"Grades Offered (Boarding)\", Formula!A:B, 2, FALSE),(VLOOKUP(\"Grades Offered\", Formula!A:B, 2, FALSE))),\",\",)','=SUBSTITUTE(VLOOKUP(\"Enrollment\", Formula!A:B, 2, FALSE),\" students\",)','=IFERROR(SUBSTITUTE(VLOOKUP(\"Acceptance Rate\", Formula!A:B, 2, FALSE),\"%\",),)','=SUBSTITUTE(VLOOKUP(\"Average Class Size\", Formula!A:B, 2, FALSE),\" students\",)','=SUBSTITUTE(VLOOKUP(\"Teacher : Student Ratio\", Formula!A:B, 2, FALSE),\"1:\",)','=SUBSTITUTE(VLOOKUP(\"% Students Boarding\", Formula!A:B, 2, FALSE),\"%\",)','=SUBSTITUTE(VLOOKUP(\"% International Students\", Formula!A:B, 2, FALSE),\"%\",)','=SUBSTITUTE(VLOOKUP(\"% Faculty with Advanced Degree\", Formula!A:B, 2, FALSE),\"%\",)','=IFERROR(IF(VLOOKUP(\"Average SAT Score\", Formula!A:B, 2, FALSE)>1550,ROUND((VLOOKUP(\"Average SAT Score\", Formula!A:B, 2, FALSE)*2/3),0),VLOOKUP(\"Average SAT Score\", Formula!A:B, 2, FALSE)),\"\")','=IFERROR(VLOOKUP(\"Average ACT Score\", Formula!A:B, 2, FALSE),\"\")','','','=VLOOKUP(\"ESL Courses Offered\", Formula!A:B, 2, FALSE)','=VLOOKUP(\"Summer Program Offered\", Formula!A:B, 2, FALSE)','=TEXTJOIN(\" \",1,IMPORTXML(CONCATENATE(\"https://www.boardingschoolreview.com/\",TEXTJOIN(\"-\",1,SUBSTITUTE($A$2,\" \",\"-\"),\"profile\")),\"//span[@itemprop=\'address\']\"))','=E2','=TEXTJOIN(\", \",1,Formula!J1:J3)','','','=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(IMPORTXML(CONCATENATE(\"http://www.bestplaces.net/climate/city/\",TEXTJOIN(\"/\",1,VLOOKUP(E2,Formula!D:E,2,FALSE),D2)),\"//*\[@id=\'dgHeader\'\]/tbody/tr\[2\]/td\[2\]/p\[2\]\"),char(10),),\", where a higher score indicates a more comfortable year-around climate. The US average for the comfort index is 54. Our index is based on the total number of days annually within the comfort range of 70-80 degrees, and we also applied a penalty for days of excessive humidity.\",\".\"),\"Sperling\'s\",\"The\")','=IFERROR(REGEXREPLACE((IMPORTXML(CONCATENATE(\"https://www.bestplaces.net/city/\",LOWER(SUBSTITUTE(TEXTJOIN(\"/\",1,VLOOKUP(E2,Formula!D:E,2,FALSE),D2),\" \",\"\"))),\"//div\[@class=\'side-padded\'\]/div\[@class=\'row\'\]/div\[@class=\'12u\'\]/p\")),\"\n\",\"\"),\"\")','=IFERROR(REGEXREPLACE(TEXTJOIN(\" \",1,(IMPORTXML(CONCATENATE(\"https://www.niche.com/k12/\",TEXTJOIN(\"-\",1,SUBSTITUTE(Getinfo!$A$2,\" \",\"-\"),SUBSTITUTE(Getinfo!$D$2,\" \",\"-\"),Getinfo!E2)),\"//p\[@class=\'premium-paragraph__text\'\]"))),\"\n\",\"\"),)','','=IFERROR(VLOOKUP(\"Organization Memberships\", Formula!G:H, 2, FALSE),)','','','','','','','=IF(F2=\"All-girls\",\"N/A\",HLOOKUP(\"Sports\", Formula!J1:N98, 2, FALSE))','=AP2','=AP2','=IF(F2=\"All-boys\",\"N/A\",HLOOKUP(\"Sports\", Formula!J1:N98, 2, FALSE))','=AS2','=AS2','=IFERROR(HLOOKUP(\"Extracurriculars\", Formula!J1:P98, 2, FALSE),)','=TEXTJOIN(\" \",1,\"Boarding:\",VLOOKUP(\"Yearly Tuition\", Formula!A:B, 2, FALSE))','=IFERROR(VLOOKUP(\"Avg. Financial Aid Grant\", Formula!A:B, 2, FALSE),)','=IFERROR(VLOOKUP(\"Merit Scholarships Offered\", Formula!A:B, 2, FALSE),)','=IF(AY2=\"No\",\"N/A\",)','','','=IFERROR(SUBSTITUTE(VLOOKUP(\"Interview Required\", Formula!G:H, 2, FALSE),\"Yes\",\"Required\"),\"\")','','=IFERROR(SUBSTITUTE(VLOOKUP(\"Application Deadline\", Formula!A:B, 2, FALSE),\"None /\",\"None / rolling\"),)','=IFERROR(SUBSTITUTE(VLOOKUP(\"Application Fee\", Formula!G:H, 2, FALSE),\"$\",),)','=IFERROR(TEXTJOIN(\", \",1,IMPORTXML(CONCATENATE(\"https://www.boardingschoolreview.com/\",TEXTJOIN(\"-\",1,SUBSTITUTE(Getinfo!$A$2,\" \",\"-\"),"profile\")),\"//*/td\[@id=\'matriculation-name\'\]")),)'];
  var i = 0
  while ((59-i)>0) {  
    GIF.getRange(2 , i+1).setValue(Info[i]);
    i++;
  } //sets the info in row 2


  //Formla sheet populate-----------------------------------
  FUA.getRange(1 , 1).setValue('=TRANSPOSE(SPLIT(TEXTJOIN(\"; ;\",1,TRANSPOSE(IMPORTXML(CONCATENATE(\"https://www.boardingschoolreview.com/\",TEXTJOIN(\"-\",1,SUBSTITUTE(Getinfo!$A$2,\" \",\"-\"),\"profile\")),\"//*/tr[@class=\'three_columns\']/td[@class=\'property-name\']\"))),\";\"))');
  FUA.getRange(1 , 2).setValue('=TRANSPOSE(SPLIT(TEXTJOIN(\";\",1,TRANSPOSE(IMPORTXML(CONCATENATE(\"https://www.boardingschoolreview.com/\",TEXTJOIN(\"-\",1,SUBSTITUTE(Getinfo!$A$2,\" \",\"-\"),\"profile\")),\"//*/tr[@class=\'three_columns\']/td[@class=\'property-value\']\"))),\";\"))');
  FUA.getRange(1 , 7).setValue('=TRANSPOSE(SPLIT(TEXTJOIN(\";\",1,TRANSPOSE(IMPORTXML(CONCATENATE(\"https://www.niche.com/k12/\",TEXTJOIN(\"-\",1,SUBSTITUTE(Getinfo!$A$2,\" \",\"-\"),SUBSTITUTE(Getinfo!$D$2,\" \",\"-\"),Getinfo!E2)),\"//div[@class=\'scalar__label\']\"))),\";\"))');  
  FUA.getRange(1 , 8).setValue('=TRANSPOSE(SPLIT(TEXTJOIN(\";\",1,TRANSPOSE(IMPORTXML(CONCATENATE(\"https://www.niche.com/k12/\",TEXTJOIN(\"-\",1,SUBSTITUTE(Getinfo!$A$2,\" \",\"-\"),SUBSTITUTE(Getinfo!$D$2,\" \",\"-\"),Getinfo!E2)),\"//div[@class=\'scalar__value\']\"))),\";\"))');
  FUA.getRange(1 , 10).setValue('=IMPORTXML(CONCATENATE(\"https://www.niche.com/k12/\",TEXTJOIN(\"-\",1,SUBSTITUTE(Getinfo!$A$2,\" \",\"-\"),SUBSTITUTE(Getinfo!$D$2,\" \",\"-\"),Getinfo!E2)),\"//h6[@class=\'popular-entity__name\']\")');
  FUA.getRange(1 , 11).setValue('=TRANSPOSE(IMPORTXML(CONCATENATE(\"https://www.boardingschoolreview.com/\",TEXTJOIN(\"-\",1,SUBSTITUTE(Getinfo!$A$2,\" \",\"-\"),\"profile\")),\"//*/td[@class=\'table_name_cell\']\"))');   
  FUA.getRange(3 , 11).setValue('=TRANSPOSE(IMPORTXML(CONCATENATE(\"https://www.boardingschoolreview.com/\",TEXTJOIN(\"-\",1,SUBSTITUTE(Getinfo!$A$2,\" \",\"-\"),\"profile\")),\"//*/td[@class=\'table_value_cell value_cell_1\']\"))'); 
  //populates the IMPORTXML Functions
  
  var KTQ = ['K3:K','L3:L','M3:M','N3:N','O3:O','P3:P','Q3:Q']; 
  for (i=0;i<7;i++){
    var catcher = '=REGEXREPLACE(TEXTJOIN(\"\",1,('+KTQ[i]+')),\"\\n\",\"\")';
    FUA.getRange(2 , i+11).setValue(catcher);
  } //populates seven columns of textjoin functions to catch columns of text

  var StateCode = ["AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN","IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV","NH","NJ","NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN","TX","UT","VT","VA","WA","WV","WI","WY","DC"];
  var StateName = ["Alabama","Alaska","Arizona","Arkansas","California","Colorado","Connecticut","Delaware","Florida","Georgia","Hawaii","Idaho","Illinois","Indiana","Iowa","Kansas","Kentucky","Louisiana","Maine","Maryland","Massachusetts","Michigan","Minnesota","Mississippi","Missouri","Montana","Nebraska","Nevada","New Hampshire","New Jersey","New Mexico","New York","North Carolina","North Dakota","Ohio","Oklahoma","Oregon","Pennsylvania","Rhode Island","South Carolina","South Dakota","Tennessee","Texas","Utah","Vermont","Virginia","Washington","West Virginia","Wisconsin","Wyoming","District of Columbia"];
  for (i=0;i<51;i++){
    FUA.getRange(1+i , 4).setValue(StateCode[i]);
    FUA.getRange(1+i , 5).setValue(StateName[i]);
  } //populates a column of state codes and names for a vlookup
    
}
