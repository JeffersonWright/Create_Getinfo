function CreateTextjoin() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Sheet1'); //sets the active sheet to the variable "sheet"
  sheet.getRange('E5:P30').merge(); //merges cells E5:P30 to make the output display box
  
  sheet.getRange('E5').setVerticalAlignment("Top"); //sets the vertical alignment for cell E5
  sheet.getRange('A:Z').setHorizontalAlignment("Left"); //sets the horizontal alignment for all relevant cells
  
  sheet.getRange('E1:M2').setBorder(true,true,true,true,true,true,'black',SpreadsheetApp.BorderStyle.SOLID); //sets the borders and color of cells
  sheet.getRange('E5').setBorder(true,true,true,true,true,true,'black',SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('P3').setBorder(true,true,true,true,true,true,'black',SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('P3').setBackgroundColor("#ffff00");
  
  for (i=1; i<6; i++) {  
    sheet.setColumnWidth(i, 110);
  } //sets the width of columns A:E to 110
  for (i=6; i<15; i++) {  
    sheet.setColumnWidth(i, 67);
  } //sets the width of columns A:E to 67
  
  
  
  sheet.getRange('E5').setValue("=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(textjoin($F$3,1,FILTER(IF(LEFT($A:$D,LEN(N3))=N3,$A:$D,),MOD(ROW($A:$D)+$F$4,($F$4+1))=0)),F1,F2),H1,H2),J1,J2),L1,L2)");
  //puts the main textjoin in cell E5; this is the heart of the script

  //The rest of this fills in the rest of the cells, and the sheet is completed and ready for use
  sheet.getRange('E1').setValue("Remove Char:");
  sheet.getRange('E2').setValue("Replace With:");
  sheet.getRange('E3').setValue("Join With:");
  sheet.getRange('E4').setValue("Skip Lines:");
  sheet.getRange('F3').setValue(", ");
  sheet.getRange('F4').setValue("0");
  sheet.getRange('G3').setValue("Char Total:");
  sheet.getRange('G4').setValue("Current:");
  sheet.getRange('J3').setValue("Removed");
  sheet.getRange('J4').setValue("Replaced");
  sheet.getRange('M3').setValue("Starts with");
  sheet.getRange('M4').setValue("Shortened");
  sheet.getRange('N2').setValue("     Single-line output, copy then paste values");
  sheet.getRange('P3').setValue("=E5");
  sheet.getRange('R1').setValue("=Proper(E5)");
  sheet.getRange('R2').setValue("=Upper(E5)");
  sheet.getRange('R3').setValue("=Lower(E5)");
  
  for (i=7; i<15; i=i+2) {  
    sheet.getRange(1, i).setValue("Occured");
  } //sets cells G1, I1, K1, and M1 to "Occured"
    sheet.getRange('J3').setValue("Removed");
  sheet.getRange('G2').setValue("=IF(ISBLANK(F1),\"0\",(LEN(textjoin($F$3,1,FILTER(IF(LEFT($A:$D,LEN(N3))=N3,$A:$D,),MOD(ROW($A:$D)+$F$4,($F$4+1))=0)))-LEN(SUBSTITUTE(textjoin($F$3,1,FILTER(IF(LEFT($A:$D,LEN(N3))=N3,$A:$D,),MOD(ROW($A:$D)+$F$4,($F$4+1))=0)),F1,)))/(LEN(F1)))");
  sheet.getRange('I2').setValue("=IF(ISBLANK(H1),\"0\",(LEN(textjoin($F$3,1,FILTER(IF(LEFT($A:$D,LEN(N3))=N3,$A:$D,),MOD(ROW($A:$D)+$F$4,($F$4+1))=0)))-LEN(SUBSTITUTE(textjoin($F$3,1,FILTER(IF(LEFT($A:$D,LEN(N3))=N3,$A:$D,),MOD(ROW($A:$D)+$F$4,($F$4+1))=0)),H1,)))/(LEN(H1)))");
  sheet.getRange('K2').setValue("=IF(ISBLANK(J1),\"0\",(LEN(textjoin($F$3,1,FILTER(IF(LEFT($A:$D,LEN(N3))=N3,$A:$D,),MOD(ROW($A:$D)+$F$4,($F$4+1))=0)))-LEN(SUBSTITUTE(textjoin($F$3,1,FILTER(IF(LEFT($A:$D,LEN(N3))=N3,$A:$D,),MOD(ROW($A:$D)+$F$4,($F$4+1))=0)),J1,)))/(LEN(J1)))");
  sheet.getRange('M2').setValue("=IF(ISBLANK(L1),\"0\",(LEN(textjoin($F$3,1,FILTER(IF(LEFT($A:$D,LEN(N3))=N3,$A:$D,),MOD(ROW($A:$D)+$F$4,($F$4+1))=0)))-LEN(SUBSTITUTE(textjoin($F$3,1,FILTER(IF(LEFT($A:$D,LEN(N3))=N3,$A:$D,),MOD(ROW($A:$D)+$F$4,($F$4+1))=0)),L1,)))/(LEN(L1)))");
  sheet.getRange('H3').setValue("=Len(textjoin($F$3,1,FILTER(IF(LEFT($A:$D,LEN(N3))=N3,$A:$D,),MOD(ROW($A:$D)+$F$4,($F$4+1))=0)))");
  sheet.getRange('H4').setValue("=Len(E5)");
  sheet.getRange('K3').setValue("=SUM(ArrayFormula({G2,I2,K2,M2}*LEN({F1,H1,J1,L1})))");
  sheet.getRange('K4').setValue("=SUM(ArrayFormula({G2,I2,K2,M2}*LEN({F2,H2,J2,L2})))");
  sheet.getRange('N4').setValue("=H3-H4");  
  
}
