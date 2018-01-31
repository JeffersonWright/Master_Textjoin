function CreateTextjoin() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Sheet1'); //sets the active sheet to the variable "sheet"
  sheet.getRange('E6:P30').merge(); //merges cells E5:P30 to make the output display box
  
  sheet.getRange('E6').setVerticalAlignment("Top"); //sets the vertical alignment for cell E5
  sheet.getRange('A:Z').setHorizontalAlignment("Left"); //sets the horizontal alignment for all relevant cells
  
  sheet.getRange('E1:M2').setBorder(true,true,true,true,true,true,'black',SpreadsheetApp.BorderStyle.SOLID); //sets the borders and color of cells
  sheet.getRange('E6').setBorder(true,true,true,true,true,true,'black',SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('O4').setBorder(true,true,true,true,true,true,'black',SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('O4').setBackgroundColor("#ffff00");
  
  for (i=1; i<6; i++) {  
    sheet.setColumnWidth(i, 110);
  } //sets the width of columns A:E to 110
  for (i=6; i<15; i++) {  
    sheet.setColumnWidth(i, 67);
  } //sets the width of columns A:E to 67
  
  
  
  sheet.getRange('E6').setValue("=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TEXTJOIN($F$3,1,FILTER(IF(LEFT($A:$D,LEN(F5))=F5,$A:$D,),MOD(ROW($A:$D)+$F$4,($F$4+1))=0)),F1,F2),H1,H2),J1,J2),L1,L2)");
  //puts the main textjoin in cell E5; this is the heart of the script

  //The rest of this fills in the rest of the cells, and the sheet is completed and ready for use
  sheet.getRange('E1').setValue("Remove Char:");
  sheet.getRange('E2').setValue("Replace With:");
  sheet.getRange('E3').setValue("Join With:");
  sheet.getRange('E4').setValue("Skip Lines:");
  sheet.getRange('E5').setValue("Starts With:");
  sheet.getRange('F3').setValue(", ");
  sheet.getRange('F4').setValue("0");
  sheet.getRange('H3').setValue("Char Total:");
  sheet.getRange('H4').setValue("Current:");
  sheet.getRange('H5').setValue("Entries:");
  sheet.getRange('K3').setValue("Removed");
  sheet.getRange('K4').setValue("Replaced");
  sheet.getRange('K5').setValue("Shortened");
  sheet.getRange('N3').setValue("     Single-line output, copy then paste values");
  sheet.getRange('O4').setValue("=E6");
  sheet.getRange('R1').setValue("=Proper(E6)");
  sheet.getRange('R2').setValue("=Upper(E6)");
  sheet.getRange('R3').setValue("=Lower(E6)");
  
  for (i=7; i<15; i=i+2) {  
    sheet.getRange(1, i).setValue("Occured");
  } //sets cells G1, I1, K1, and M1 to "Occured"
  sheet.getRange('G2').setValue("=IF(ISBLANK(F1),\"0\",(LEN(textjoin($F$3,1,FILTER(IF(LEFT($A:$D,LEN(F5))=F5,$A:$D,),MOD(ROW($A:$D)+$F$4,($F$4+1))=0)))-LEN(SUBSTITUTE(textjoin($F$3,1,FILTER(IF(LEFT($A:$D,LEN(F5))=F5,$A:$D,),MOD(ROW($A:$D)+$F$4,($F$4+1))=0)),F1,)))/(LEN(F1)))");
  sheet.getRange('I2').setValue("=IF(ISBLANK(H1),\"0\",(LEN(textjoin($F$3,1,FILTER(IF(LEFT($A:$D,LEN(F5))=F5,$A:$D,),MOD(ROW($A:$D)+$F$4,($F$4+1))=0)))-LEN(SUBSTITUTE(textjoin($F$3,1,FILTER(IF(LEFT($A:$D,LEN(F5))=F5,$A:$D,),MOD(ROW($A:$D)+$F$4,($F$4+1))=0)),H1,)))/(LEN(H1)))");
  sheet.getRange('K2').setValue("=IF(ISBLANK(J1),\"0\",(LEN(textjoin($F$3,1,FILTER(IF(LEFT($A:$D,LEN(F5))=F5,$A:$D,),MOD(ROW($A:$D)+$F$4,($F$4+1))=0)))-LEN(SUBSTITUTE(textjoin($F$3,1,FILTER(IF(LEFT($A:$D,LEN(F5))=F5,$A:$D,),MOD(ROW($A:$D)+$F$4,($F$4+1))=0)),J1,)))/(LEN(J1)))");
  sheet.getRange('M2').setValue("=IF(ISBLANK(L1),\"0\",(LEN(textjoin($F$3,1,FILTER(IF(LEFT($A:$D,LEN(F5))=F5,$A:$D,),MOD(ROW($A:$D)+$F$4,($F$4+1))=0)))-LEN(SUBSTITUTE(textjoin($F$3,1,FILTER(IF(LEFT($A:$D,LEN(F5))=F5,$A:$D,),MOD(ROW($A:$D)+$F$4,($F$4+1))=0)),L1,)))/(LEN(L1)))");
  sheet.getRange('I3').setValue("=Len(textjoin($F$3,1,FILTER(IF(LEFT($A:$D,LEN(F5))=F5,$A:$D,),MOD(ROW($A:$D)+$F$4,($F$4+1))=0)))");
  sheet.getRange('I4').setValue("=Len(E6)");
  sheet.getRange('I5').setValue("=COUNTA(FILTER(IF(LEFT($A:$D,LEN(F5))=F5,$A:$D,),MOD(ROW($A:$D)+$F$4,($F$4+1))=0))");
  sheet.getRange('L3').setValue("=SUM(ArrayFormula({G2,I2,K2,M2}*LEN({F1,H1,J1,L1})))");
  sheet.getRange('L4').setValue("=SUM(ArrayFormula({G2,I2,K2,M2}*LEN({F2,H2,J2,L2})))");
  sheet.getRange('L5').setValue("=I3-I4");  
  
}
