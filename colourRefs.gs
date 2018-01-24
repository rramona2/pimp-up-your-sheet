function getHexValue() {
  var colourSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ColourRefs");
  for (c=1;c<22;c+=2){
    for (r=1;r<11;r++){
      if(c>1 && r>7) {
        break;}
    var hexCode = colourSheet.getRange(r,c).getBackground();
    colourSheet.getRange(r,c+1).setValue(hexCode);
 }
}
}