function formatSheet1() {

//Get this spreadsheet and sheet 1
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var sheet1 = ss.getSheetByName("Sheet1");  

//Get full range of data, last column and last row
  var allCells = sheet1.getDataRange();
  var numColumns = sheet1.getLastColumn();
  var numRows = sheet1.getLastRow();

  //Get certain ranges: header row, comments column
  var headerRow = sheet1.getRange(1, 1, 1, numColumns);
  var commentsColumn = sheet1.getRange(2, numColumns, numRows, 1);
  
//Add borders to all cells with data in them
  allCells.setBorder(true, true, true, true, true, true);    
  
//Change font to calibri and font size to 11
  allCells.setFontFamily("Calibri");
  allCells.setFontSize(11);  

//Change font size in Comments column to 9
  commentsColumn.setFontSize(9);  
  
//Set wrap to Row 1
  headerRow.setWrap(true);
  
//Set wrap to Comments column
  commentsColumn.setWrap(true);
  
//Set horizontal alignment to centre for all cells
  allCells.setHorizontalAlignment("center");

//Set horizontal alignment of Comments column to left
  commentsColumn.setHorizontalAlignment("left");

//Change vertical alignment of data to middle
  allCells.setVerticalAlignment("middle"); 
  
//Change vertical alignment of header row (row 1) to bottom  
  headerRow.setVerticalAlignment("bottom");

//Change Class number format to 00
  var classColumn = sheet1.getRange(2, 2, numRows, 1);
  classColumn.setNumberFormat("00"); 

//Change date formats to dd/mm
  var dateColumns = sheet1.getRange(2, 8, numRows, 11);
  dateColumns.setNumberFormat("dd/mm"); 

//Add bolding to header row
  headerRow.setFontWeight("bold");
 
//Fill Header Row with colours using Hexidecimal codes
  var brown = "#dd7e6b";
  var red = "#dd7e6b";
  var orange = "#f9cb9c";
  var yellow = "#ffe599";
  var green = "#b6d7a8";
  var blue = "#a4c2f4";
  var purple = "#b4a7d6";
  var magenta = "#d5a6bd";
  
  var ag1 = sheet1.getRange("A1:G1");
  ag1.setBackground (blue); 

  var h1 = sheet1.getRange("H1");
  h1.setBackground (green);  

  var i1 = sheet1.getRange("I1");
  i1.setBackground (purple); 
  
  var j1 = sheet1.getRange("J1");
  j1.setBackground (magenta); 
  
  var kl1 = sheet1.getRange("K1:L1");
  kl1.setBackground (red); 
  
  var m1 = sheet1.getRange("M1");
  m1.setBackground (yellow); 
  
  var n1 = sheet1.getRange("N1");
  n1.setBackground (orange); 
  
  var o1 = sheet1.getRange("O1");
  o1.setBackground (purple); 
  
  var pq1 = sheet1.getRange("P1:Q1");
  pq1.setBackground (magenta); 

  var r1 = sheet1.getRange("R1");
  r1.setBackground (red); 

  var s1 = sheet1.getRange("S1");
  s1.setBackground (brown);

//Delete columns T to Z
  var totalColumns = sheet1.getMaxColumns();
  var numOfColumnsToDelete = totalColumns - numColumns;
  sheet1.deleteColumns(numColumns+1, numOfColumnsToDelete);
  
//Delete unused rows (maxRows - last Row)
  var totalRows = sheet1.getMaxRows();
  var numOfRowsToDelete = totalRows - numRows;
  sheet1.deleteRows(numRows+1, numOfRowsToDelete);
  
//Set widths of columns A to R with autoResize
  for (i=1; i<8; i++) {  
  sheet1.autoResizeColumn(i);  
  } 

//Set width of columns with dates to 60
  for (i=8; i<19; i++) {  
  sheet1.setColumnWidth(i, 60); 
  }   
  
//Set width of Comments column to 120
  sheet1.setColumnWidth(numColumns, 120); 
}

function hideRowsColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName("Sheet1");

//hide columns L to R
  var columnsToHide = sheet1.getRange("L1:R1");
  sheet1.hideColumn(columnsToHide); 

//hide french classes (rows 4 & 10)
  var row4 = sheet1.getRange("A4");
  var row10 = sheet1.getRange("A10"); 
  sheet1.hideRow(row4);
  sheet1.hideRow(row10);
}

function unhideRowsColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName("Sheet1");

//unhide all the columns
  var numColumns = sheet1.getLastColumn();
  var allColumns = sheet1.getRange(1, 1, 1, numColumns);
  sheet1.unhideColumn(allColumns);

//unhide all the rows  
  var numRows = sheet1.getLastRow();
  var allRows = sheet1.getRange(1, 1, numRows, 1);
  sheet1.unhideRow(allRows);    
}

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