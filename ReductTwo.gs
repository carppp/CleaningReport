function reducetwo() {
  var SpreadSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1FVHoWxEc3Ru4Mhe7mhTJKItomeOZVR2EAXKgoYQKhEQ/edit#gid=1416143441");
  var Sheet = SpreadSheet.getSheetByName("統計");
  var iindex1 = Sheet.getRange(7, 3).getValue();
  var oindex = Sheet.getRange(16, 3).getValue();
  var rindex = Sheet.getRange(4, 3).getValue();
  Sheet.getRange(7, 3).setValue(iindex1-2);
  Sheet.getRange(16, 3).setValue(oindex-2);
  Sheet.getRange(4, 3).setValue(rindex-2);
}
