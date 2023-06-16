function nodata() {
  var inputSpreadSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1FVHoWxEc3Ru4Mhe7mhTJKItomeOZVR2EAXKgoYQKhEQ/edit#gid=121839247");
  var inputSheet = inputSpreadSheet.getSheetByName("表單回應 1");
  var outputSheet = inputSpreadSheet.getSheetByName("統計");
  var Today =new Date();
  var date = Today.getFullYear()+"/"+(Today.getMonth()+1)+"/"+Today.getDate();
  var a = inputSheet.getRange(2,6,1000,1).getValues();
  var cindex1 = outputSheet.getRange(7, 3).getValue();
  var oindex = outputSheet.getRange(16, 3).getValue();
  var rindex = outputSheet.getRange(4, 3).getValue();
  var num=0,tmp=0;
  var class=[];
  var classroom = false,outside=false,recycle=false;
  
  for(var i =1;i<1000;i++)
  {
    var data = new Date(a[i-1][0]);
    var compare = data.getFullYear()+"/"+(data.getMonth()+1)+"/"+data.getDate();
    if(compare == date)
    {
      num=num+1;
      class[tmp] = inputSheet.getRange(1+i, 2).getValue();
      tmp++;
    }
  }
  for(var i =0;i<6;i++)
  {
    if(class[i] == "內掃")
      classroom = true;
    if(class[i] == "外掃")
      outside = true;
    if(class[i] == "回收")
      recycle = true;
  }
  if(classroom == false)
  {
    outputSheet.getRange(7, 3).setValue(cindex1+1);
  }
  if(outside == false)
  {
    outputSheet.getRange(16, 3).setValue(oindex+1);
  }
  if(recycle == false)
  {
    outputSheet.getRange(4, 3).setValue(rindex+1);
  }
}
