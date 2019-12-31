function syncAllDataFromPROD(){
   var devEnv=getEnv('dev');
   var prodEnv=getEnv('prod');

   var devSheet=SpreadsheetApp.openById(devEnv.schedulerSpreadsheetId); 
   var prodSheet=SpreadsheetApp.openById(prodEnv.schedulerSpreadsheetId); 
   
   var sheetList=['members','role-definition', 'history','absent','2019Template','registry'];
   
   for (i=0;i<sheetList.length;i++){
      if(devSheet.getSheetByName(sheetList[i])==null)
         devSheet.insertSheet(sheetList[i]);
      SyncSheet(prodSheet.getSheetByName(sheetList[i]), devSheet.getSheetByName(sheetList[i]));
   } 
}


function simulate(){
  //syncAllDataFromPROD();
 
  trigger_checkDaily('dev');
   


}




