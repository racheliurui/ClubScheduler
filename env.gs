function getEnv(envCode){

  var env={};
  env.envCode=envCode;
  
  if(envCode=='dev'){
      env.clubName='aaaaa';
      env.robotemail='aaa.vpe@gmail.com';
      env.schedulerSpreadsheetId='aaaaaaaaaaaaaa';
      env.weeklyAgendaFolderId=''aaaaaaaaaaaaaa';
      env.rootFolderId='aaaaaaaaaaaaaa';
      env.debugEmail='aaa@gmail.com>';
  }
  
  
  if(envCode=='prod'){
      env.clubName='aaaaa';
      env.robotemail='aaa.vpe@gmail.com';
      env.schedulerSpreadsheetId='aaaaaaaaaaaaaa';
      env.weeklyAgendaFolderId=''aaaaaaaaaaaaaa';
      env.rootFolderId='aaaaaaaaaaaaaa';

  }

  var ss= SpreadsheetApp.openById(env.schedulerSpreadsheetId); 
  var config=getConfigFromSheet(ss.getSheetByName("config"));
  
  env.weeklyAgendaTemplateFileId = getFileUnderFolderByName(config['WeeklyAgendaTemplateFileName'],env.rootFolderId);
  env.config=config;


  return env;

}