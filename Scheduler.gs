/**
 * @OnlyCurrentDoc
*/

/***
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
Lib Date and extension
***/

function isSameDate(date1, date2){
  
  if(date1.getFullYear()==date2.getFullYear() && date1.getMonth()==date2.getMonth() && date1.getDate()==date2.getDate())
     return true;
     
  else 
     return false;
}


function getMeetingDateBySheetName(sheetName){
  var dateCharactors = sheetName.split("-");
  if(dateCharactors.length!=3)
    {
       SpreadsheetApp.getUi().alert("not a qualified meeting sheet with name " + sheetName);
       return null;
    }
  var year = dateCharactors[0];
  var month =dateCharactors[1]-1;
  var day = dateCharactors[2];
  
  var  someday = new Date();
  someday.setFullYear(dateCharactors[0], (dateCharactors[1]-1), dateCharactors[2]);
 
  return someday;
}

/**
Give a DateString return date object
**/
function getDateByDateString(DateString){
  var dateCharactors = sheetName.split("-");
  var year = dateCharactors[0];
  var month =dateCharactors[1]-1;
  var day = dateCharactors[2];
  
  var  someday = new Date();
  someday.setFullYear(dateCharactors[0], (dateCharactors[1]-1), dateCharactors[2]);
 
  return someday;
}

function getMeetingSheetNameByDate(date){
  
  var year = date.getFullYear();
  var month =date.getMonth()+1;
  var day = date.getDate();

   return year+'-'+ month+'-'+day;
}


function getFormattedDate(date){
var monthNames = [
  "Jan", "Feb", "Mar",
  "Apr", "May", "Jun", "Jul",
  "Aug", "Sep", "Oct",
  "Nov", "Dec"
];
 return date.getDate()+ " " +monthNames[date.getMonth()]+" " +date.getFullYear();


}


function dateDiff(dt1, dt2){

// get milliseconds
var t1 = dt1.getTime();
var t2 = dt2.getTime();
return parseInt((t1-t2)/(24*3600*1000));


}


function getWeekNum(thedate){
    var onejan = new Date(thedate.getFullYear(),0,1);
    return Math.ceil((((thedate - onejan) / 86400000) + onejan.getDay()+1)/7);

}


function getPreviousMeetings(meetingDate,meetingGapDays,numOfMeetings){
     var dateList=new Array();
     for (var i=0;i<numOfMeetings;i++){
         var newDate=new Date();
         newDate.setDate(meetingDate.getDate()-meetingGapDays*(i+1));
         dateList.push(newDate);
       }
  return dateList;
}

function getNextMeetingDates(FirstMeetingDate, meetingGapDays, numOfForcasts){
  var today=new Date();
  var nextMeeingDate;
  
  if(typeof FirstMeetingDate =='string')
   nextMeeingDate=Date.parse(FirstMeetingDate);
  else
   nextMeeingDate=FirstMeetingDate;
  while(nextMeeingDate<today){
    nextMeeingDate=new Date(+nextMeeingDate+meetingGapDays*3600000*24);
  }  
  var dateList=new Array();
  dateList.push(nextMeeingDate);
  for (var i=0;i<numOfForcasts;i++){
         var newDate=new Date();
         newDate.setDate(nextMeeingDate.getDate()+ meetingGapDays*(i+1));
         dateList.push(newDate);
       }
  return dateList;

}

function getPreviousMeetingDate(FirstMeetingDate, MeetingGapDays){
  var today=new Date();
  var nextMeeingDate=FirstMeetingDate;
  var previousMeetingDate=FirstMeetingDate;
  while(nextMeeingDate<today){
    previousMeetingDate=nextMeeingDate;
    nextMeeingDate=new Date(+nextMeeingDate+MeetingGapDays*3600000*24);
  }
  return previousMeetingDate;
}


// https://tc39.github.io/ecma262/#sec-array.prototype.includes
if (!Array.prototype.includes) {
  Object.defineProperty(Array.prototype, 'includes', {
    value: function(searchElement, fromIndex) {

      if (this == null) {
        throw new TypeError('"this" is null or not defined');
      }

      // 1. Let O be ? ToObject(this value).
      var o = Object(this);

      // 2. Let len be ? ToLength(? Get(O, "length")).
      var len = o.length >>> 0;

      // 3. If len is 0, return false.
      if (len === 0) {
        return false;
      }

      // 4. Let n be ? ToInteger(fromIndex).
      //    (If fromIndex is undefined, this step produces the value 0.)
      var n = fromIndex | 0;

      // 5. If n ≥ 0, then
      //  a. Let k be n.
      // 6. Else n < 0,
      //  a. Let k be len + n.
      //  b. If k < 0, let k be 0.
      var k = Math.max(n >= 0 ? n : len - Math.abs(n), 0);

      function sameValueZero(x, y) {
        return x === y || (typeof x === 'number' && typeof y === 'number' && isNaN(x) && isNaN(y));
      }

      // 7. Repeat, while k < len
      while (k < len) {
        // a. Let elementK be the result of ? Get(O, ! ToString(k)).
        // b. If SameValueZero(searchElement, elementK) is true, return true.
        if (sameValueZero(o[k], searchElement)) {
          return true;
        }
        // c. Increase k by 1. 
        k++;
      }

      // 8. Return false
      return false;
    }
  });
}

 

/***
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
Google Drive and Google Sheet
***/
function getFileUnderFolderByName(fileName,folderId){  
  var folder=DriveApp.getFolderById(folderId);
  var files=folder.getFiles();
  while (files.hasNext()) {
     var file = files.next();
     if (file.getName()==fileName)
        return file.getId();
        
}
     return null;
}

function getFolderIdByName(folderName){
   var folders = DriveApp.getFolders();
   while (folders.hasNext()) {
      var folder = folders.next();
      if (folder.getName()==folderName)
        return folder.getId();
    }
  }
  





function publishDoc(documentid){

  var currentSheet = DriveApp.getFileById(documentid);
  currentSheet.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  


}


function getEditableUrl(documentid){
   return "https://docs.google.com/spreadsheets/d/"+documentid+"/edit#gid=0"
}

function getPublishedUrl(documentid){

   return "https://docs.google.com/spreadsheets/d/"+documentid+"/pubhtml"
}

function getDownloadUrl(documentid){

   return "https://docs.google.com/spreadsheets/d/"+documentid+"/export?format=xlsx"
}

function revokeEditing(documentid){

  var currentSheet = DriveApp.getFileById(documentid);
  currentSheet.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  
}

function allowEditing(documentid,email){

  var currentSheet = DriveApp.getFileById(documentid); 
  currentSheet.addEditor(email);  
  
}


/**
This is used when create new agendas
**/
function saveTemplateSpreadsheetToFolder(sourceSheetid, newSheetFileName, destinationFolderId){

    var destFolder = DriveApp.getFolderById(destinationFolderId); 
    var newSheetId = DriveApp.getFileById(sourceSheetid).makeCopy(newSheetFileName, destFolder).getId();
    return newSheetId;

   } 


/*copy everything from source to target 
Tested, even if targetSheet originally is smaller than sourcesheet scope, google sheet will expand targetsheet automatically
*/
function SyncSheet(sourceSheet, targetSheet){

  targetSheet.clearContents();
  var MaxRow=sourceSheet.getMaxRows();
  var MaxCol=sourceSheet.getMaxColumns(); 
  
  
  var obj=sourceSheet.getSheetValues(1, 1, sourceSheet.getMaxRows(), sourceSheet.getMaxColumns());

  var rangeString = "A1:"+columnToLetter(MaxCol)+(obj.length);
  
  targetSheet.getRange(rangeString).setValues(obj);

}


function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function letterToColumn(letter)
{
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++)
  {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}


/*
 This will return the spreadsheet object with the given file name 
*/
function getOrCreateSpreadsheetByName(fileName,folderId){

   var sheetId=getFileUnderFolderByName(fileName,folderId);
   if (null==sheetId){
      sheetId=SpreadsheetApp.create(fileName).getId();
      var file = DriveApp.getFileById(sheetId);
      DriveApp.getFolderById(folderId).addFile(file);
      file.getParents().next().removeFile(file);

   }
   return SpreadsheetApp.openById(sheetId);
}


function getOrCreateSheetBySheetName(spreadsheet, sheetName){
   var sheet=spreadsheet.getSheetByName(sheetName);
   if(null==sheet)
      sheet=spreadsheet.insertSheet().setName(sheetName);
   return sheet;
}



function rowToObject(header, rowobj){
  var obj={};
  for (var i=0;i<header.length;i++){
    obj[header[i]]=rowobj[i];  
  }

  return obj;
}


function objToRow(header, valueObj){
   var row=new Array();
   for (var i=0;i<header.length;i++){
     row.push(valueObj[header[i]])
    }
    return row;
}

function clearSheetAndWriteValuesBack(sheet, header, valueArrayList){
   var values=new Array();
   values[0]=header;
  
   for (var i=0;i<valueArrayList.length;i++){
     values[i+1]= objToRow(header, valueArrayList[i]);
   }
   var rangeName = "A1:"+columnToLetter(header.length)+(valueArrayList.length+1).toString();
   var range = sheet.getRange(rangeName);
  
    sheet.clearContents();    
    range.setValues(values);

}


/*
  Transform Sheet data (with header) into object list and return (no filter applied)
*/
function getFullList(sheet){   
    var obj = sheet.getSheetValues(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
    var header=getHeader(sheet);
    var fullList=new Array();
  
    //get filtered List
    var currentRow = 2;
    while(null!= obj[currentRow-1][0] &&  obj[currentRow-1][0].toString().length>0){  
         var currentObj=rowToObject(header,obj[currentRow-1]);
            fullList.push(currentObj);        
            currentRow ++;
   }
   return fullList;      
}



/*
  get filtered list and clean the existing sheet
*/
function getRefreshedList(sheet, cutOverDate, dateColNum){   
    var obj = sheet.getSheetValues(1, 1, sheet.getMaxRows(), 5);
    var header=getHeader(sheet);
    var filterredList=new Array();
  
    //get filtered List
    var currentRow = 2;
    while(null!= obj[currentRow-1][0] &&  obj[currentRow-1][0].toString().length>0){  
      if(obj[currentRow-1][dateColNum]>cutOverDate){     
         var currentObj=rowToObject(header,obj[currentRow-1]);
            filterredList.push(currentObj);        
      }              
            currentRow ++;
   }
  
   clearSheetAndWriteValuesBack(sheet, header, filterredList);
       
  
   return filterredList;      
}



function getHeader(sheet){
  var obj = sheet.getSheetValues(1, 1, 1, sheet.getMaxColumns());
  var header=new Array();
  var currentColum = 1;
    while(null!= obj[0][currentColum-1] &&  obj[0][currentColum-1].toString().length>0){  
      header.push(obj[0][currentColum-1].toString())             
      currentColum ++;
    }
  return header;

}


/**
Reset sheet header to align with defined value
**/
function setHeader(header, sheet){
   var values=new Array();
   values[0]=header;
  

   var rangeName = "A1:"+columnToLetter(header.length)+"1";
   var range = sheet.getRange(rangeName);  
   range.setValues(values);
}


/**
If meeting sheet exist, then return;
If meeting sheet not exist, then creat using global config and return;
**/
function getMeetingSheetIdByDate(meetingdate,env){
    var meetingAgendaFileName=getMeetingSheetNameByDate(meetingdate);
    var sheetId=getFileUnderFolderByName(getMeetingSheetNameByDate(meetingdate),env.weeklyAgendaFolderId);
    return sheetId;
}

function createNewMeetingSheetByDate(meetingDate, env){
     var weeklyAgendaTemplateFileID= env.weeklyAgendaTemplateFileId;        
     return saveTemplateSpreadsheetToFolder(env.weeklyAgendaTemplateFileId, getMeetingSheetNameByDate(meetingDate), env.weeklyAgendaFolderId);
}







/***
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
Scheduler Sheet
***/

/**
Centrolized header definition
1) reset the header to re-enforce format
2) get the object list from the sheet
**/
function getHeadersBySheet(sheet){
   var sheetName=sheet.getName();
   switch(sheetName) {
      case 'members':
         return ['Name','Email','InActive'];
      case 'role-definition':
         return ['roleName','entryRoleName','numOfRolesPerMeeting','roleMinSec','roleTargetSec','roleMaxSec','roleBuzzSec'];
      case 'history':
         return ['Date','Name','Role'];
      case 'absent':
         return ['StartDate','EndDate','Name','Role'];
      case '2019Template':
         return ['RoleName','RoleDisplayName'];
      case 'registry':
         return ['Date','Role','Name'];
      default:
          return null;
   }
   return null;
}



/**
Common shared function to use header definition and specified sheet 
1) reset the header to re-enforce format
2) get the object list from the sheet
**/
function getValueListFromSheet(sheet){
  setHeader(getHeadersBySheet(sheet),sheet);
  return getFullList(sheet); 
}




/**
Common shared function to get key value config from sheet using A and B columne
**/
function getConfigFromSheet(sheet){ 
    var obj =sheet.getSheetValues(1, 1, sheet.getMaxRows(), 2);
    var row=0;
    var col=0; 
    var config={};   
    //loop the mem rows
    var currentRow = 2;
    while(null!= obj[currentRow-1][0] &&  obj[currentRow-1][0].toString().length>0){           
            config[obj[currentRow-1][0]]=obj[currentRow-1][1];             
            currentRow ++;
   }   
   return config;
}


/***
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
Tab Members
***/



/**
Current fields: 
Name	Email	InActive
**/

/*
get memberList Containing (only filter out active members)
Containing,
Basic Contact Info
*/
function getActivememberList(memberList){
    var ActiveMemList=[];
    for (var i=0;i<memberList.length;i++){
        if(memberList[i].InActive==false){
           ActiveMemList.push(memberList[i]);
        }
         
    }
    return ActiveMemList;    
}


function getMailByNameFromMemberList(name,memberList){
   for (var i=0;i<memberList.length;i++){
     if(memberList[i].Name.localeCompare(name)==0){
        return memberList[i].Email;
     }     
   }   
   return null;
}

function getMailListByNameList(nameList,memberList){
   var MailList="";
   var MailSplitter =",";
   for (var i=0;i<memberList.length;i++){
     for(var j=0; j<nameList.length;j++){
      if(memberList[i].Name.localeCompare(nameList[j])==0){
        MailList =MailList+ memberList[i].Name+ '<'+memberList[i].Email+ '>'+MailSplitter;  
      }  
     }
   }
   return MailList;
}



/*
 return mail list, 
 filterActiveOnly should be true or false (bool)
*/

function getMailList(memberList){
   var MailList="";
   var MailSplitter =",";
   for (var i=0;i<memberList.length;i++){
          MailList =MailList+ memberList[i].Name+ '<'+memberList[i].Email+ '>'+MailSplitter;   
          //MailList =MailList+memberList[i].Email+ MailSplitter; 
   }
   return MailList;
}

/***
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
Tab Roles
***/


/**
Current fields: 
roleName	entryRoleName	numOfRolesPerMeeting	roleMinSec roleTargetSec roleMaxSec roleBuzzSec
**/


/*
According to roleName and RoleLimitationList, get the entry requirement for a certain role
If there's no limitation then return null
*/
function getLimitationForRole(RoleName,RoleInfoList){
    for (i=0;i<RoleInfoList.length;i++){
          if(RoleInfoList[i].roleName.localeCompare(RoleName)==0 )
            {
                if(RoleInfoList[i].entryRoleName.length>0)
                  return RoleInfoList[i].entryRoleName;
                else
                  return null;
            }
      }//for 

    return null;
}



/***
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
Tab History
***/

/**
Current fields: 
Date	Name Role
**/




/*
add new items into history
*/
function mergeNewHistory(historyList, newItemList, newDate){
   var newHistoryList;
   
   newHistoryList=deleteExpiredHistory(historyList,365*2);
   newHistoryList=deleteHistoryByDate(newDate,newHistoryList);  
   for( i=0;i<newItemList.length;i++){
      var newItem={};
      newItem.Date=newDate;
      newItem.Role=newItemList[i][0];
      newItem.Name=newItemList[i][1];
      newHistoryList.push(newItem);
   }
   return newHistoryList;
}


/*
For a given date, delete existing items from the list
*/
function deleteHistoryByDate(certainDate, historyList){
   var newHistoryList=new Array();
   for (var i=0;i<historyList.length;i++){  
        if(! isSameDate(historyList[i].Date,certainDate))
               newHistoryList.push(historyList[i]);
      
   }
   return newHistoryList;
}

/*
filter out expired items
*/
function deleteExpiredHistory(historyList, expireDays){
       var cutOverDate=new Date();
       cutOverDate.setDate(cutOverDate.getDate()-expireDays);
       var newHistoryList=new Array();
       for (var i=0;i<historyList.length;i++){          
          if(historyList[i].Date>cutOverDate)
               newHistoryList.push(historyList[i]);      
       }       
       return newHistoryList;
}



/*
For a given date list filter out qualified history
*/
function history_filterHistoryListByDateList(dateList, historyList){
   var certainDateHistoryList=new Array();
   for (var i=0;i<historyList.length;i++){   
       for (var j=0;j<dateList.length;j++){  
          if(isSameDate(historyList[i].Date,dateList[j]))
               certainDateHistoryList.push(historyList[i]);
       }
   }
   return certainDateHistoryList;
}


/*
read history data and directly calculate to get last role date
Return a hashmap,
memLastRoleHisMap[Name][Role]=LatestDate
*/
function history_getMemLastRoleHistory(historyList,daysTreatAsInactive){
    var memLastRoleHisMap = {}; 
    var today= new Date();
    var currentDayGap=0;
    //loop the history rows, get Date, UniqueId, Name, Role    
    for (var i=0;i<historyList.length;i++){      
       var currentDate= historyList[i].Date;
       var currentName= historyList[i].Name;
       var currentRole=historyList[i].Role;               
        //if new member
       if(!(currentName in memLastRoleHisMap)){
        var currentUserMap={};
        currentUserMap[currentRole] = currentDate;
        currentUserMap["lastmeetingDate"] = currentDate; 
        currentUserMap["daysSinceLastMeeting"] = dateDiff(today,currentDate); 
        memLastRoleHisMap[currentName]=currentUserMap;
       }else{
         
         //update lastmeetingDate 
         if(memLastRoleHisMap[currentName]["lastmeetingDate"]<currentDate){
           memLastRoleHisMap[currentName]["lastmeetingDate"] = currentDate;
           currentDayGap=dateDiff(today,currentDate);
           memLastRoleHisMap[currentName]["daysSinceLastMeeting"] = currentDayGap;
           if (currentDayGap>daysTreatAsInactive)
             memLastRoleHisMap[currentName]["isActive"]= "no";
           else
             memLastRoleHisMap[currentName]["isActive"]= "yes";
         }
         
         //member exist, role not exist
         if(!(currentRole in memLastRoleHisMap[currentName]))
           memLastRoleHisMap[currentName][currentRole]=currentDate;
         //member exist, role exist, role date need update)
         else if(memLastRoleHisMap[currentName][currentRole]<currentDate)
           memLastRoleHisMap[currentName][currentRole]=currentDate;
         //else do nothing 
      } 
     }
    
   return memLastRoleHisMap;     
}

/***
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
Tab Absent
***/

/*
Current fields: 
StartDate	EndDate	Name Role
*/


/**
For a given meeting date, return absent list
**/
function getAbsentListForCertainDate(meetingDate,absentList){
  var absentListForCertainDate=new Array();
  var absentRoleNameListForCertainDate=new Array();
  
  for(var i=0;i<absentList.length;i++){
    if(   ((absentList[i].StartDate<=meetingDate) && (absentList[i].EndDate>=meetingDate)) || (isSameDate(absentList[i].StartDate,meetingDate))){
    
       if(absentList[i].Role.length>0){
           var item={};
           item.Role=absentList[i].Role;
           item.Name=absentList[i].Name;
           absentRoleNameListForCertainDate.push(item);
       }          
       else    
         absentListForCertainDate.push(absentList[i].Name);
    
    }
       
  }
  
  var absentInfoForCertainDate={};
  
  absentInfoForCertainDate.absentNameList=absentListForCertainDate;
  absentInfoForCertainDate.absentRoleNameList=absentRoleNameListForCertainDate;
  return absentInfoForCertainDate;
}



/*
Delete expired absent list
*/
function deleteExpiredAbsentItems(absentList){
   var newAbsentList=new Array();
   var today=new Date();
   for (var i=0;i<absentList.length;i++){  
        if(absentList[i].EndDate>today)
            newAbsentList.push(absentList[i]);
      
   }
   return newAbsentList;
}


/***
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
Sheet Meeting Agendas
***/

function setAgendaStatus(agendaSheet, status){
     agendaSheet.getRange('J2').setValue(status);
}

function setAgendaStatusList(agendaSheet, statusList){
  agendaSheet.getRange('J2').setDataValidation(null);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statusList)
    .setAllowInvalid(false)
    .setHelpText('Please set the correct status.')
    .build();
   agendaSheet.getRange('J2').setDataValidation(rule);
}

function getAgendaStatus(agendaSheet){
    return agendaSheet.getRange('J2').getValue();
}


/**
   Reset the given meeting agenda by,
   1) Reset the member sheet with updated member list
   2) OverRide Agenda and Forcast for the meeting Agenda from Template
   3) Set the Doc to be Shared
**/
//prepare the upcomming meeting agenda to be ready to populate content
function resetMeetingAgenda(meetingSpreadSheet,templateSpreadSheet, memberSheet){
  //update MemList
  SyncSheet(memberSheet,meetingSpreadSheet.getSheetByName("Member"));
  
  //refresh the Agenda
  var oldSheet=meetingSpreadSheet.getSheetByName("Agenda");
  if (null!=oldSheet)
    meetingSpreadSheet.deleteSheet(oldSheet);
  templateSpreadSheet.getSheetByName("Agenda").copyTo(meetingSpreadSheet).setName("Agenda");
  
  oldSheet=meetingSpreadSheet.getSheetByName("ForeCast");
  if (null!=oldSheet)
    meetingSpreadSheet.deleteSheet(oldSheet);
  templateSpreadSheet.getSheetByName("ForeCast").copyTo(meetingSpreadSheet).setName("ForeCast");
  
  //make sure the doc is shared
  publishDoc(meetingSpreadSheet.getId());

}

/**

For a given agenda, clean the roles
**/
function cleanMeetingAgenda(agendaSheet){
  agendaSheet.getRange('D5:D28').clearContent();
}

/**
get to-be-archived content as history from sheet
** This Only works for the existing template
**/
function getAgendaFromSheet(agendaSheet){
   var obj=agendaSheet.getSheetValues(5, 3, agendaSheet.getMaxRows(), 2);
   var newObj=new Array();
   var role;
   var index;
   for (var i=0;i<obj.length;i++){
       role=obj[i];
       index=role[0].indexOf("#");
       if(null!=index && index>0)
         role[0]=role[0].slice(0, index-1);
       if ((null!=obj[i][1]) && obj[i][1].length>0 && (role[0]!="Sergeant at Arms")&& (role[0]!="Club Business"))
         newObj.push(role);
   }
   
   return newObj;
}

function getNameByRoleFromMeetingAgendaItems(role, meetingItemList){
    for(var i=0;i<meetingItemList.length;i++){
       if( meetingItemList[i][0]==role)
        return meetingItemList[i][1];
    }
    return null;
}




/**populate content to agenda

**/
function populateScheduedAgendaToSheet(agendaList, meetingSp,env){

   
   var agendaSheet=meetingSp.getSheetByName("Agenda");
   var forcastSheet=meetingSp.getSheetByName("ForeCast");
   
   //read planned roleList 
   var obj=agendaSheet.getSheetValues(5, 3, agendaSheet.getMaxRows(), 1);
   matchNWrite(obj,agendaSheet,'D',5,agendaList[0],env);
   agendaSheet.getRange("C2").setValue(agendaList[0].meetingDate);
   setAgendaStatus(agendaSheet,"Scheduled");
   obj=forcastSheet.getSheetValues(5, 2, forcastSheet.getMaxRows(), 1);
   matchNWrite(obj,forcastSheet,'C',5,agendaList[1],env);
   forcastSheet.getRange("C4").setValue(agendaList[1].meetingDate);
   forcastSheet.getRange("F4").setValue(agendaList[2].meetingDate);

   obj=forcastSheet.getSheetValues(5, 5, forcastSheet.getMaxRows(), 1);
   matchNWrite(obj,forcastSheet,'F',5,agendaList[2],env);

}



function matchNWrite(plannedRoleObj,pupulateSheet,populateCol,populateStartRow,agendaFullList,env){
   var PlannedRoleNameList=[];
   for (var i=0;i<plannedRoleObj.length;i++){
      var currentRole=plannedRoleObj[i][0];
      if(null!=currentRole && currentRole.length>0)
         PlannedRoleNameList.push(currentRole);
   }
   var agendaList=agendaFullList.meetingItems;
   var SAA = {
               RoleName : 'Sergeant at Arms',
               RoleDisplayName:'Sergeant at Arms',
               Name:env.config["SAA"]
               }; 
   agendaList.push(SAA);
   
   var ClubBusiness = {
               RoleName : 'Club Business',
               RoleDisplayName:'Club Business',
               Name:env.config["President"]
               }; 
   agendaList.push(ClubBusiness);
   
   var roleObj=new Array();   
   for(var i=0;i<PlannedRoleNameList.length;i++){      
       roleObj[i]=new Array();
       roleObj[i][0]="";
       for (var j=0;j<agendaList.length;j++){   
            var test1=PlannedRoleNameList[i];
            var test2=agendaList[j].RoleDisplayName;
            var test3=(test1.localeCompare(test2)==0)
            if(PlannedRoleNameList[i].localeCompare(agendaList[j].RoleDisplayName)==0){
              if(agendaList[j].Name==null)
                 roleObj[i][0]= '';    
              else
                 roleObj[i][0]= agendaList[j].Name;             
            }
         }   
   }   
  var rangeName=populateCol+populateStartRow+":"+ populateCol +(roleObj.length+4);
  pupulateSheet.getRange(rangeName).setValues(roleObj);   


}


function getNameListByRoleFromAgendaSheet(roleName,agendaSheetItemList){
     var nameListString="";
     for (var i=0;i<agendaSheetItemList.length;i++){
        if(agendaSheetItemList[i][0]==roleName)
          nameListString=nameListString+agendaSheetItemList[i][1]+",";
     }

    return nameListString;
}


/***
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
tab registry
***/

/*
Current fields: 
Date Name Role
*/


/**
For a given meeting date, return matched registry list : RegistryList(Name, Role),
**/
function getRegistryInfoForCertainDate(meetingDate,registryList){
  var matchedRegistryList=new Array();
  
  for(var i=0;i<registryList.length;i++){
    if((isSameDate(registryList[i].Date,meetingDate))){       
       var item={};
       item.Name=registryList[i].Name;
       item.Role=registryList[i].Role;
       matchedRegistryList.push(item);       
       }
  }  
  return matchedRegistryList;
}



/*
Delete expired registerred list
*/
function deleteExpiredRegistryItems(registryList){
   var newRegistryList=new Array();
   var today=new Date();
   for (var i=0;i<registryList.length;i++){  
        if(registryList[i].Date>today)
            newRegistryList.push(registryList[i]);
      
   }
   return newRegistryList;
}



/***
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
cmd generate Agenda
***/


/**
Reconcile to get new Agenda,
Reset spreadsheet
Populate content 
return document id
**/
function generateAndPopulateAgenda(dateList,meetingSpreadSheet,env){
   var templateSpreadSheet=SpreadsheetApp.openById(env.weeklyAgendaTemplateFileId);
   var ss= SpreadsheetApp.openById(env.schedulerSpreadsheetId);     
   resetMeetingAgenda(meetingSpreadSheet,templateSpreadSheet, ss.getSheetByName('members'));   

   var registeredMemberList=getValueListFromSheet(ss.getSheetByName("members"));
   var rolesInfoList=getValueListFromSheet(ss.getSheetByName("role-definition"));   
   var history=getValueListFromSheet(ss.getSheetByName("history"));   
   var emptyAgendaItems=getValueListFromSheet(ss.getSheetByName(env.config['WeeklyAgendaTemplateFileName']));
   var absentList=getValueListFromSheet(ss.getSheetByName("absent"));   
   var registryList=getValueListFromSheet(ss.getSheetByName("registry"));   
   
   var memHistoryMap=history_getMemLastRoleHistory(history,env.config['AbsentDaysTreatAsInActive']);
   
   var sheduledMeetingList=generateScheduleContentForGiveDateList(dateList,emptyAgendaItems,absentList, memHistoryMap,registeredMemberList,rolesInfoList,registryList);
   
   populateScheduedAgendaToSheet(sheduledMeetingList,meetingSpreadSheet,env);
   
}




/***********************************************************************
Core Scheduling Start
***********************************************************************
***********************************************************************
***********************************************************************
***********************************************************************
**/

/**
For a given date and a list of upcomming meeting date list
//todo interface change
**/
function generateScheduleContentForGiveDateList(dateList,emptyAgendaItems,absentList, memHistoryMap,registeredMemberList,RoleInfoList,registryList){  
   var sheduledMeetingList=new Array();   
   //only return name string list
   var activeMemList=filterInActiveMemberNameList(registeredMemberList);
   
   for(var i=0;i<dateList.length;i++){
       sheduledMeetingList=scheduleForCertainMeeting(dateList[i],sheduledMeetingList,emptyAgendaItems,activeMemList,RoleInfoList,memHistoryMap,absentList,registryList);
    }
    
   return sheduledMeetingList;
}

/**
For a given date, 
sheduledMeetingList: used to avoid duplicate role assignment
templateAgenda: initial role list
activeMemList: memeber list to select from
RoleInfoList: role info 
memHistoryMap: member's history attendance
absentList: member's absent registry
registerredRoleList:member's role registry

Result: generate a new schedule and push to existing scheduled meeting list and return
**/
function scheduleForCertainMeeting(meetingDate,sheduledMeetingList,emptyAgendaItems,activeMemList,RoleInfoList,memHistoryMap,absentList,registryList){ 
  
    var absentListOnMeetingDate=getAbsentListForCertainDate(meetingDate,absentList);  //Object 
    var absentNameList=absentListOnMeetingDate.absentNameList;
    var absentRoleNameList=absentListOnMeetingDate.absentRoleNameList;  
  
    var registryList=getRegistryInfoForCertainDate(meetingDate,registryList); //RegistryList(Name, Role, Assigned),MatchedName
    
    var newAgenda={};
    newAgenda.meetingItems=duplicateEmptyAgendaItems(emptyAgendaItems);
    newAgenda.meetingDate=meetingDate;
    
    //assign all registerred roles first
    newAgenda=assignRegisterredRole(absentNameList,registryList, newAgenda);
    
    //filter out absent member; inactive member; and member already have pre-assigned role
    var filteredMemList=filterAbsentMembers(activeMemList,absentNameList);
    filteredMemList=filterMemeberAlreadyInAgenda(filteredMemList, newAgenda.meetingItems); 
   
    // loop through agenda, assign rest agenda role using active; available member list;
    for(var i=0;i<newAgenda.meetingItems.length;i++){
        if(null==newAgenda.meetingItems[i].Name || newAgenda.meetingItems[i].Name.toString().length==0){
               var toBeAssignedRole=newAgenda.meetingItems[i].RoleName;
               var entryRole=getLimitationForRole(toBeAssignedRole,RoleInfoList);
               //copy current available user list
               var baseMemList=filteredMemList.slice(0);
               //filter out member don't want to do this role
               baseMemList=filteredMemDontWantTheRole(toBeAssignedRole,baseMemList,absentRoleNameList);
               //filter out member already have same role in existing scheduled meeting list
               baseMemList=filterMemeberListRemovingMemberWithSameRoleScheduled(toBeAssignedRole,baseMemList,sheduledMeetingList);
               //filter out member hasn't done the entry role
               baseMemList=filterMemberListRemovingMemberNotQualifiedForRole(entryRole,baseMemList,memHistoryMap);
               //get the highest ranking member
               var topQualifiedMem=getFirstRankingMemForARole(toBeAssignedRole,baseMemList,memHistoryMap);
               if(topQualifiedMem!=null){
                  newAgenda.meetingItems[i].Name =topQualifiedMem;
                  filteredMemList=removeMemberFromList(topQualifiedMem,filteredMemList);
               }
        }
    }
    
    sheduledMeetingList.push(newAgenda);
    return sheduledMeetingList;
}



/***********************************************************************
Core Scheduling End
***********************************************************************
***********************************************************************
***********************************************************************
***********************************************************************
**/

/***********************************************************************
Ranking Start
***********************************************************************
***********************************************************************
***********************************************************************
***********************************************************************
**/


//memHistoryMap[Name][Role]=LatestDate
function getFirstRankingMemForARole(role,memList,memHistoryMap){
   
   if(memList==null || memList.length==0)
     return null;
   if(memList.length==1)
     return memList[0];

   var topQualifedMem=memList[0];
   for (i=1;i<memList.length;i++){
      topQualifedMem=getTheWinner(topQualifedMem, memList[i], role,memHistoryMap);
   }      
   return topQualifedMem;
}


function getTheWinner(topMem, newMem,role, memHistoryMap){
   if (null==memHistoryMap[topMem]|| null==memHistoryMap[topMem][role])
    return topMem;
   if (null==memHistoryMap[newMem]|| null==memHistoryMap[newMem][role])
    return newMem;
    
   if (memHistoryMap[topMem][role]< memHistoryMap[newMem][role])
       return topMem;
   else 
       return newMem;
}


/***********************************************************************
Ranking End
***********************************************************************
***********************************************************************
***********************************************************************
***********************************************************************
**/


/**
***********************************************************************
Assign registerred role start
***********************************************************************
***********************************************************************
***********************************************************************
***********************************************************************
**/


/**
Given the empty agenda, try to fit in all registered roles and then return the filled agenda
**/
function assignRegisterredRole(absentNameList,registryList, newAgenda){
   var registry;
   for(var i=0;i<registryList.length; i++){
      registry=registryList[i];
      //check it's valid registration
      if (registry.Name==null || registry.Name.length==0)
        continue;     
      //check if the user not registerred for absent
      if (absentNameList.includes(registry.Name))
        continue;
      //find a slot to fit in
      for(var j=0;j<newAgenda.meetingItems.length; j++){
          //registerred role still available
          if(isRoleAssignableForRegistry(registryList[i],newAgenda.meetingItems[j])){
           newAgenda.meetingItems[j].Name=registryList[i].Name;
           break;
          }   
      }              
   }             
   return newAgenda;
}


/**
check if the registry: Name, Role is suitable to assign to current meetingitem
registry: Name, Role
meetingItem: RoleName	RoleDisplayName

1) Role name must match
2) Meeting Item must be available to be assigned

**/
function isRoleAssignableForRegistry(registry, meetingItem){   
    if (registry.Role!=meetingItem.RoleName)
       return false;
    if (meetingItem.Name!=null && meetingItem.Name.length>0)
       return false;       
    return true;   
}

/**
***********************************************************************
Assign registerred role End
***********************************************************************
***********************************************************************
***********************************************************************
***********************************************************************
**/


/**
***********************************************************************
Filter MemberList Start
***********************************************************************
***********************************************************************
***********************************************************************
***********************************************************************
**/

function filterInActiveMemberNameList(memList){
   var filteredMemList=new Array();
   for (var i=0; i<memList.length;i++){
      if (memList[i].InActive==true)
        continue;
      filteredMemList.push(memList[i].Name);
   }
   return filteredMemList;
}

function filterAbsentMembers(memList,absentNameList){
   var filteredMemList=new Array();
   for (var i=0; i<memList.length;i++){
      if (absentNameList.includes(memList[i]))
        continue;
      filteredMemList.push(memList[i]);
   }
   return filteredMemList;
}
/**

Registered MemList: Name	Email	InActive
Return a new member list,
1) remove any member that already has a role in current meetingItems
**/
function filterMemeberAlreadyInAgenda(memList, meetingItems){
   var filteredMemList=new Array();
   for (var i=0; i<memList.length;i++){
      if (isInAgenda(memList[i], meetingItems))
        continue;      
      filteredMemList.push(memList[i]);
   }
   return filteredMemList;
}


/**

MemList: Name	Email	InActive
Return a new member list, without member already has same role in scheduled meetings
**/
function filterMemeberListRemovingMemberWithSameRoleScheduled(role, memList,scheduledMeetingList){
   var filteredMemList=new Array();
   for (var i=0; i<memList.length;i++){
      if (! hasSameRoleInScheduleList(role,memList[i],scheduledMeetingList))    
        filteredMemList.push(memList[i]);
    }
    return filteredMemList;
}


/**
Filter out member that hasn't done the entry role yet
**/
function filterMemberListRemovingMemberNotQualifiedForRole(entryRole,memList,memHistoryMap){
   
   if(entryRole==null || entryRole.length==0)
     return memList;

   var filteredMemList=new Array();
   for (var i=0; i<memList.length;i++){
      if (hasDoneEntryRole(entryRole,memList[i],memHistoryMap))                
         filteredMemList.push(memList[i]);
    }
    return filteredMemList;
}

/**
absentRoleNameList: Name Role

**/
function filteredMemDontWantTheRole(role,memList,absentRoleNameList){
   var filteredMemList=new Array();
   for (var i=0; i<memList.length;i++){
        if(! isAvodingTheRole(role,memList[i],absentRoleNameList))
          filteredMemList.push(memList[i]);
    }
    return filteredMemList;
}

function isAvodingTheRole(role,name,absentRoleNameList){
    for (var i=0; i<absentRoleNameList.length;i++){
        if(absentRoleNameList[i].Role==role && absentRoleNameList[i].Name==name)
          return true;
    }
    return false;
}

function removeMemberFromList(name,memList){
   var filteredMemList=new Array();
   for (var i=0; i<memList.length;i++){
      if (memList[i]!=name)                
         filteredMemList.push(memList[i]);
    }
    return filteredMemList;
}

/**
***********************************************************************
Filter MemberList End
***********************************************************************
***********************************************************************
***********************************************************************
***********************************************************************
**/



/**
***********************************************************************
Common Start
***********************************************************************
***********************************************************************
***********************************************************************
***********************************************************************
**/

function hasSameRoleInScheduleList(role,name,scheduledMeetingList){
   
    for (var i=0; i<scheduledMeetingList.length;i++){
        if(hasSameRoleInMeetingItems(role,name,scheduledMeetingList[i].meetingItems))
          return true;
    }
    return false;
}


function hasSameRoleInMeetingItems(role,name,meetingItems){
   
    for (var i=0; i<meetingItems.length;i++){
        
        if (meetingItems[i].RoleName==role && meetingItems[i].Name!=null && meetingItems[i].Name==name)
           return true;
    }
    
    return false;
}

/**
check if a given name is already in the agenda
**/
function isInAgenda(name, meetingItems){

   for (var i=0; i<meetingItems.length;i++){
        if (meetingItems[i].Name==name)
          return true;
   }
   return false;
}


/**
check if the name has done entry role 
memLastRoleHisMap[Name][Role]=LatestDate
**/
function hasDoneEntryRole(entryRole,name,memHistoryMap){

   if(memHistoryMap[name]==null)
     return false;
   if(memHistoryMap[name][entryRole]==null)
     return false;
   return true;
}


function duplicateEmptyAgendaItems(emptyAgendaItems){
   var agendaItems=JSON.parse(JSON.stringify(emptyAgendaItems));
   return agendaItems;
}

/**
***********************************************************************
Common End
***********************************************************************
***********************************************************************
***********************************************************************
***********************************************************************
**/

/***
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
cmd archive agenda
***/
/**
for special meeting 
1) clean all the content
2) clean status rule
3) set agenda to be readonly
**/
function archiveSpecialMeeting(meetingSpreadSheet){
  var agendaSheet=meetingSpreadSheet.getSheetByName("Agenda");
  cleanMeetingAgenda(agendaSheet);
  setAgendaStatusList(agendaSheet, ["N/A"]);
  setAgendaStatus(agendaSheet,"N/A");
  revokeEditing(meetingSpreadSheet.getId());
}


function archiveMeetingAgendaWithHistory(meetingSpreadSheet,env){
  var agendaSheet=meetingSpreadSheet.getSheetByName("Agenda");
  archiveHistoryByDate(meetingDate,agendaSheet,env);
  setAgendaStatusList(meetingSpreadSheet, ["Archived"]);
  setAgendaStatus(agendaSheet,"Archived");
    revokeEditing(meetingSpreadSheet.getId());
}


/**
Archive Certain meeting to History;
It will get the meeting detail by the given date and archive into History sheet.
**/
function archiveHistoryByDate(meetingDate,agendaSheet,env){
  var agendaItems=getAgendaFromSheet(agendaSheet);
  var historySheet=SpreadsheetApp.openById(env.schedulerSpreadsheetId).getSheetByName('history');
  var historyHeader=getHeader(historySheet);
  var historyList=getValueListFromSheet(historySheet);   

  var newHistoryList=mergeNewHistory(historyList, agendaItems, meetingDate);
  clearSheetAndWriteValuesBack(historySheet, historyHeader, newHistoryList);
}





function archiveMeetingAgendaByDate(meetingDate,env){
  var sheetId=getMeetingSheetIdByDate(meetingDate,env);
  var meetingSpreadSheet=SpreadsheetApp.openById(sheetId);
  var agendaSheet=meetingSpreadSheet.getSheetByName("Agenda");
  archiveHistoryByDate(meetingDate,agendaSheet,env);
  setAgendaStatus(agendaSheet,"Archived");
  revokeEditing(sheetId);
}



/***
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
Lib Trigger
***/

/***
Run Every Day
1) Check if last meeting is set as "Ready To Archive" or "Archived" or "N/A"
2) If not, send email to TM and VPE (With Instructions) to update the item and change status to "Ready To Archive"

**/
function trigger_checkDaily(envCode){
  var env=getEnv(envCode);
  var FirstMeetingDate=env.config['FirstMeetingDate'];
  var MeetingGapDays=env.config['MeetingGapDays'];
  
  
/**
**************************
Run Previous Meeting
**************************
**/
  //get last meeting date
  var tobeArchivedDate=getPreviousMeetingDate(FirstMeetingDate,MeetingGapDays);
  var previousMeetingAgendaSheetId=getMeetingSheetIdByDate(tobeArchivedDate,env);
  
  // non prod env, will create a dummpy old agenda
  if(previousMeetingAgendaSheetId==null){   
   if(envCode=='prod')
     throw 'Error. trigger_checkDaily, last meeting agenda not exist.';
    else{
     //for test purpose
     previousMeetingAgendaSheetId=createNewMeetingSheetByDate(tobeArchivedDate, env);
     var dummyMeetingSheet=SpreadsheetApp.openById(previousMeetingAgendaSheetId).getSheetByName("Agenda");     
     setAgendaStatusList(dummyMeetingSheet, ['N/A']);
     setAgendaStatus(dummyMeetingSheet, 'N/A'); 
     }
  }

  var meetingSpreadSheet=SpreadsheetApp.openById(previousMeetingAgendaSheetId);
  var previousAgendaSheet=meetingSpreadSheet.getSheetByName("Agenda");
  var previousAgendaStatus=getAgendaStatus(previousAgendaSheet);
  
  //archive previous special meeting
  if(previousAgendaStatus=='N/A'){
     archiveSpecialMeeting(meetingSpreadSheet);
  }
  

  //previous meeting not special meeting('N/A') and not Archived
  if(previousAgendaStatus!='N/A' && previousAgendaStatus!='Archived' && previousAgendaStatus!='Ready To Archive'){
     setAgendaStatusList(previousAgendaSheet, ['Ready To Archive','N/A', 'Pending Confirm Attendance']);
     setAgendaStatus(previousAgendaSheet, 'Pending Confirm Attendance');
     sendArchiveReminder(tobeArchivedDate, meetingSpreadSheet.getId(),env);  
     return;
  }
  
  if(previousAgendaStatus=='Ready To Archive'){
     setAgendaStatusList(previousAgendaSheet, ['Archived'])
     setAgendaStatus(previousAgendaSheet, 'Archived'); 
     archiveMeetingAgendaByDate(tobeArchivedDate,env);
  }
  
  

  
/**
**************************
Check Next Meeting
**************************
**/
   
   
   previousAgendaStatus=getAgendaStatus(previousAgendaSheet);
   //Only if previous meeting is archived, then the next meeting agenda will be generated
   if(previousAgendaStatus=='Archived' || previousAgendaStatus=='N/A'){
   
      var dateList=getNextMeetingDates(FirstMeetingDate,MeetingGapDays,2);  
      var nextMeetingSheetId=getMeetingSheetIdByDate(dateList[0],env);
      var nextMeetingSpreadSheet;
      if(nextMeetingSheetId==null){
        nextMeetingSheetId=createNewMeetingSheetByDate(dateList[0], env);
        nextMeetingSpreadSheet=SpreadsheetApp.openById(nextMeetingSheetId);
        generateAndPopulateAgenda(dateList,nextMeetingSpreadSheet,env);
        var agendaSheet=nextMeetingSpreadSheet.getSheetByName("Agenda");
        setAgendaStatusList(agendaSheet, ['Draft','Ready to Publish','N/A']);
        setAgendaStatus(agendaSheet, 'Draft');
     }
   
     nextMeetingSpreadSheet=SpreadsheetApp.openById(nextMeetingSheetId);
     var nextMeetingAgendaStatus=getAgendaStatus(nextMeetingSpreadSheet.getSheetByName("Agenda"));
   
     if(nextMeetingAgendaStatus =='N/A'){
         //clean the meeting agenda
         return;
     }
   
     if(nextMeetingAgendaStatus =='Ready to Publish'){
         trigger_publishAgenda(dateList[0],nextMeetingSpreadSheet,env);
         return;
     }
   
    if(nextMeetingAgendaStatus!='Draft' && dateDiff(new Date(),dateList[0])>3){
       sendNewAgendaApproveReminder(dateList[0], meetingSpreadSheet.getId(),env);  
     }else{
        trigger_publishAgenda(dateList[0],nextMeetingSpreadSheet,env);
   
     }
   
   
    }//If previous agenda is archived
  


}


function trigger_publishAgenda(meetingDate, meetingSpreadSheet, env){
   //label Agenda as published
   var agendasheet=meetingSpreadSheet.getSheetByName("Agenda");
   setAgendaStatusList(agendasheet, ['Published']);
   setAgendaStatus(agendasheet, 'Published');
   email_newAgendaGeneratedReminder(meetingDate,meetingSpreadSheet.getId(),env);

}



/***
Run Every Month
1) Sending VPE Report to commitee members
**/
function trigger_VPEReport(envCode){
  var env=getEnv(envCode);
  var VPEReportId=refreshVPEReport(env);
  revokeEditing(VPEReportId);
  var url=getEditableUrl(VPEReportId);  
  
  var shedulerurl=getEditableUrl(env.schedulerSpreadsheetId);
  //add VPE to have access to scheduler
  /**
  TODO
  var VPE=env.config['VPE'];
  var email=getMailByName(VPE);
  if(email.indexOf("@gmail.com")>-1)
     allowEditing(shedulerurl,email);
  
  sendVPEReport(url,env);  
 **/
}




/***
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
Email
***/
function sendArchiveReminder(tobeArchivedDate,meetingSpreadSheetId,env){
 var meetingSpreadSheet= SpreadsheetApp.openById(meetingSpreadSheetId);
 var agendaUrl = "https://docs.google.com/spreadsheets/d/"+meetingSpreadSheetId+"/edit#gid=0";
 var subject="Pending Action: update previous Meeting Agenda, VPE and TM to Update agenda for " + getFormattedDate(tobeArchivedDate);  
          
 var htmlBody = HtmlService.createHtmlOutputFromFile('ArchiveReminder').getContent();


 var meetingItemList=getAgendaFromSheet(meetingSpreadSheet.getSheetByName("Agenda")); 
 var nameList=new Array(); 
 //var emailList= getMailListByNameList(nameList);
 var TM=getNameByRoleFromMeetingAgendaItems("Toastmaster", meetingItemList);
 var VPE=env.config['VPE'];
 nameList.push(TM);
 nameList.push(VPE);
 var memberList=getValueListFromSheet(SpreadsheetApp.openById(env.schedulerSpreadsheetId).getSheetByName("members")); 
 var emailList= getMailListByNameList(nameList,memberList);
 
 
 var htmlBody = htmlBody.replace("%name%", TM+","+VPE);
 var htmlBody = htmlBody.replace("%url%", agendaUrl);
 
if(env.envCode!='prod')
   emailList=env.debugEmail;
 
 MailApp.sendEmail({
    to: emailList,
    subject: subject,
    htmlBody: htmlBody,

  });
 
}


//send VPE Report
function sendVPEReport(url,env){
 var clubName=env.clubName;
 var subject=clubName+" Committee: Updated VPE Report";       
 var htmlBody = HtmlService.createHtmlOutputFromFile('VPEReport').getContent();
 var nameList=[env.config["VPE"],env.config["SAA"],env.config["President"],env.config["VPM"]]; 
 var memberList=getValueListFromSheet(SpreadsheetApp.openById(env.schedulerSpreadsheetId).getSheetByName("members")); 
 var emailList= getMailListByNameList(nameList,memberList);
 
 var htmlBody = htmlBody.replace("%VPM%", env.config["VPM"]);
 var htmlBody = htmlBody.replace("%url%", url);
 
 if(env.envCode!='prod')
   emailList=env.debugEmail;
 MailApp.sendEmail({
    to: emailList,
    subject: subject,
    htmlBody: htmlBody,

  });
 
}

function sendVPEHelper(url,env){
 var subject="VPE Helper -- How to use scheduler";       
 var htmlBody = HtmlService.createHtmlOutputFromFile('VPEHelper').getContent();
 var nameList=[env.config["VPE"]]; 
  var memberList=getValueListFromSheet(SpreadsheetApp.openById(env.schedulerSpreadsheetId).getSheetByName("members")); 

 var emailList= getMailListByNameList(nameList,memberList);
 
 var htmlBody = htmlBody.replace("%VPE%", env.config["VPE"]);
 var htmlBody = htmlBody.replace("%SchedulerURL%", url);
 var htmlBody = htmlBody.replace("%VPEReportUrl%", getVPEReportURL(env));
  
 if(env.envCode!='prod')
   emailList=env.debugEmail;

 MailApp.sendEmail({
    to: emailList,
    subject: subject,
    htmlBody: htmlBody,

  });
 
}




       
function sendNewAgendaApproveReminder(meetingDate,meetingSheetId, env){
 var clubName=env.clubName;
 var currentAgenda=getAgendaFromSheet(SpreadsheetApp.openById(meetingSheetId).getSheetByName("Agenda"));
 var SchedulerUrl=getEditableUrl(env.schedulerSpreadsheetId);
 var subject=clubName+" VPE Heads up for New Agenda : " + getFormattedDate(meetingDate);       
 var htmlBody = HtmlService.createHtmlOutputFromFile('DraftMeetingPendingApproval').getContent();
 
 var memberList=getValueListFromSheet(SpreadsheetApp.openById(env.schedulerSpreadsheetId).getSheetByName("members")); 
 

 var VPE=env.config["VPE"];
 var VPEEmail=getMailByNameFromMemberList(VPE,memberList);

 var htmlBody = htmlBody.replace("%VPE%", VPE);
 htmlBody = htmlBody.replace("%VPEEmail%",  VPEEmail);
 htmlBody = htmlBody.replace("%AgendaUrl%", getEditableUrl(meetingSheetId));
 htmlBody = htmlBody.replace("%SchedulerURL%", SchedulerUrl);
 

 if(env.envCode!='prod')
   emailList=env.debugEmail;


 MailApp.sendEmail({
    to: emailList,
    subject: subject,
    htmlBody: htmlBody,

  });
 
}       

//triggerred when a new agenda being published
function email_newAgendaGeneratedReminder(meetingDate,meetingSheetId, env){
 
 var currentAgenda=getAgendaFromSheet(SpreadsheetApp.openById(meetingSheetId).getSheetByName("Agenda"));
 var clubName=env.clubName;
 var subject=clubName+" Announced New Meeting : " + getFormattedDate(meetingDate);       
 var htmlBody = HtmlService.createHtmlOutputFromFile('WeeklyMeetingReminder').getContent();
 
 var memberList=getValueListFromSheet(SpreadsheetApp.openById(env.schedulerSpreadsheetId).getSheetByName("members")); 
 var emailList= getMailList(memberList);
 
 //get details from scheduled list 
 var speakers=getNameListByRoleFromAgendaSheet("Speaker",currentAgenda);
 var TM=getNameListByRoleFromAgendaSheet("Toastmaster",currentAgenda);
 TM=TM.replace(",","");
 var VPE=env.config["VPE"];
 var VPEEmail=getMailByNameFromMemberList(VPE,memberList);
 
 var htmlBody = htmlBody.replace("%TM%", TM);
  htmlBody = htmlBody.replace("%TMEmail%", getMailByNameFromMemberList(TM,memberList));
  htmlBody = htmlBody.replace("%VPE%", VPE);
  htmlBody = htmlBody.replace("%VPEEmail%",  VPEEmail);
  htmlBody = htmlBody.replace("%AgendaUrl%", getEditableUrl(meetingSheetId));
  htmlBody = htmlBody.replace("%Speakers%", speakers);

 if(env.envCode!='prod')
   emailList=env.debugEmail;


 MailApp.sendEmail({
    to: emailList,
    subject: subject,
    htmlBody: htmlBody,

  });
 
}


function getNameStringListByRoleFromMeetingAgenda(roleName,meetingItems){
     var nameListString="";
     for (var i=0;i<meetingItems.length;i++){
        if(meetingItems[i].RoleName==roleName)
          nameListString=nameListString+meetingItems[i].Name+",";
     }

    return nameListString;
}

function getNameListByRoleFromMeetingAgenda(roleName,scheduledMeeting){
     var meetingItems=scheduledMeeting.meetingItems;
     var nameList=new Array();
     for (var i=0;i<meetingItems.length;i++){
        if(meetingItems[i].RoleName==roleName)
          nameList.push(meetingItems[i].Name);
     }

    return nameList;
}




/***
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
************************************************************************************************************************************************************************************************
VPE Reports
***/

/*
populate meeting sheet link to VPE report 
*/
function refreshVPEReport(env){
     var VPEReportFileName=env.config["VPEReportFileName"];
     var spreadsheet = getOrCreateSpreadsheetByName(VPEReportFileName,env.rootFolderId);
     var header=new Array("MeetingDate","MeetingUrl");
     var sheet=getOrCreateSheetBySheetName(spreadsheet,"AgendaList");
     var meetingLinkList=generateMeetingLinkList(env.config["FirstMeetingDate"],env.config["MeetingGapDays"],env);

     clearSheetAndWriteValuesBack(sheet, header, meetingLinkList);
    
     //update history report sheet
     sheet=getOrCreateSheetBySheetName(spreadsheet,"HistorySummary");
     reconcileMemberLastRoleHistorySheet(sheet,env);
     
     return spreadsheet.getId();    
}


function getVPEReportURL(env){
 var VPEReportFileName=env.config["VPEReportFileName"];
 var spreadsheet = getOrCreateSpreadsheetByName(VPEReportFileName,env.rootFolderId);
 return getEditableUrl(spreadsheet.getId());

}


/**
Refresh the meeting agenda link (from first meeting date till today)
**/
function generateMeetingLinkList(FirstMeetingDate,MeetingGapDays,env) {  
  var MeetingLinkList = [];
  var currentMeetingDate=FirstMeetingDate;
  var meetingLink="";
  var meetingSheetId="";
  var today=new Date();
  
  while(currentMeetingDate<today){
    meetingSheetId=getMeetingSheetIdByDate(currentMeetingDate,env);
    if (meetingSheetId!=null){
        var currentMeetingLink={
                  MeetingDate : currentMeetingDate,
                  MeetingUrl: getEditableUrl(meetingSheetId)
          }    
        MeetingLinkList.push(currentMeetingLink); 
    
    }
   currentMeetingDate=new Date(+currentMeetingDate+MeetingGapDays*3600000*24);
  }
  
  var nextMeetingDates=getNextMeetingDates(FirstMeetingDate,MeetingGapDays,1);
  meetingSheetId=getMeetingSheetIdByDate(nextMeetingDates[0],env);
  if (meetingSheetId!=null){
     var currentMeetingLink={
                  MeetingDate : nextMeetingDates[0],
                  MeetingUrl: getEditableUrl(meetingSheetId)
          }    
     MeetingLinkList.push(currentMeetingLink); 
   }

   
  return MeetingLinkList;     
}



/*
1) clean current sheet
2) calclute from History Sheet to get rawData
3) get all roles 
4) populate to sheet using correct sequence (all the members will be populated, even members without values)
*/
function reconcileMemberLastRoleHistorySheet(lastRoleHisSheet,env){
     lastRoleHisSheet.clearContents();      
     //get the memList & roleList
     
     var ss= SpreadsheetApp.openById(env.schedulerSpreadsheetId);      
     var memList = getValueListFromSheet(ss.getSheetByName("members"));
     var roleList = getValueListFromSheet(ss.getSheetByName("role-definition"));   
     var historyList=getValueListFromSheet(ss.getSheetByName("history"));   
     var reconciledResultMap=history_getMemLastRoleHistory(historyList);
     
     var values=new Array();
     //set header
     values[0]=new Array("Name","LastMeetingDate","DaysSinceLastMeetingDate","IsActive");
     
     for (var i=0;i<roleList.length;i++){
        values[0].push(roleList[i].roleName);
     }
     
     // set content
     var currentMemName;
     var currentRole;
     for (var j=0;j<memList.length;j++){
         currentMemName= memList[j].Name;
         values[j+1]=new Array();
         values[j+1][0]=(currentMemName).toString();
         if(reconciledResultMap[currentMemName]!=null && reconciledResultMap[currentMemName].lastmeetingDate!=null){
           values[j+1][1]=reconciledResultMap[currentMemName].lastmeetingDate;
           values[j+1][2]=reconciledResultMap[currentMemName].daysSinceLastMeeting;
           values[j+1][3]=reconciledResultMap[currentMemName].isActive;
         }else{
           values[j+1][1]='';
           values[j+1][2]='';
           values[j+1][3]='';
         
         }
         for (var k=0;k<roleList.length;k++){
            currentRole = roleList[k].roleName;
            if((currentMemName in reconciledResultMap) && (currentRole in reconciledResultMap[currentMemName]))
               values[j+1][k+4]=reconciledResultMap[currentMemName][currentRole];
            else
               values[j+1][k+4]="";
         
         }     
     }
     var rangeName = "A1:"+columnToLetter(roleList.length+4)+(memList.length+1).toString();
     var range = lastRoleHisSheet.getRange(rangeName);
     range.setValues(values);
}




