//This script created by Alice Keeler will allow you to take a roster of students on the first tab and after executing the script a tab will be created with the students name.

  function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('TemplateTab')
      .addItem('Run TemplateTab', 'templateTabs') //(caption, function name)
  .addSeparator()
  .addItem('Make Copies', 'copyTemplate2') //(caption, function name)
  .addItem('Delete Unedited Tabs', 'deleteBlanks') //(caption, function name)
  .addItem('Email Tab', 'emailTab') //(caption, function name)
   .addItem('Move to Front', 'moveFront') //(caption, function name)
  .addItem('Create New', 'newTemplateTab') //(caption, function name)
      .addSeparator()
  
      .addToUi();
}


function templateTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];//This is an array of all the sheets in the spreadsheet 
  
  sheet.getRange(1,2).setValue('Count');//add a column to the original spreadsheet that will count the number of cells
  
  
  var sheetTemp = ss.getSheets()[1];
  var range = sheet.getDataRange();
  var values = range.getValues();//This creates an array of the range
  
  

  
  
    
 //Run the script to create the tabs 
   ss.setActiveSheet(sheet)
    ss.renameActiveSheet('roster');
   ss.setActiveSheet(sheetTemp);
  ss.renameActiveSheet('template');
   
  
    var lastRow = range.getLastRow();

 
//need to loop going through each item in column A and creating a sheet for each name
    //need to know how to call cell A2... etc... 
    for(var k=1; k < lastRow; k++ ){
      try{
        ss.setActiveSheet(sheetTemp);
        var tabName = values[k][0];
        ss.duplicateActiveSheet();
        ss.renameActiveSheet(tabName);
        ss.moveActiveSheet(k+2);
        var row = k+1;
        sheet.getRange(row,2).setValue('=COUNTA(\''+tabName+'\'!A1:Z900)');
          }
      catch(err){
        tabName = values[k][0]+" " + k;
      ss.duplicateActiveSheet();
        ss.renameActiveSheet(tabName);
      ss.moveActiveSheet(k+2);
      var row = k+1;
        sheet.getRange(row,2).setValue('=COUNTA(\''+tabName+'\'!A1:Z900)');
      }
        }
  
  
  //sets user back to the roster tab. 
  ss.setActiveSheet(sheet);
  ss.getSheets()[0].hideSheet();
  ss.getSheets()[1].hideSheet();
  }



function deleteBlanks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];//This is an array of all the sheets in the spreadsheet 
  var sheetTemp = ss.getSheets()[1];
  var range = sheet.getDataRange();
  var values = range.getValues();//This creates an array of the range
  
  var temp = ss.getSheetByName('template');
  
  sheet.getRange(1,2).setValue('=COUNTA(template!A1:Z900)');
  var count = sheet.getRange(1,2).getValue();
  Logger.log('count '+count);
  
  var last = sheet.getLastRow();
  var sto = last+1;
  
  for(var i=2;i<sto;i++){
    try{
    
    var getCount = sheet.getRange(i,2).getValue();
    var getTabName = sheet.getRange(i,1).getValue();
    Logger.log('getCount '+getCount+' ' +getTabName);
    if(getCount == count){
     var byebye =  ss.getSheetByName(getTabName);
      Logger.log('Sheet to Delete '+byebye);
      ss.deleteSheet(byebye);
    }
    else{}
    }
    catch(err){}
    
  }
  
}
  
  
function copyTemplate(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var ssUrl  = ss.getUrl();
  var ssId = ss.getId();
  var name = ss.getName();
  
  var ask = ui.prompt('How many copies would you like to make?');
  var respond = ask.getResponseText();
  
  if(isNaN(respond)){
    ui.alert('You did not enter a number. Please try again');
  }
  else{
  try{
    for(var i=0; i<respond; i++){
      var j=i+1;
    SpreadsheetApp.openById(ssId).copy(name+' ' +j)
    
  }
    ui.alert('Your '+respond+' copies are in your Google Drive. Copy and Paste your roster into each one and run the TemplateTab script.');
  }
  catch(err){
    ui.alert('Sorry, there was an error.');
  }
  }
 
    
  }

function copyTemplate2(){
  
  
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = ss.getName();
  var ui = SpreadsheetApp.getUi();
  
  var template = ss.getSheetByName('template');
  
  
   var ask = ui.prompt('How many copies would you like to make?');
  var respond = ask.getResponseText();
  
  if(isNaN(respond)){
    ui.alert('You did not enter a number. Please try again');
  }
  else{
  try{
    for(var i=0; i<respond; i++){
      var j=i+1;
   
      var templateTab = SpreadsheetApp.openById('1YVnU-4intMuy0_2eISoj0wUVQ17SDvBVKPAVWwY7-lY');
  var fresh = templateTab.copy('Copy '+j+' '+name);
  var freshUrl = fresh.getUrl();
 var freshId = fresh.getId();
 var freshed = SpreadsheetApp.openById(freshId);
  var templateFresh = freshed.getSheets()[1];
  freshed.deleteSheet(templateFresh);
      
  
      //copy the template to the new spreadsheet copies. 
      
//  template.copyTo(freshed);
  
      ss.getSheetByName('template').copyTo(freshed);
      //show template
      freshed.getSheets()[1].showSheet();
      
    
  }
    ui.alert('Your '+respond+' copies are in your Google Drive. Copy and Paste your roster into each one and run the TemplateTab script.');
  }
  catch(err){
    ui.alert('Sorry, there was an error.');
  }
  }
 
    
  }
  
  
  
  
  
  
  
  



function emailTab(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet = ss.getSheetByName('roster');
  var docId = ss.getId();
  var ui = SpreadsheetApp.getUi();
  var docName = ss.getName();
  var ask = ui.alert('Make sure email address is in column C on the roster tab. Does column C contain email addresses?', ui.ButtonSet.YES_NO_CANCEL);
  var response = ask.NO;
  Logger.log(response);
  
 /*
  
  if(response=='NO'||'CANCEL'){
    ss.setActiveSheet(sheet);
  }
  */
  
 
  Logger.log('UH');
  //get rid of extra rows
   var row = sheet.getLastRow();
  var last = row+1;
  var max = sheet.getMaxRows();
  var number = max-last;
  for(var j=max; j<last; j++){
    sheet.deleteRows(last, number);
    Logger.log(max);
  }
  try{
    var data = sheet.getDataRange().getValues();//data on the sheet
    Logger.log(data[1][0]);
    var len = data.length;
   
    for(var i=2; i< len; i++){
      var tabName = data[i][0];//first column, tab name
      var email = data[i][2];
      var name = data[i][0]+' ' +data[i][1];//concatenate their name together
     var sheetId = ss.getSheetByName(tabName).getSheetId();
      var sheetUrl = 'https://docs.google.com/spreadsheets/d/'+docId+'/edit#gid='+sheetId;
      try{
        GmailApp.sendEmail(email, 'Link to your tab for '+docName, 'Find your tab located at '+sheetUrl);
      }
      catch(err){
      Logger.log('Uh that didn\'t work');
      }//if email won't go through
      
    }
    //hide the roster back
    ss.getSheetByName('roster').hideSheet();//hide roster
    ui.alert('Link to tab\'s sent');
  }
  catch(err){
    Logger.log('Loop didn\'t work');
    }
  
}
  
function moveFront(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var numSheets = ss.getNumSheets();
  ss.insertSheet('delete', numSheets);
  var lastS = ss.getSheetByName('delete');
  
 
  
  for(var i=2; i<numSheets; i++){
   var last = lastS.getLastRow();
    var newLast = last+1;
    Logger.log(ss.getSheets()[i].getName());
    var tabName = ss.getSheets()[i].getName();
    lastS.getRange(newLast, 1).setValue(tabName);
    lastS.getRange(newLast, 3).setValue(ss.getSheetByName(tabName).getIndex());
    try{
    //figure out the range of this bad boy
    var mCol =  ss.getSheetByName(tabName).getLastColumn();
 var mRow =   ss.getSheetByName(tabName).getLastRow();
  Logger.log(mRow+', '+mCol);
 var a1 = lastS.getRange(1,1,mRow,mCol).getA1Notation();
  Logger.log(a1);
    }
    catch(err){}
   
    var count = '=countA(\''+tabName+'\'!'+a1+')';
   lastS.getRange(newLast,2).setValue(count);
    
    
  }
  
  //sort by count
  lastS.getRange(1, 1,lastS.getLastRow(),3).sort([{column: 2, ascending: false}]);
  lastS.getRange(1, 3,lastS.getLastRow(),3).sort([{column: 3, ascending: true}]);
  
  Logger.log('move the tabs');
  //move the tabs
  var lastRow = lastS.getLastRow();
  Logger.log('Last Row = '+lastRow);
  var data = lastS.getDataRange().getValues();
  for(var j=0; j<lastRow; j++){
   var tabName = data[j][0];
    Logger.log('tabName = '+tabName);
    var k = j+3;
    var m = data[j][2]
    
    var sheet = ss.getSheetByName(tabName);
    Logger.log('Sheet to move '+sheet);
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(m);
//    ss.setActiveSheet(sheet);
    Logger.log('tabName to move '+tabName);
   
   
//    ss.moveActiveSheet(k);
    
  }
  
//var roster = ss.getSheetByName('roster');
//  ss.setActiveSheet(roster);
//  ss.moveActiveSheet(0);
// var template = ss.getSheetByName('template');
//  ss.setActiveSheet(template);
//  ss.moveActiveSheet(1);
  
 ss.deleteSheet(lastS); //delete the delete tab
  

  
//  hideSheets();
  
}

function hideSheets(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
    
 var roster = ss.getSheetByName('roster');
  ss.setActiveSheet(roster);
 ss.moveActiveSheet(0);
   roster.hideSheet();
 var template = ss.getSheetByName('template');
 ss.setActiveSheet(template);
 ss.moveActiveSheet(1);
  template.hideSheet();
 
}
  



function newTemplateTab(){
  
  //1YVnU-4intMuy0_2eISoj0wUVQ17SDvBVKPAVWwY7-lY
  
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = ss.getName();
  
  var template = ss.getSheetByName('template');
  
  
  //copy the TemplateTab and make fresh
  var templateTab = SpreadsheetApp.openById('1YVnU-4intMuy0_2eISoj0wUVQ17SDvBVKPAVWwY7-lY');
  var fresh = templateTab.copy('Fresh '+name);
  var freshUrl = fresh.getUrl();
 var freshId = fresh.getId();
 var freshed = SpreadsheetApp.openById(freshId);
  var templateFresh = freshed.getSheets()[1];
  freshed.deleteSheet(templateFresh);
  
  template.copyTo(freshed);
  freshed.getSheets()[1].showSheet();
 
  var ui = SpreadsheetApp.getUi();
  ui.alert('Fresh TemplateTab created. Check Google Drive or the URL is '+freshUrl);
  
  
}