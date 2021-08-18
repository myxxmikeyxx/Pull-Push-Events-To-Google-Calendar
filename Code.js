// https://stackoverflow.com/questions/42345467/communicate-between-sidebar-and-modal-dialogue-box-in-google-script
// https://stackoverflow.com/questions/9895082/javascript-populate-drop-down-list-with-array
// https://developers.google.com/apps-script/guides/html/templates#index.html_3
// http://jsfiddle.net/yYW89/
function onOpen(e){
  var ui = SpreadsheetApp.getUi()
      ui.createMenu('Calendar Addon')
      .addItem('First Run (Click Me)', 'firstRun')
      .addItem('URL', 'showUrl')
      .addItem('Help URL', 'showHelpUrl')
      .addItem('Dialog', 'showDialog')
      .addItem('SideBar', 'showSidebar')
      .addItem('Popup', 'popUp') 
      .addItem('DropDown User Calendars', 'userCalendarsHTML')
      .addItem('New Show URL', 'showurlNew')
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Get Events From Calendar')
          .addItem('Get Events', 'getEvents')
          .addItem('Example Format', 'getImportExample')
                  .addSeparator()
          .addItem('Delete Extra Sheets', 'DeleteSheets')
          .addItem('Calendar ID Help', 'showUrl'))
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Add Events To Calendar')
          .addItem('Add Events', 'addEvents')
          .addItem('Auto Fill Force', 'autoFill')
          .addItem('Example Format', 'getExportExample')
                  .addSeparator()
          .addItem('Delete Extra Sheets', 'DeleteSheets')
          .addItem('Calendar ID Help', 'showUrl'))
      .addToUi(); 
}

function functionName(){
	alert('Hello world');
}

function DeleteSheets(){
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Delete Sheets', 'Would you like to delete ALL EXTRA sheets and their data? \n This will not Delete Imported Dates or Add Events or the content. \n Do First Run to remove all data. ', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var ssa = SpreadsheetApp.getActiveSpreadsheet();
    var ssall = ssa.getSheets();
    for(var i=0; i < ssall.length; i++){
      Logger.log(ssall[i] + " Sheet name in place " + i); 
      
      var itt = ssa.getSheetByName('Imported Dates');
      var check2 = ssa.getSheetByName('Imported Dates');

      if (!itt) {
        ssa.insertSheet('Imported Dates',1);
      }
      if (!check2) {
        ssa.insertSheet('Add Events',0);
      }
    }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetsCount = ss.getNumSheets();
  var sheets = ss.getSheets();
    for (var i = 0; i < sheetsCount; i++){
      var sheet = sheets[i]; 
      var sheetName = sheet.getName();
      Logger.log(sheetName);
      if (sheetName != "Imported Dates"){
       if  (sheetName != "Add Events"){
        Logger.log("DELETE!" + sheet);
        ss.deleteSheet(sheet);
       }
      }else {
        Logger.log("No sheets to delete");
      }
      
    }
  }else if (response == ui.Button.NO){
    Logger.log("Clicked No For Deleting Extra Sheets.");
  }else {
    Logger.log("Clicked Close X.");
  }
}

function firstRun(){
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('First Run', 'Would you like to clear all sheets and data? \n If you select No nothing will happen.', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    Logger.log("User Clicked YES to clear all sheets.");
    var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var ssa = SpreadsheetApp.getActiveSpreadsheet();
    var ssall = ssa.getSheets();
    var calID = ss.getRange(1, 10).getValue();
    Logger.log(calID);
    
    for(var i=0; i < ssall.length; i++){
      var itt = ssa.getSheetByName('Imported Dates');
      var check2 = ssa.getSheetByName('Imported Dates');
      if (!itt) {
        ssa.insertSheet('Imported Dates',1);
        var msg = Browser.msgBox("Imported dates sheet did not exist.")    
        } else {
            ssa.deleteSheet(ssa.getSheetByName('Imported Dates'));
            ssa.insertSheet('Imported Dates',1);
        }
      if (!check2) {
        ssa.insertSheet('Add Events',0);
        var msg = Browser.msgBox("Add Events sheet did not exist.")      
        }else {
            ssa.deleteSheet(ssa.getSheetByName('Add Events'));
            ssa.insertSheet('Add Events',0);
        }
    }
    formatAddEvents();
    formatImportDates();
    
    var response = ui.alert('Do you know your Calendar ID?' , 'Click No if you do not know your Calendar ID. \n If you know your Calendar ID Click YES.' , ui.ButtonSet.YES_NO);
    // Process the user's response.
    if (response == ui.Button.YES) {
      Logger.log('The user clicked "Yes" That they do know calendar ID');
      calendarID();
    } else if (response == ui.Button.NO){
      Logger.log('The user clicked "No" That they don\'t know calendar ID');
      showHelpUrl();
    } else {
      Logger.log('The user clicked the close button in the dialog\'s title bar.');
      resetBack(calID);
    }
  } else if (response == ui.Button.NO) {
    Logger.log('Didn\'t want to clear the Documnet.');
  } else {
    Logger.log('The user clicked the close button in the dialog\'s title bar.');
  }
}
function resetBack(calID){
  SpreadsheetApp.setActiveSheet(ssa.getSheets()[0]);// send me back to first sheet
  ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ss.getRange(1, 10).setValue(calID);
  SpreadsheetApp.setActiveSheet(ssa.getSheets()[1]);// send me to the second sheet
  ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); 
  ss.getRange(1, 6).setValue(calID);
  SpreadsheetApp.setActiveSheet(ssa.getSheets()[0]);// send me back to first sheet
}

function calendarID(){
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ssa = SpreadsheetApp.getActiveSpreadsheet();
  var response = ui.prompt("Calendar Link", "Link which Calendar ID to use: ", ui.ButtonSet.OK);
  if (response.getResponseText() != "")
  {
  Logger.log("User Calendar ID:" + response.getResponseText());
  SpreadsheetApp.setActiveSheet(ssa.getSheets()[0]);// send me back to first sheet
  ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ss.getRange(1, 10).setValue(response.getResponseText());
  SpreadsheetApp.setActiveSheet(ssa.getSheets()[1]);// send me to the second sheet
  ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); 
  ss.getRange(1, 6).setValue(response.getResponseText());
  SpreadsheetApp.setActiveSheet(ssa.getSheets()[0]);// send me back to first sheet
  ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  }else{
    Logger.log("User Calendar ID:" + response.getResponseText());
    Logger.log("Nothing entered for clandar ID.");
    var msg = Browser.msgBox("Nothing was entered, nothing will change.");
  }
}

function userCalendarsHTML(){ 
    var htmlOutput = HtmlService.createHtmlOutputFromFile('UserCalendars')
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Help Dialog');
}

function getData(){
  Logger.log(getUserCalendars());
  var test = ['test1', 'test2', 'last one']
  return test;
}

function getUserCalendars(){
  var calArray =  CalendarApp.getAllCalendars();
  Logger.log(calArray);
  Logger.log(calArray.getName())
  var calNameArray;
  var calIDArray;
  for (var i = 0; i < calArray.length; i++) {
    calNameArray[i] = calArray[i].getName();
    calIDArray[i] = calArray[i].getId();
  }
  return calNameArray;
}

function getLotsOfThings(){
  var myArray = ['Apple', 'Banana', 'things']
  return myArray;
}

function formatAddEvents(){
  var ssa = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(ssa.getSheets()[0]);// send me back to first sheet
  ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ss.getRange(1, 1).setValue("Event Title").setBackground("LightGrey").setFontWeight("Bold").setHorizontalAlignment("Center");
  ss.getRange(1, 2).setValue("Start Date").setBackground("LightGrey").setFontWeight("Bold").setHorizontalAlignment("Center");
  ss.getRange(1, 3).setValue("Start Time").setBackground("LightGrey").setFontWeight("Bold").setHorizontalAlignment("Center");
  ss.getRange(1, 4).setValue("End Date").setBackground("LightGrey").setFontWeight("Bold").setHorizontalAlignment("Center");
  ss.getRange(1, 5).setValue("End Time").setBackground("LightGrey").setFontWeight("Bold").setHorizontalAlignment("Center");
  ss.getRange(1, 6).setValue("Location").setBackground("LightGrey").setFontWeight("Bold").setHorizontalAlignment("Center");
  ss.getRange(1, 7).setValue("Description").setBackground("LightGrey").setFontWeight("Bold").setHorizontalAlignment("Center");
  ss.getRange(1, 8).setValue(" Auto Fills Start Date Time ").setBackground("LightGrey").setFontWeight("Bold").setHorizontalAlignment("Center");
  ss.getRange(1, 9).setValue(" Auto Fills End Date Time ").setBackground("LightGrey").setFontWeight("Bold").setHorizontalAlignment("Center");
}

function formatImportDates(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssa = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  sheet.autoResizeColumns(7, 3); // Sets the first 6 columns to a width that fits their text.
  SpreadsheetApp.setActiveSheet(ssa.getSheets()[1]);// send me to the second sheet
  ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); 
  ss.getRange(1, 1).setValue("Event Title").setBackground("LightGrey").setFontWeight("Bold").setHorizontalAlignment("Center");
  ss.getRange(1, 2).setValue("Start Date").setBackground("LightGrey").setFontWeight("Bold").setHorizontalAlignment("Center");
  ss.getRange(1, 3).setValue("End Date").setBackground("LightGrey").setFontWeight("Bold").setHorizontalAlignment("Center");
  ss.getRange(1, 4).setValue("Location").setBackground("LightGrey").setFontWeight("Bold").setHorizontalAlignment("Center");
  ss.getRange(1, 5).setValue("Description").setBackground("LightGrey").setFontWeight("Bold").setHorizontalAlignment("Center");
  SpreadsheetApp.setActiveSheet(ssa.getSheets()[0]);// send me back to first sheet
  ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

function showUrl() {
 var htmlOutput = HtmlService
    .createHtmlOutput('Read Obtain Calendar\'s ID from <a href="https://yabdab.zendesk.com/hc/en-us/articles/205945926-Find-Google-Calendar-ID" target="_blank">this site</a> for help!')
    .setWidth(400) //optional
    .setHeight(300); //optional
SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Help Dialog Title');
}

function showHelpUrl() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('Help')
     .setWidth(400) //optional
     .setHeight(300); //optional
 SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Help Dialog');
 }

function popUp() {
  var s='The meeting is at <strong>10:30AM</strong>.'
    + '<br />'
    + '<input type="button" value="OK" onClick="google.script.host.close();" />';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(s), 'PopUp');
}

function showNoChangeDialog() {
  var html = HtmlService.createHtmlOutputFromFile('NoChange')
      .setWidth(400)
      .setHeight(80);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html, 'No change was made.');
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Page')
      .setTitle('My custom sidebar')
      .setWidth(300);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);
}

function getEvents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ssa = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(ssa.getSheets()[1]);// send me back to hard coded Imported Dates sheet
  ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ssas = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (ssas.getRange(1, 6).getValue() != ""){
    var cal = CalendarApp.getCalendarById(ssas.getRange(1, 6).getValue());
  }else{
    var msg = Browser.msgBox("You must have a Calendar ID.");
    return;
  }
  var d = new Date();
  var year = d.getFullYear();
  var month = d.getMonth();
  var day = d.getDate();
  var oneYear = new Date(year + 1, month, day);
  Logger.log(oneYear);
  var events = cal.getEvents( d , oneYear);
  Logger.log(events);
  var lr = ss.getLastRow(); 
  Logger.log(lr);
  if (lr > 1)
  {
    ss.getRange(2, 1, lr-1, 5).clearContent().clearFormat().clear();    
  }
  for(var i = 0;i<events.length;i++){
    var title = events[i].getTitle();
    Logger.log(title);
    var sd = events[i].getStartTime();
    Logger.log(sd);
    Logger.log(new Date(sd).getHours());
    var ed = events[i].getEndTime();
    Logger.log(ed);
    var loc = events[i].getLocation();
    Logger.log(loc);
    var des = events[i].getDescription();
    Logger.log(des);
    ss.getRange(i+2, 1).setValue(title);
    var sdh = new Date(sd).getHours();
    var edh = new Date(ed).getHours();
    if (sdh == 0.0 && edh == 0.0) {
      ss.getRange(i+2, 2).setValue(sd);
      ss.getRange(i+2, 2).setNumberFormat("mm/dd/yyyy");
    }else{
      ss.getRange(i+2, 2).setValue(sd);
      ss.getRange(i+2, 2).setNumberFormat("mm/dd/yyyy h:mm:ss AM/PM");
    }
    if (sdh == 0.0 && edh == 0.0) {
      ss.getRange(i+2, 3).setValue(sd);
      ss.getRange(i+2, 3).setNumberFormat("mm/dd/yyyy");
    }else{
      ss.getRange(i+2, 3).setValue(ed);
      ss.getRange(i+2, 3).setNumberFormat("mm/dd/yyyy h:mm:ss AM/PM");
    }
    ss.getRange(i+2, 4).setValue(loc);
    ss.getRange(i+2, 5).setValue(des);
    }
}

function addEvents(){
  autoFill();
  var ssas = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ssa = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(ssa.getSheets()[0]);// send me back to first sheet, hard coded to be add events.
  ssas = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ssa = SpreadsheetApp.getActiveSpreadsheet();
  if (ssas.getRange(1, 10).getValue() != ""){
    var cal = CalendarApp.getCalendarById(ssas.getRange(1, 10).getValue());
  }else{
    var msg = Browser.msgBox("You must have a Calendar ID.");
  }
  var lr = ssas.getLastRow();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetsCount = ss.getNumSheets();
  var sheets = ss.getSheets();
      if(ss.getActiveSheet().getName() == "Add Events"){
        if (lr <= 1 ){
          //if it has no events added to the list, then it does nothing
        }else{
          var data = ss.getRange("A2:I" + lr).getValues();
          for(var i=0; i<data.length; i++){
            if (data[i][8] == ""){
              var msg = Browser.msgBox("You must have a Start date and time for an event.");
              break;
            }else {
              if (data[i][9] == ""){
              }else{
                var CalData = "[" + data[i][0] + ", " + data[i][7] + ", "  + data[i][8] + ", "  + data[i][5] + ", "  + data[i][6] + "]" ;
                var NewStartDate = new Date(data[i][7]);
                var NewEndDate = new Date(data[i][8]);
                if( NewStartDate.getHours() == NewEndDate.getHours()){ //redunten since I did this with an auto fill formula
                  cal.createEvent( data[i][0], NewStartDate, NewEndDate, {location: data[i][5] , description: data[i][6]} );
                }else{
                  cal.createEvent( data[i][0], NewStartDate, NewEndDate, {location: data[i][5] , description: data[i][6]} );
                }
              }
            }
          }
        }
      }else{
          var msg = Browser.msgBox("Add Events sheet Needs to be the first Sheet. (please move it and try again)");
      }
}


function onEdit(e) {
  // need to edit formulas so they grab the correct information
  Logger.log(e);
  fillDown(e);

  if (e.range.getSheet().getSheetName() == 'Add Events' && e.range.getColumn() == 1) {
    e.range.setHorizontalAlignment("Center");
  }  
  if (e.range.getSheet().getSheetName() == 'Add Events' && e.range.getColumn() == 2) {
    // e.range.setValue(e.value != '' ? e.setNumberFormat('mm/dd/yy') : '');
    // e.range.offset(0,1).setValue(e.value== 'TRUE' ? new Date() : '');
    e.range.setNumberFormat(e.value != '' ? 'm/d/yy' : '');
    e.range.setHorizontalAlignment("Center");
  }
  if (e.range.getSheet().getSheetName() == 'Add Events' && e.range.getColumn() == 3) {
    e.range.setNumberFormat(e.value != '' ? 'hh:mm A/P"M"' : '');
    e.range.setHorizontalAlignment("Center");
  }
  if (e.range.getSheet().getSheetName() == 'Add Events' && e.range.getColumn() == 4) {
    // e.range.setValue(e.value != '' ? e.setNumberFormat('mm/dd/yy') : '');
    // e.range.offset(0,1).setValue(e.value== 'TRUE' ? new Date() : '');
    e.range.setNumberFormat(e.value != '' ? 'm/d/yy' : '');
    e.range.setHorizontalAlignment("Center");
  }
  if (e.range.getSheet().getSheetName() == 'Add Events' && e.range.getColumn() == 5) {
    e.range.setNumberFormat(e.value != '' ? 'hh:mm A/P"M"' : '');
    e.range.setHorizontalAlignment("Center");
  }
  
  var sheetToWatch= 'Add Events',
      columnToWatch = 3, columnToStamp = 8;
  if (e.range.columnStart !== columnToWatch || e.source.getActiveSheet()
    .getName() !== sheetToWatch || !e.value) return;
  e.source.getActiveSheet()
    .getRange(e.range.rowStart, columnToStamp)
  fillDown(e);

  var sheetToWatch= 'Add Events',
      columnToWatch = 2, columnToStamp = 9;
  if (e.range.columnStart !== columnToWatch || e.source.getActiveSheet()
    .getName() !== sheetToWatch || !e.value) return;
  e.source.getActiveSheet()
  .getRange(e.range.rowStart, columnToStamp)
  fillDown(e);  
  
  var sheetToWatch= 'Add Events',
      columnToWatch = 3, columnToStamp = 9;
  if (e.range.columnStart !== columnToWatch || e.source.getActiveSheet()
    .getName() !== sheetToWatch || !e.value) return;
  e.source.getActiveSheet()
  .getRange(e.range.rowStart, columnToStamp)
  fillDown(e);
  
  var sheetToWatch= 'Add Events',
      columnToWatch = 4, columnToStamp = 9;
  if (e.range.columnStart !== columnToWatch || e.source.getActiveSheet()
    .getName() !== sheetToWatch || !e.value) return;
  e.source.getActiveSheet()
  .getRange(e.range.rowStart, columnToStamp)
  fillDown(e);
  
  var sheetToWatch= 'Add Events',
      columnToWatch = 5, columnToStamp = 9;
  if (e.range.columnStart !== columnToWatch || e.source.getActiveSheet()
    .getName() !== sheetToWatch || !e.value) return;
  e.source.getActiveSheet()
  .getRange(e.range.rowStart, columnToStamp)
  fillDown(e);
}

function fillDown(v) {
    
  //maybe use this when its edited
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ss.getRange(2, 2).setNumberFormat("M/dd/yyyy");
  ss.getRange(2, 8).setFormula('=IF(A2 = "", IF(B2 = "", "" , "Must Have Event Title"),IF(C2="", TEXT(B2,"mmm dd, yyy HH:mm:ss") ,TEXT(B2 + C2 ,"mmm dd, yyyy HH:mm:ss")))');
  var lr = v.range.rowStart;
  var filldown = ss.getRange(2, 8, lr-1);
  ss.getRange(2,8).copyTo(filldown);
  ss.getRange(2, 9).setFormula('=IF(A2 = "", IF(B2 = "", "" , "Must Have Event Title") ,IF(D2 = "", IF(B2 = "", "", IF(C2 ="", TEXT(B2 + 1,"mmm dd, yyy HH:mm:ss"),IF(E2 = "", TEXT(B2 + (C2 + 0.1),"mmm dd, yyyy HH:mm:ss"), TEXT(B2 + E2, "mmm dd, yyyy HH:mm:ss") ))), IF( E2 = "", IF(D2 = B2 , TEXT(D2 + 1, "mmm dd, yyy") & " 00:00:00", TEXT(D2 + E2,"mmm dd, yyyy HH:mm:ss") ), TEXT(D2 + E2 ,"mmm dd, yyyy HH:mm:ss")) ))');
  filldown = ss.getRange(2, 9, lr-1);
  ss.getRange(2,9).copyTo(filldown);
}

function autoFill(){
  // Log information about the data validation rule for cell A1.
  var cell = SpreadsheetApp.getActive().getRange('B2');
  var rule = cell.getDataValidation();
  if (rule != null) {
    var criteria = rule.getCriteriaType();
    var args = rule.getCriteriaValues();
    Logger.log('The data validation rule is %s %s', criteria, args);
  } else {
    Logger.log('The cell does not have a data validation rule.');
    Logger.log(cell.getNumberFormat());
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ss.getRange(2, 8).setFormula('=IF(A2 = "", "",IF(C2="", TEXT(B2,"mmm dd, yyy HH:mm:ss") ,TEXT(B2 + C2 ,"mmm dd, yyyy HH:mm:ss")))');
  var lr = ss.getLastRow(); 
  var filldown = ss.getRange(2, 8, lr-1);
  ss.getRange(2,8).copyTo(filldown);
  //orginal -> ss.getRange(2, 9).setFormula('=IF(D2 = "", IF(B2 = "", "",TEXT(B2 + (C2 + 0.0625),"mmm dd, yyyy HH:mm:ss")), TEXT(D2 + E2,"mmm dd, yyyy HH:mm:ss") )');
  ss.getRange(2, 9).setFormula('=IF(A2 = "",  IF(B2 = "", "" , "Must Have Event Title"),IF(D2 = "", IF(B2 = "", "", IF(C2 ="", TEXT(B2 + 1,"mmm dd, yyy HH:mm:ss"),IF(E2 = "", TEXT(B2 + (C2 + 0.1),"mmm dd, yyyy HH:mm:ss"), TEXT(B2 + E2, "mmm dd, yyyy HH:mm:ss") ))), IF( E2 = "", IF(D2 = B2 , TEXT(D2 + 1, "mmm dd, yyy") & " 00:00:00", TEXT(D2 + E2,"mmm dd, yyyy HH:mm:ss") ), TEXT(D2 + E2 ,"mmm dd, yyyy HH:mm:ss")) ))');
  filldown = ss.getRange(2, 9, lr-1);
  ss.getRange(2,9).copyTo(filldown);
}

function getImportExample(){
  
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Format Sheets', 'This will clear the contents of Imported Dates to show output format. \n Are you sure? ', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheets()[1]);
  ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 
  ss.getRange(2, 1).setValue("Test Event").setHorizontalAlignment("Center");
  ss.getRange(2, 2).setValue(new Date()).setHorizontalAlignment("Center");
  var d = new Date();
  var year = d.getFullYear();
  var month = d.getMonth();
  var day = d.getDate();
  var oneDay = new Date(year, month, day + 1);
  ss.getRange(2, 3).setValue(oneDay).setHorizontalAlignment("Center");
  ss.getRange(2, 4).setValue("Chicago, IL, USA").setHorizontalAlignment("Center");
  ss.getRange(2, 5).setValue("Just the Description if one was added.").setHorizontalAlignment("Center");
  ss.getRange(3, 1).setValue("Test Event").setHorizontalAlignment("Center");
  ss.getRange(3, 2).setValue(new Date()).setHorizontalAlignment("Center");
  ss.getRange(3, 3).setValue(new Date()).setHorizontalAlignment("Center");
  ss.getRange(3, 4).setValue("Chicago, IL, USA").setHorizontalAlignment("Center");
  ss.getRange(3, 5).setValue("Just the Description if one was added.").setHorizontalAlignment("Center");
  }
  else {
    Logger.log("Canceled Example Formatting, Import Dates.");
    showNoChangeDialog();
  }
}

function getExportExample(){
  
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Format Sheets', 'This will clear the contents of Add Events to show output format. \n Are you sure? ', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheets()[0]);
  ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 
  ss.getRange(2, 1).setValue("Test Event").setHorizontalAlignment("Center");
  //ss.getRange(2, 2).setValue(new Date()).setHorizontalAlignment("Center");
  var d = new Date();
  var year = d.getFullYear();  
  var month = 0;
  if((d.getMonth() + 1) < 13){
    month = (d.getMonth() + 1);
  }else{
    d.setMonth(0);
    month = (d.getMonth() + 1);
  }
  var day = d.getDate();
  var oneDay = new Date();
  // Add a day
  oneDay.setDate(d.getMonth(), d.getDate() + 1, d.getYear());
  oneDay = oneDay.getDate();
  var twoDay = new Date();
  // Add two days
  twoDay.setDate(d.getMonth(), d.getDate() + 2, d.getYear());
  twoDay = twoDay.getDate();
  var hours = d.getHours();
  var minutes = d.getMinutes();
  ss.getRange(2, 2).setValue(month +"/" + day + "/" + year).setHorizontalAlignment("Center").setNumberFormat('m/d/yy');
  ss.getRange(2, 3).setValue(hours + ":" + minutes).setHorizontalAlignment("Center").setNumberFormat('hh:mm A/P"M"');
  //ss.getRange(2, 3).setValue(oneDay).setHorizontalAlignment("Center");
  ss.getRange(2, 6).setValue("Chicago, IL, USA").setHorizontalAlignment("Center");
  ss.getRange(2, 7).setValue("Just the Description if one was added.").setHorizontalAlignment("Center");
  ss.getRange(2, 8).setFormula('=IF(A2 = "",  IF(B2 = "", "" , "Must Have Event Title"),IF(C2="", TEXT(B2,"mmm dd, yyy HH:mm:ss") ,TEXT(B2 + C2 ,"mmm dd, yyyy HH:mm:ss")))');
  ss.getRange(2, 9).setFormula('=IF(A2 = "",  IF(B2 = "", "" , "Must Have Event Title"),IF(D2 = "", IF(B2 = "", "", IF(C2 ="", TEXT(B2 + 1,"mmm dd, yyy HH:mm:ss"),IF(E2 = "", TEXT(B2 + (C2 + 0.1),"mmm dd, yyyy HH:mm:ss"), TEXT(B2 + E2, "mmm dd, yyyy HH:mm:ss") ))), IF( E2 = "", IF(D2 = B2 , TEXT(D2 + 1, "mmm dd, yyy") & " 00:00:00", TEXT(D2 + E2,"mmm dd, yyyy HH:mm:ss") ), TEXT(D2 + E2 ,"mmm dd, yyyy HH:mm:ss")) ))');
  ss.getRange(3, 1).setValue("Test Event").setHorizontalAlignment("Center");
  ss.getRange(3, 2).setValue(month +"/" + (oneDay) + "/" + year).setHorizontalAlignment("Center").setNumberFormat('m/d/yy');
  ss.getRange(3, 3).setValue(hours + ":" + minutes).setHorizontalAlignment("Center").setNumberFormat('hh:mm A/P"M"');
  ss.getRange(3, 4).setValue(month +"/" + (twoDay) + "/" + year).setHorizontalAlignment("Center").setNumberFormat('m/d/yy');
  ss.getRange(3, 5).setValue(hours + ":" + minutes).setHorizontalAlignment("Center").setNumberFormat('hh:mm A/P"M"');
  ss.getRange(3, 6).setValue("Chicago, IL, USA").setHorizontalAlignment("Center");
  ss.getRange(3, 7).setValue("Just the Description if one was added.").setHorizontalAlignment("Center");
  ss.getRange(3, 8).setFormula('=IF(A3 = "",  IF(B2 = "", "" , "Must Have Event Title"),IF(C3="", TEXT(B3,"mmm dd, yyy HH:mm:ss") ,TEXT(B3 + C3 ,"mmm dd, yyyy HH:mm:ss")))');
  ss.getRange(3, 9).setFormula('=IF(A3 = "",  IF(B2 = "", "" , "Must Have Event Title"),IF(D3 = "", IF(B3 = "", "", IF(C3 ="", TEXT(B3 + 1,"mmm dd, yyy HH:mm:ss"),IF(E3 = "", TEXT(B3 + (C3 + 0.1),"mmm dd, yyyy HH:mm:ss"), TEXT(B3 + E3, "mmm dd, yyyy HH:mm:ss") ))), IF( E3 = "", IF(D3 = B3 , TEXT(D3 + 1, "mmm dd, yyy") & " 00:00:00", TEXT(D3 + E3,"mmm dd, yyyy HH:mm:ss") ), TEXT(D3 + E3 ,"mmm dd, yyyy HH:mm:ss")) ))');
  
  }
  else {
    Logger.log("Canceled Example Formatting, Export Dates.");
    showNoChangeDialog();
  }
}
