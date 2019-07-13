/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */

var TZ = "G1";
var settingsSheet = "Settings";
var inputSheet = "Events";
var vaildCol = 16; // Q is 17th letter
var idCol = 15; // P is 16th letter
var startCol = 13;
var endCol = 14;
var nameCol = 5;
var notesCol = 6;
var statusCol = 9; 
var locCol = 7;
var emailCol = 8;

var addedMsg = "Added";

function onOpen(e) {
    
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var callist = ss.getSheetByName(settingsSheet);
  var timezone = callist.getRange(TZ).getValue().toString();
  ss.setSpreadsheetTimeZone(timezone);
  var menus = ui.createMenu('Event Adder');
      menus.addItem('Update Calendar List', 'getCals')
      menus.addItem('Add Events', 'addCal')
      menus.addItem('Show logs', 'showLog')
      menus.addToUi();
  
    Logger.log("onOpen Done");
}

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}

function getCals() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var callist = ss.getSheetByName(settingsSheet);
   var calendars = CalendarApp.getAllCalendars();
       for(var i = 0; i < 1000 ; i++){
    callist.getRange(2+i,1).setValue("");
    callist.getRange(2+i,2).setValue("")     
  }
   for(var i = 0; i < calendars.length ; i++){
   var calname = calendars[i].getName();
    var calid = calendars[i].getId();
    callist.getRange(2+i,1).setValue(calname);
    callist.getRange(2+i,2).setValue(calid)     
  }
  createCalPicker();
  Logger.log("Cals refreshed");

}

function addCal(){
  Logger.clear();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var eventsl = ss.getSheetByName(inputSheet);
  
  var lastColumn = ss.getLastColumn();
  var eventList = ss.getRange("A:Q").getValues();

  for(var rowNum = 1; rowNum < 10; rowNum++){
    var eventRow = eventList[rowNum];
    if(eventRow[vaildCol] == "Vaild" && eventRow[statusCol] != addedMsg){
      Logger.log("Event: " + eventRow)
      var eventCal = CalendarApp.getCalendarById(eventRow[idCol]);
      if(eventCal == null){
        Logger.log("CalApp fail");
        continue;
      }
      var notes = {};
      if(eventRow[notesCol]){
      notes.description = eventRow[notesCol];
      }
      var newEvent = eventCal.createEvent(eventRow[nameCol],new Date(eventRow[startCol]), new Date(eventRow[endCol]), notes);
      if(newEvent == null){
        Logger.log("Event fail");
        continue;
      }
      if(eventRow[locCol])
        newEvent.setLocation(eventRow[locCol]);
      if(!isNaN(eventRow[emailCol]) && eventRow[emailCol]){
        var minutes = parseInt(eventRow[emailCol] * 24 * 60) + 1;
        newEvent.addEmailReminder(minutes);
      }
        
      
      eventsl.getRange(rowNum+1,statusCol+1).setValue(addedMsg);

    }
    
  }}
    

function test(){
Logger.log(CalendarApp.getCalendarById("Sss"))
}



function createCalPicker(){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = ss.getSheetByName(settingsSheet);
  var callist = ss.getSheetByName(inputSheet);
  var picker = callist.getRange('A2:A1000');
  var range = settings.getRange('A2:A100');
var rule = SpreadsheetApp.newDataValidation().requireValueInRange(range, true).build();
picker.setDataValidation(rule);
  
  picker = callist.getRange('H2:H1000');
    range = settings.getRange(settingsSheet +'!G3:G13');
  
  //rule = SpreadsheetApp.newDataValidation().requireValueInRange(range, true).build();
//picker.setDataValidation(rule);
  
}

function showLog() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(Logger.getLog());

}
