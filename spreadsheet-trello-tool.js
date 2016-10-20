// -----------------------------------------------------------------------------
// SpreadSheet Trello Tool
// plugin for google drive to dump a trello board
// -----------------------------------------------------------------------------

function onOpen(){
  var ss = SpreadsheetApp.getActive();
  var menuEntries = [{name: "My custom function", functionName: "myFunction"}];
  ss.addMenu("Trello", menuEntries);
}

function myFunction() {
  var msg = "Hello SpreadsheetApp";
  Browser.msgBox(msg);
  SpreadsheetApp.getActive().getActiveCell().setValue(msg);
}
