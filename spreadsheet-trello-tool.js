// -----------------------------------------------------------------------------
// SpreadSheet Trello Tool
// plugin for google drive to dump a trello board
// -----------------------------------------------------------------------------

var appKey = "";
var token = "";
var boardName = "";
var boardId = "";

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Trello')
    .addItem('Dump Board', 'dumpBoard')
    .addSeparator()
    .addSubMenu(ui.createMenu('Configuration')
      .addItem('Create Config Sheet', 'createConfigSheet'))
    .addToUi();
}

function dumpBoard() {
  Browser.msgBox("Start to dump trello board");
  var ss = SpreadsheetApp.getActive().getSheetByName("Config");
  var values = ss.getRange("B1:B4").getValues();

  boardName = values[0];
  boardId = values[1];
  appKey = values[2];
  token = values[3];

  backup(boardName,boardId);
}

function backup(boardName,boardId) {
  var allData = "";
  allData = getBackupData(boardId,"?cards=all&actions=commentCard&actions_limit=1000&card_attachments=true&lists=all&card_checklists=all&members=all");

  // Parse the data:
  var cards = JSON.parse(allData).cards;
  var lists = JSON.parse(allData).lists;
  var actions = JSON.parse(allData).actions;
  var members = JSON.parse(allData).members;

  populateBoardSheet(boardId,boardName,cards,lists,actions,members);
}

function getBackupData(boardID,data) {
  var url = constructTrelloURL("boards/" + boardID + data);
  var resp = UrlFetchApp.fetch(url, {"method": "get"});
  return resp.getContentText();
}

function constructTrelloURL(baseURL){
  if (baseURL.indexOf("?") == -1) {
    return "https://trello.com/1/"+ baseURL +"?key="+appKey+"&token="+token;
  } else {
    return "https://trello.com/1/"+ baseURL +"&key="+appKey+"&token="+token;
  }
}

function populateBoardSheet(boardID,boardName,cards,lists,actions,members) {
  try {
    var board = createBoardBackupSheet(boardName);
    var headings = board.getRange(1,1,1,headingCount).getValues()[0];
    for (var i = 0; i < headings.length; i++) {
      headings[i] = getHeadingKey(headings[i]);
    }

    var listIds = new Array();
    var checklistIds = new Array();
    var actionCardIds = new Array();
    var allRows = new Array();

    // Configure lookup arrays;
    for (var i=0; i<lists.length; i++) {listIds.push(lists[i].id);}
    for (i=0;i<actions.length;i++) {actionCardIds.push(actions[i].data.card.id);}

    var row = 1;
    for (i = 0; i < cards.length;i++) {

      var card = cards[i];
      for (var propertyName in card) {
        var lowerName = propertyName.toLowerCase();
        if (card.hasOwnProperty(propertyName) && propertyName != lowerName) {
          card[lowerName] = card[propertyName];
        }
      }

      var name = "";
      var storyPoints = "";
      var acceptanceCriteria = "";
      var otherChecklists = "";
      var dueDate = "";
      if (card.due != null) {
        dueDate = card.due;
      }

      nap = parseCardName(isScrum,card.name);
      card = parseCardChecklists(card,headings);
      comments = parseActionComments(card,actions,actionCardIds);
      var list = parseCardList(card,lists,listIds);
      var cardStatus = "open";

      if (card.closed) {
        cardStatus = "closed";
      }

      if (!openOnly || (cardStatus == "open" && list.status == "open")) {
        Logger.log("Card %s has status %s", card.idShort,cardStatus);

        card.duedate = dueDate;
        card.title = nap.name;
        card.description = card.desc;
        card.userstory = card.desc;
        card.cost = nap.cost;
        card.created = new Date(1000*parseInt(card.id.substring(0,8),16))

        card.storypoints = nap.points;
        card.consumedpoints = nap.consumedPoints;
        card.cardstatus = cardStatus;
        card.list = list.name;
        card.liststatus = list.status;

        card.checklists = card.other;
        card.other = card.checklists;
        card.labels = parseLabels(card.labels);
        card.attachments = parseAttachments(card.attachments);
        card.actions = parseActionComments(card,actions,actionCardIds);
        card.members = parseMembers(card,members);
        // This must be done last
        card.id = card.idShort;

        if (canProcessList(card.list, listFilter) && canProcessCard(card,columnFilter,headings)) {
          row++;
          var rowData = getDataForRow(card,headings);
          allRows.push(rowData[0]);
        }
      }
    }

    if (row > 1) {
      board.getRange(2,1,allRows.length,headings.length).setValues(allRows).setVerticalAlignment("top");
      boardDataFound = true;
    }

    Browser.msgBox('Dump ready');
  } catch (e) {
    Browser.msgBox("Unable to dump Trello backup data. Consider not dumping data, or backing up only open cards. Error was: " + e.message);
  }
}

function createBoardBackupSheet(boardName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheets = ss.getSheets();
  var sheetNo = sheets.length;
  var reusedSheet = false;
  var newName = sheetNo + " - " + boardName;
  var newSheet = null;
  var headers = ss.getSheetByName("Config").getRange("B11:T11");
  headingCount = headers.getNumColumns();

  for (var i=0;i<sheets.length;i++) {
   var nm = sheets[i].getName();
   if (nm.substr(0,nm.indexOf(" ")) == sheetNo.toFixed(0)) {
     newSheet = ss.getSheetByName(nm);
     var dataArea;
     if (newName !== nm) {
       newSheet.setName(newName);
       dataArea = newSheet.getDataRange();
     }
     else {
       var existingData = newSheet.getDataRange();

       dataArea = newSheet.getRange(headers.getRow(),headers.getColumn(), existingData.getNumRows(), headers.getNumColumns());
     }

     dataArea.clearContent();

     break;
   }
  }

  if (newSheet === null) {
   newSheet = ss.insertSheet(newName,sheetNo-1);
  }

  headers.copyTo(newSheet.getRange(1,1,1,headers.getNumColumns()));

  newSheet.setFrozenRows(1);
  return newSheet;
}

function createConfigSheet() {
  var ss = SpreadsheetApp.getActive();
  var configSheet = ss.getSheetByName("Config");
  if(configSheet == null) {
    fillConfigSheet();
  }else{
    Browser.msgBox("Ya existe la hoja");
  }
}

function fillConfigSheet() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.insertSheet("Config");
  sheet.getRange("A1").setValue("board name");
  sheet.getRange("A2").setValue("board id");
  sheet.getRange("A3").setValue("app token");
  sheet.getRange("A4").setValue("app key");

  sheet.getRange("A7").setValue("get trello app key and token");
  sheet.getRange("A10").setValue("available columns");
    sheet.getRange("B10").setValue("Id");
    sheet.getRange("C10").setValue("Title");
    sheet.getRange("D10").setValue("Story Points");
    sheet.getRange("E10").setValue("Cost");
    sheet.getRange("F10").setValue("Consumed Points");
    sheet.getRange("G10").setValue("User Story");
    sheet.getRange("H10").setValue("Acceptance Criteria");
    sheet.getRange("I10").setValue("Card Status");
    sheet.getRange("J10").setValue("List");
    sheet.getRange("K10").setValue("List Status");
    sheet.getRange("L10").setValue("Checklists");
    sheet.getRange("M10").setValue("Labels");
    sheet.getRange("N10").setValue("Attachments");
    sheet.getRange("O10").setValue("Actions");
    sheet.getRange("P10").setValue("Members");
    sheet.getRange("Q10").setValue("URL");
    sheet.getRange("R10").setValue("Due Date");
    sheet.getRange("S10").setValue("Questions");
    sheet.getRange("T10").setValue("Created");
  sheet.getRange("A11").setValue("selected columns * copy and paste the columns");

  // Assign background color
  var cells = ["A1:A4","A7","A10","A11"]
  var i=0,
      arryLngth = cells.length;
  for (i=0; i < arryLngth; i+=1) {
    sheet.getRange(cells[i]).setBackground("#CCCCCC");
  };
  sheet.getRange("B10:T10").setBackground("#CCCCCC");
}
