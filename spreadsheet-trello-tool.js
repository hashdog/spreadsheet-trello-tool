// -----------------------------------------------------------------------------
// SpreadSheet Trello Tool
// plugin for google drive to dump a trello board
// -----------------------------------------------------------------------------

var appKey = "";
var token = "";
var boardName = "";
var boardId = "";

function onOpen(){
  var ss = SpreadsheetApp.getActive();
  var menuEntries = [{name: "Dump Board", functionName: "dumpBoard"}];
  ss.addMenu("Trello", menuEntries);
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

      nap = parseCardName(card.name);
      card = parseCardChecklists(card,headings);
      comments = parseActionComments(card,actions,actionCardIds);
      var list = parseCardList(card,lists,listIds);
      var cardStatus = "open";

      if (card.closed) {
        cardStatus = "closed";
      }

      if (cardStatus == "open" && list.status == "open") {
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

        row++;
        var rowData = getDataForRow(card,headings);
        allRows.push(rowData[0]);
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

function getDataForRow(card,headings) {
  var rowData = [];

  for (var i = 0;i < headings.length;i++) {
    if (card.hasOwnProperty(headings[i].toLowerCase())) {
      rowData.push(card[headings[i].toLowerCase()]);
    }
    else {
      rowData.push("");
    }
  }

  return [rowData];
}

function getHeadingKey(heading) {
  return heading.toLowerCase().replace(/ /g,'').trim();
}

function parseMembers(card, members) {
  var mbr = ""

  for (var a=0;card.idMembers && a<card.idMembers.length; a++) {
    if (a>0) {mbr=mbr+"\n";}
    for (var b=0;b<members.length;b++) {
      if (members[b].id == card.idMembers[a]) {
        mbr = mbr + members[b].fullName;
        break;
      }
    }

  }

  return mbr;
}

function parseLabels(labels) {
  var lb = ""
  for (var l =0;l<labels.length; l++) {
    if (l>0) {lb=lb+"\n";}
    if (labels[l].name.length <= 0) {
      lb = lb + labels[l].color;
    }
    else {
      lb = lb + labels[l].name;
    }
  }

  return lb;
}

function parseAttachments(attachments) {
  var at = ""

  for (var a =0;a<attachments.length; a++) {
    if (a>0) {at=at+"\n";}
    at = at + attachments[a].url;
  }

  return at;
}

function parseCardName(name) {
  var nap = {name:"",points:"",cost:"",consumedPoints:""};

  nap.name = name;
  if (name.charAt(0) == "(" && name.indexOf(")") != -1) {
     nap.points = name.substr(1,name.indexOf(")")-1);
     nap.name = name.substr(name.indexOf(")")+1);
  }

  if (nap.name.indexOf("[") != -1 && nap.name.indexOf("]") != -1 && (nap.name.indexOf("[") < nap.name.indexOf("]")-1)) {

    nap.consumedPoints = nap.name.substring(nap.name.indexOf("[")+1, nap.name.indexOf("]"));

    nap.name = nap.name.substring(0,nap.name.indexOf("[")) + nap.name.substr(nap.name.indexOf("]")+1);
  }

  nap.cost = getCostFromName(nap.name);
  nap.name = getNameWithoutCost(nap.name);

  return nap;
}

function parseCardList(card,lists,listIds) {
  var x = listIds.indexOf(card.idList)
  var list = {name:"",status:""};
  list.name = lists[x].name;
  if (lists[x].closed) { list.status = "closed";} else {list.status = "open"};

  return list;
}

function parseCardChecklists(card,headings) {
  var newCard = card;
  newCard.other = "";

  for (var k=0;k < card.checklists.length; k++) {
    var cl = card.checklists[k];

    for (var j=0; cl.checkItems && j < cl.checkItems.length; j++) {

      var cli = cl.checkItems[j];

      if (cli.state == "incomplete") {

        var checkListItemState = getCheckListItemState(cli.state);

        var name = getHeadingKey(cl.name);

        if (headings.indexOf(name) >= 0) {
          if (!newCard.hasOwnProperty(name)) {
            newCard[name] ="";
          }
          newCard[name] += checkListItemState + cli.name + "\n\n";
        }
        else {
          if (j==0) {
            newCard.other += "**********\n" + cl.name + "\n**********\n"
          }
          newCard.other += checkListItemState + cli.name + "\n\n";
        }
      }
    }
  }

  return newCard;
}

function getCheckListItemState(state) {
  if (state != "incomplete") {
    return "\u2611 ";
  } else {
    return "\u2610 ";
  }
 }

function parseActionComments(card,actions,actionCardIds) {
  var comments = "";

  var i= actionCardIds.indexOf(card.id);
  while (i!=-1) {

    if (comments!="") {comments = comments + "\n\n";}
    var name = "** Member Name Unavailable **";
    if (actions[i].memberCreator && actions[i].memberCreator.fullName) {
      name = actions[i].memberCreator.fullName;
    }
    comments = comments + name + ":\n----------\n" + actions[i].data.text + "\n----------";

    if (i==actionCardIds.length-1) {
      i= -1;
    }
    else {
      i= actionCardIds.indexOf(card.id,i+1);
    }
  }

  return comments;
}

function getNameWithoutCost(name) {
  var nameWithoutCost = name.trim();

  if (name.length > 0 && name.substr(name.length-1) == ")" && name.indexOf("(Cost: ") != -1) {
    nameWithoutCost = name.substr(0,name.indexOf("(Cost: ")).trim();
  }

  return nameWithoutCost;
}

function getCostFromName(name) {
  var cost ="";

  if (name.length > 0 && name.substr(name.length-1) == ")" && name.indexOf("(Cost: ") != -1) {
    cost = name.substring(name.indexOf("(Cost: ")+7,name.length-1)+"";
  }

  return cost;
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
