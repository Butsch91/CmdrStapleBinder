// StplOnOpen

// **********************************************
// function OnOpen()
//
// Function executed everytime the Spreadsheet is
// opened or refreshed
//
// **********************************************

function OnOpen() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ConfigSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  var GenDeckList = ConfigSht.getRange(2, 9).getValue();
  var GenStapleRef = ConfigSht.getRange(3, 9).getValue();
  var RefreshStapleCount = ConfigSht.getRange(4, 9).getValue();
  
  // SETS THE "DECKLISTS" LIST TAB HEADER ACCORDING TO THE DIFFERENT TAB NAMES
  fcnListHeader();
  
  // GENERATES DECKLISTS
  if(GenDeckList == 'Enable') subGenerateDeckListTab();
  
  // GENERATES ALL STAPLES REFERENCES TO OTHER DECKS
  if(GenStapleRef == 'Enable') fcnGenerateStapleRef(); 
  
  // UPDATES ALL STAPLES IN STAPLE LIST TAB
  if(RefreshStapleCount == 'Enable') fcnRefreshAllStapleCount();
  
  // CLEARS TRANSFER TAB
  fcnClearTransfer();

  // CREATES THE "SPECIAL FUNCTIONS" MENU
  fcnCreateMenu();
 
}

// **********************************************
// function fcnCreateMenu()
//
// 
//
// **********************************************
function fcnCreateMenu()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var GenlMenuButtons = [{name: "Add New Deck", functionName: "fcnAddNewDeck"}, {name: "Hide Columns", functionName: "fcnHideColumns"}, {name: "Show Columns", functionName: "fcnShowColumns"}];
  var StapleMenuButtons = [{name: "Refresh DeckList Tab Staples", functionName: "subGenerateDeckListTab"}, {name: "Refresh All Staple Counts and Cross References", functionName: "fcnRefreshAllStapleCount"}];
  var DeckMenuButtons = [{name: "Refresh Selected Deck", functionName: "fcnRefreshDeck"}, {name: "Add Card Line Below", functionName: "fcnAddCardLine"}, {name: "Remove Card", functionName: "fcnRemoveCard"}, {name: "Sort Section by Staple Order (Select Section Header)", functionName: "fcnSortSectionStaple"}, {name: "Sort Section by Card Name (Select Section Header)", functionName: "fcnSortSectionName"}];
  
  ss.addMenu("General Fctn", GenlMenuButtons);
  ss.addMenu("Staple Fctn", StapleMenuButtons);
  ss.addMenu("Deck Fctn", DeckMenuButtons);

}

// **********************************************
// function fcnListHeader()
//
// 
//
// **********************************************

function fcnListHeader()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();
  var ListSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DeckLists");
  var MaxCol = ListSht.getMaxColumns();
  var DeckSht = '';
  var Tab = 3;
  var Col = 0;
  var DeckName = '';
  var ConfigSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  
  // HIGHLIGHTS THE CURRENT SHEET
  ListSht.setTabColor('red');
  
  for (Col = 2; Col <= MaxCol; Col++){
    DeckSht = SpreadsheetApp.getActiveSpreadsheet().getSheets()[Tab];
    //DeckSht.setTabColor(null);
    DeckName = DeckSht.getSheetName();
    DeckSht.getRange(1,1).setValue(DeckName);
    ListSht.getRange(1, Col).setValue(DeckName);
    Tab++;
  }
  
  // RESETS THE CURRENT SHEET HIGHLIGHT
  ListSht.setTabColor(null);

}



// **********************************************
// function fcnClearTransfer()
//
// THIS FUNCTION ADDS A NEW DECK TAB BY COPYING
// THE "NEW" TAB AND ADDING IT AFTER THE LAST DECK
// AND CREATING A NEW ROW IN THE "DECKLISTS" TAB
//
// **********************************************

function fcnClearTransfer()
{
  var TrsfrSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transfer");
  var TrsfrMaxRow = TrsfrSht.getMaxRows();
  var TrsfrRow = 8;
  
  // HIGHLIGHTS THE CURRENT SHEET
  TrsfrSht.setTabColor('red');
  
  // Clears both Decks
  TrsfrSht.getRange(2,1).setValue('');
  TrsfrSht.getRange(5,1).setValue('');
  
  // Clears Current Transfer List
  for (var ClrRow = TrsfrRow; ClrRow <= TrsfrMaxRow; ClrRow++){
    TrsfrSht.getRange(ClrRow,1).setValue('');
  }
  
  // RESETS THE CURRENT SHEET HIGHLIGHT
  TrsfrSht.setTabColor(null);
}