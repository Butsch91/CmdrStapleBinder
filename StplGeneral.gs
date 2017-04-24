// StplGeneral

// **********************************************
// function fcnHideColumns()
//
// Hide all columns with the "hide" title
// 
// 
// **********************************************

function fcnHideColumns() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();								
  var actShtName = actSht.getSheetName();
  var MaxRow = actSht.getMaxRows();
  var MaxCol = actSht.getMaxColumns();
  var ColTitle;
  
  if(actShtName != 'Config' && actShtName != 'Conversion' && actShtName != 'Test') { 
    for (var Col = 1; Col <= MaxCol; Col++){
      ColTitle = actSht.getRange(1, Col).getValue();
      if(ColTitle == 'hide') {
        var ColRng = actSht.getRange(1, Col, MaxRow, 1);
        
        actSht.hideColumn(ColRng);
      }
    }
  }
}

// **********************************************
// function fcnShowColumns()
//
// Hide all columns with the "hide" title
// 
// 
// **********************************************

function fcnShowColumns() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();								
  var actShtName = actSht.getSheetName();
  var MaxRow = actSht.getMaxRows();
  var MaxCol = actSht.getMaxColumns();
  var ColTitle;
  
  if(actShtName != 'Config' && actShtName != 'Conversion' && actShtName != 'Test') { 
    for (var Col = 1; Col <= MaxCol; Col++){
      ColTitle = actSht.getRange(1, Col).getValue();
      if(ColTitle == 'hide') {
        var ColRng = actSht.getRange(1, Col, MaxRow, 1);
        actSht.unhideColumn(ColRng);
      }
    }
  }
}

// **********************************************
// function fcnAddNewDeck()
//
// THIS FUNCTION ADDS A NEW DECK TAB BY COPYING
// THE "NEW" TAB AND ADDING IT AFTER THE LAST DECK
// AND CREATING A NEW COLUMN IN THE "DECKLISTS" TAB
//
// **********************************************

function fcnAddNewDeck()
{
  var ConfigSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var NbSheets = ss.getNumSheets();
  var LastDeck = NbSheets - 3;
  var NewSht = ss.getSheetByName("New");
  var ListSht = ss.getSheetByName("DeckLists");
  var ListShtName = ListSht.getSheetName();
  var ListMaxCol = ListSht.getMaxColumns();
  
    // OPENS PROMPT TO NAME NEW DECK AND RENAME INSERTED TAB
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter New Deck Name :');
  
  if (response.getResponseText() !='' && response.getSelectedButton() == ui.Button.OK){
    
    var DeckName = response.getResponseText();
    // INSERTS "NEW DECK" TAB BEFORE "DECKLISTS" TAB
    ss.insertSheet(DeckName, LastDeck, {template: NewSht});
  
    // INSERTS COLUMN "NAME" AT THE END OF "DECKLISTS" TAB
    ListSht.insertColumnAfter(ListMaxCol);
    ListSht.getRange(1, ListMaxCol+1).setValue(DeckName);
  }
}