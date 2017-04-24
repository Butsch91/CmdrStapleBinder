// StplSubroutines

// **********************************************
// function subClearDeckStaple()
//
// 
//
// **********************************************

function subClearDeckStaple(DeckName)
{
  var ListSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DeckLists");
  var ListMaxRow = ListSht.getMaxRows();
  var ListMaxCol = ListSht.getMaxColumns();
  var ListColName = '';

  for (var Col=1; Col <= ListMaxCol; Col++){
    ListColName = ListSht.getRange(1,Col).getValue();
    if (DeckName == ListColName){
      ListSht.getRange(3,Col,40).setValue('');
    }
  }
}

// **********************************************
// function subGenerateDeckListTab()
//
// This function browses through each deck tab and
// updates the Staples in the DeckLists tab
//
// **********************************************

function subGenerateDeckListTab()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var NbSheets = ss.getNumSheets();
  var LastDeck = NbSheets - 3;
  var DeckSht = '';
  var DeckName ='';
  var ConfigSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  
  for (var Tab = 3; Tab <= LastDeck; Tab++){
    
    DeckSht = SpreadsheetApp.getActiveSpreadsheet().getSheets()[Tab];
    DeckName = DeckSht.getSheetName();
    
    // SAVES THE DECK TAB COLOR
    var DeckColor = DeckSht.getTabColor();
    
    // HIGHLIGHTS THE CURRENT SHEET
    DeckSht.setTabColor('red');
    
    // CLEARS THE DECK STAPLE LIST IN THE DECKLISTS TAB
    subClearDeckStaple(DeckName);

    // UPDATES THE DECKLISTS TAB ACCORDING TO THE CURRENT STAPLE LIST IN THE DECK TAB    
    subCreateDeckStaple(DeckName);
    
    // RESETS THE CURRENT SHEET HIGHLIGHT
    DeckSht.setTabColor(DeckColor);
    
    //ConfigSht.getRange(Tab,1).setValue(DeckName);
  }
}

// **********************************************
// function subCreateDeckStaple()
//
// 
//
// **********************************************

function subCreateDeckStaple(DeckName)
{
  var ConfigSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  var ListSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DeckLists");
  var ListMaxRow = ListSht.getMaxRows();
  var ListMaxCol = ListSht.getMaxColumns();
  var ListColName = '';
   
  var DeckSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DeckName);
  var DeckCardName = '';
  var DeckPrevCardCount = '';
  var DeckMaxRow = DeckSht.getMaxRows();
  var ListRow = 3;
 
  // HIGHLIGHTS THE CURRENT SHEET
  ListSht.setTabColor('red');
  
  // LOOPS TO FIND APPROPRIATE COLUMN IN DECKLIST SHEET
  for (var ListCol=1; ListCol <= ListMaxCol; ListCol++){
    ListColName = ListSht.getRange(1,ListCol).getValue();
    
    // FINDS APPRIOPRIATE COLUMN
    if (DeckName == ListColName){
      
      // LOOPS TO GET CARD NAME IN DECK SHEET TO PUT IT IN DECK LIST 
      for(var DeckRow = 3; DeckRow <= DeckMaxRow; DeckRow++){
        
        // FIND CARD IN CURRENT DECK, THEN READS CARD NAME AND CARD TYPE COUNT
        DeckCardName = DeckSht.getRange(DeckRow,2).getValue();
        
        // FINDS CARD NAME IN DECK SHEET
        if (DeckCardName != '' && DeckCardName != 'End'){

          // ADDS CARD NAME TO GLOBAL STAPLE LIST SHEET
          ListSht.getRange(ListRow,ListCol).setValue(DeckCardName);
          ListRow++;
          ListSht.getRange(2,ListCol).setValue(ListRow-3);
        } 
      }
    }
  }
  
  // RESETS THE CURRENT SHEET HIGHLIGHT
  ListSht.setTabColor(null);
}


// **********************************************
// function subSearchCardDeck()
//
// 
//
// **********************************************

function subSearchCardDeck(DeckName,CardName)
{
  var ListSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DeckLists");
  var MaxRow = ListSht.getMaxRows();
  var MaxCol = ListSht.getMaxColumns();
  var RowName = '';
  var ColName = '';

  for (var Col=1; Col <= MaxCol; Col++){
    ColName = ListSht.getRange(1,Col).getValue();
    if (DeckName == ColName){
      for (var Row=3; Row <= MaxRow; Row++){
        RowName = ListSht.getRange(Row,Col).getValue();
        if (CardName == RowName){
          return 1;
        }
        else{
          if(RowName == ''){
            return Row;
          }
        }
      }
    }
  }
}


// **********************************************
// function subAddCardDeck()
//
// 
//
// **********************************************

function subAddCardDeck(DeckName,CardName,NextEmpty)
{
  var ListSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DeckLists");
  var MaxCol = ListSht.getMaxColumns();
  var ColName = '';
  var DeckStapleCnt = '';
  
  for (var Col=1; Col <= MaxCol; Col++){
    ColName = ListSht.getRange(1,Col).getValue();
    if (DeckName == ColName){
      DeckStapleCnt = ListSht.getRange(2,Col).getValue();
      ListSht.getRange(2,Col).setValue(DeckStapleCnt+1);
      ListSht.getRange(NextEmpty,Col).setValue(CardName);
    }
  }
}

// **********************************************
// function subSetCardOrderNumber()
//
// When a new card is added, it updates the Order Number
// according to the value in the Cross-Reference List
// 
// **********************************************

function subSetCardOrderNumber(CardName) {
  
  var RefSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Staple CrossRef");
  var RefMaxRows = RefSht.getMaxRows();
  var Row = 0;
  var Col = 0;
  var RefCardName;
  var CardOrderNum = 0;
  
  for (Row = 2; Row <= RefMaxRows; Row++){
    RefCardName = RefSht.getRange(Row,2).getValue();
    if (RefCardName == CardName) CardOrderNum = RefSht.getRange(Row,1).getValue();
    if (CardOrderNum != 0) Row = RefMaxRows+1;
  }
  return CardOrderNum;
  
}


// **********************************************
// function subUpdateCrossRef()
//
// This function updates the Staple Cross Reference Tab
// adding the Deck Name to the Card sent in parameters
//
// **********************************************

function subUpdateCrossRef(DeckName, CardName)
{
  var RefSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Staple CrossRef");
  var RefMaxRows = RefSht.getMaxRows();
  var RefMaxCols = RefSht.getMaxColumns();
  var RefRow = 2;
  var RefCol = 5;
  var RefCardName;
  var RefDeckName;
  var RefCountCell;
  var RefCountVal;
  var RefDeckList;
  
  for (RefRow = 2; RefRow <= RefMaxRows; RefRow++){
    RefCardName = RefSht.getRange(RefRow,2).getValue();
    if (RefCardName == CardName){
      for (RefCol = 5; RefCol <= RefMaxCols; RefCol++){
        RefDeckName = RefSht.getRange(1, RefCol).getValue();
        if (RefDeckName == DeckName){
          RefSht.getRange(RefRow,RefCol).setValue(DeckName);
          RefCountCell = RefSht.getRange(RefRow,4);
          RefCountVal = RefCountCell.getValue();
          RefCountCell.setValue(RefCountVal+1);
          RefCol = RefMaxCols+1;
        }
      }
    }
  }
}
