// StplCodeDeck

// **********************************************
// function fcnRefreshDeck()
//
// 
//
// **********************************************

function fcnRefreshDeck()
{
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var DeckSht = ss.getActiveSheet();
  var DeckName = DeckSht.getSheetName();
  var DeckMaxRow = DeckSht.getMaxRows();
  var CardCntCell;
  var CardCell;
  var OrderNumCell;
  var CardName = '';
  var CardOrderNum = 0;

  var ConfigSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");

  if (DeckName != 'Staple Binder' && DeckName != 'DeckLists' && DeckName != 'Staple CrossRef' && DeckName != 'Config'){
    
    // SAVES THE DECK TAB COLOR
    var DeckColor = DeckSht.getTabColor();
    
    // SETS DECK TAB COLOR TO RED
    DeckSht.setTabColor('red');
    
    // SETS DECK NAME IN FIRST CELL
    DeckSht.getRange(1,1).setValue(DeckName);
    
    // CLEARS DECK STAPLE LIST IN DECKS TAB
    subClearDeckStaple(DeckName);
    
    // CREATES DECK STAPLE LIST IN DECKS TAB
    subCreateDeckStaple(DeckName);
    
    // UPDATES CARDS INFO (
    for(var CardRow = 3; CardRow <= DeckMaxRow; CardRow++){
      // FIND CARD IN CURRENT DECK
      CardCntCell = DeckSht.getRange(CardRow,1);
      CardCell = DeckSht.getRange(CardRow,2);
      OrderNumCell = DeckSht.getRange(CardRow,4);
      CardName = "";
      
      // READS CARD NAME AND TYPE COUNT
      CardName = CardCell.getValue();
      
      if (CardName != '' && CardName != 'End'){

        // HIGHLIGHTS THE CARD AND DECK CELLS
        CardCell.setBackground('orange');
        
        // GETS THE ORDER NUMBER VALUE
        CardOrderNum = subSetCardOrderNumber(CardName);
        OrderNumCell.setValue(CardOrderNum);
      }

      // RESETS THE CARD CELL
      CardCell.setBackground(null);
    }
    // RESETS THE TAB COLOR
    DeckSht.setTabColor(DeckColor);
  }
}


// **********************************************
// function fcnSortSectionStaple()
//
// This function sorts all cards from a section
// according to the Staple Binder organization
//
// **********************************************

function fcnSortSectionStaple()
{  
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();
  var actShtName = actSht.getName();
  var MaxRows = actSht.getMaxRows();
  var MaxCols = actSht.getMaxColumns();
  var NumCols = MaxCols - 1;
  var SectRange;
  var SectFirstRow = actSht.getActiveCell().getRow();
  var SectNumRows;
  var EndValue;

  if (actShtName != 'Staple Binder' && actShtName != 'DeckLists' && actShtName != 'Staple CrossRef' && actShtName != 'Config'){  
    // Finds the End of the Section to determine the number or rows in the section
    for (var Row = SectFirstRow; Row <= MaxRows; Row++){
      EndValue = actSht.getRange(Row, 2).getValue();
      if (EndValue == 'End'){
        SectNumRows = Row - SectFirstRow - 1;
        Row = MaxRows + 1;
      }
    }
    
    SectRange = actSht.getRange(SectFirstRow+1, 2, SectNumRows, NumCols);
    SectRange.sort(4);
  }
}

// **********************************************
// function fcnSortSectionName()
//
// This function sorts all cards from a section
// in Card Name order
//
// **********************************************

function fcnSortSectionName()
{  
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();
  var actShtName = actSht.getName();
  var MaxRows = actSht.getMaxRows();
  var MaxCols = actSht.getMaxColumns();
  var NumCols = MaxCols - 1;
  var SectRange;
  var SectFirstRow = actSht.getActiveCell().getRow();
  var SectNumRows;
  var EndValue;

  if (actShtName != 'Staple Binder' && actShtName != 'DeckLists' && actShtName != 'Staple CrossRef' && actShtName != 'Config'){  
    // Finds the End of the Section to determine the number or rows in the section
    for (var Row = SectFirstRow; Row <= MaxRows; Row++){
      EndValue = actSht.getRange(Row, 2).getValue();
      if (EndValue == 'End'){
        SectNumRows = Row - SectFirstRow - 1;
        Row = MaxRows + 1;
      }
    }
    
    SectRange = actSht.getRange(SectFirstRow+1, 2, SectNumRows, NumCols);
    SectRange.sort(2);
  }
}


// **********************************************
// function fcnAddCardLine()
//
// This function adds a card line to a Deck Sheet 
//
// **********************************************

function fcnAddCardLine()
{
  // Load active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();
  var actShtName = actSht.getSheetName();
  var LastColumn = actSht.getMaxColumns();
  var actRng = ss.getActiveRange();
  var Row = actRng.getRowIndex();
  var NewRow = Row + 1;
  var ColCardCount = actSht.getRange(Row, 1);
  var ColCardName = actSht.getRange(Row,2);
  var ColCardActive = actSht.getRange(Row,3);
  
  if (actShtName != 'Staple Binder' && actShtName != 'DeckLists' && actShtName != 'Config' && Row > 1){
    actSht.insertRowAfter(Row);
    ColCardCount.copyTo(actSht.getRange(NewRow, 1));
    ColCardActive.copyTo(actSht.getRange(NewRow,3));
  }
}


// **********************************************
// function fcnRemoveCard()
//
// This function removes a staple card from the Deck
// and updates the Staple Cross Reference Sheet
// 
// **********************************************

// IN CONSTRUCTION

function fcnRemoveCard()
{
  // Load active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();
  var actShtName = actSht.getSheetName();
  var actRng = ss.getActiveRange();
  var RowNb = actRng.getNumRows();
  var Row = actRng.getRowIndex();

  if (actShtName != 'Staple Binder' && actShtName != 'DeckLists' && actShtName != 'Config' && Row > 1){
    actSht.deleteRows(Row, RowNb);
  }
}

