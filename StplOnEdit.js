// StplOnEdit

// **********************************************
// function SearchCardNameonEdit()
//
// 
//
// **********************************************

function fcnSearchCardNameonEdit(event)
{
  var actSht = event.source.getActiveSheet();
  var actRng = event.source.getActiveRange();
  var actShtName = actSht.getSheetName(); 

  var Row = actRng.getRowIndex();
  var Col = actRng.getColumn();
  
  var CurrDeckName =actShtName;
  var CardCntCell = actSht.getRange(Row,1);
  var CardCell = actSht.getRange(Row,2);
  var OrderNumCell = actSht.getRange(Row,4);
  var CardName = '';
  var CardFound = 0;
  var CardOrderNum = 0;
  var ConfigSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  
  var NextCardName;
  
  if (actShtName != 'Staple Binder' && actShtName != 'DeckLists' && actShtName != 'Transfer' && actShtName != 'Config' && actShtName != 'Staple CrossRef' && Row > 1 && Col == 2){
    
	// READS CARD NAME AND TYPE COUNT
    CardName = CardCell.getValue();
    
    if (CardName != '' && CardName != 'End'){
      // FIND CARD IN CURRENT DECK
      CardFound = subSearchCardDeck(CurrDeckName,CardName);
      //ConfigSht.getRange(4,3).setValue(CardFound);
      
      // IF CARD IS NOT FOUND (CardFound > 1), ADD IT TO THE DECK SPREADSHEET. CardFound Value = Next Empty Space 
      if (CardFound > 1){
        subAddCardDeck(CurrDeckName,CardName,CardFound);
        
        // GETS THE ORDER NUMBER VALUE
        CardOrderNum = subSetCardOrderNumber(CardName);
        OrderNumCell.setValue(CardOrderNum);
      }         
      // UPDATES THE STAPLE CROSS REFERENCE SHEET
      subUpdateCrossRef(CurrDeckName, CardName);
    }
    
    // IF CARD IS DELETED, CLEAR ALL VALUES IN DECK COLUMN AND REFRESH STAPLE LIST IN DECKS TAB
    if(CardName == '' && CardName != 'End'){
      OrderNumCell.setValue('');
      subClearDeckStaple(CurrDeckName);
      subCreateDeckStaple(CurrDeckName);
    }
  }
}



// **********************************************
// function fcnDeckTransfer()
//
// In the Transfer tab, this function lists all staple cards
// from the first Deck that are also in the second Deck 
// 
// **********************************************

function fcnDeckTransfer(event)
{
  var actSht = event.source.getActiveSheet();
  var actRng = event.source.getActiveRange();
  var actRow = actRng.getRow();
  var actCol = actRng.getColumn();
  var actShtName = actSht.getSheetName();

  var TrsfrSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transfer");
  var SrcDeck = TrsfrSht.getRange(2,1).getValue();  
  var DestDeck = TrsfrSht.getRange(5,1).getValue();
  var TrsfrMaxRow = TrsfrSht.getMaxRows();
  var TrsfrRow = 8;
  
  
  if (actShtName == 'Transfer' && (actRow == 2 || actRow == 5) && actCol == 1 && SrcDeck != '' && DestDeck != '' && SrcDeck != DestDeck){
  
    //var ConfigSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
    var ListSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DeckLists");
    var ListMaxRow = ListSht.getMaxRows();
    var ListMaxCol = ListSht.getMaxColumns();
    var Deck1Row = '';
    var Deck1Col = '';
    var Deck1Card = '';
    var Deck2Row = '';
    var Deck2Col = '';
    var Deck2Card = '';
    var ColName = '';
    
    // Clears Current Transfer List
    for (var ClrRow = TrsfrRow; ClrRow <= TrsfrMaxRow; ClrRow++){
      TrsfrSht.getRange(ClrRow,1).setValue('');
    }
    
    // Finds Columns for Each Deck
    for(var Col = 1; Col <= ListMaxCol; Col++){
      ColName = ListSht.getRange(1,Col).getValue();
      if(SrcDeck == ColName) Deck1Col = Col;
      if(DestDeck == ColName) Deck2Col = Col;
    }
   
    // Browse through each card of Deck 1 in Deck List Sheet
    
    for(var Deck1Row = 3; Deck1Row <= ListMaxRow; Deck1Row++){
      Deck1Card = ListSht.getRange(Deck1Row, Deck1Col).getValue();
      if (Deck1Card != ''){
        for(var Deck2Row = 3; Deck2Row <= ListMaxRow; Deck2Row++){
          Deck2Card = ListSht.getRange(Deck2Row, Deck2Col).getValue();
          if(Deck1Card == Deck2Card){
            TrsfrSht.getRange(TrsfrRow,1).setValue(Deck1Card);
            TrsfrRow++;
          }
        }
      }
    }
  }
  
  else if (SrcDeck == '' || DestDeck == '' || SrcDeck == DestDeck){
    
    // Clears Current Transfer List
    for (var ClrRow = TrsfrRow; ClrRow <= TrsfrMaxRow; ClrRow++){
      TrsfrSht.getRange(ClrRow,1).setValue('');
    }
  }
}


// **********************************************
// function fcnGathererLink()
//
// Creates a link to the Gatherer Card Page when
// a new card is entered in the Staple List
// 
// **********************************************

function fcnGathererLink() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();								
  var actShtName = actSht.getSheetName();						
  var actRng = actSht.getActiveRange();	
  var CardName = actRng.getValue();
  var RowIndex = actRng.getRowIndex();
  var ColIndex = actRng.getColumn();
  var ColName = actSht.getRange(1,ColIndex).getValue();
  //var ConfigSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  
  if(actShtName == 'Staple Binder' && typeof ColName != 'number' && ColName != "#" && CardName != "" && ColIndex >= 1 && ColIndex <= 18 && RowIndex > 1) { 
    actRng.setValue('=HYPERLINK("http://gatherer.wizards.com/Pages/Card/Details.aspx?name='+CardName+'","'+CardName+'")');
    actRng.setFontLine("none");
  }
}

// **********************************************
// function fcnDeckFunctions()
//
// When the Deck Function cell is modified,
// the appropriate function is executed
// 
// **********************************************

function fcnDeckFunctions(event) {
  
  var actSht = event.source.getActiveSheet();
  var actRng = event.source.getActiveRange();
  var actShtName = actSht.getSheetName(); 
  var Row = actRng.getRowIndex();
  var Col = actRng.getColumn();
  var SortCmd;
  
  if(actShtName != 'Staple Binder' && actShtName != 'DeckLists' && actShtName != 'Transfer' && actShtName != 'Config' && actShtName != 'Staple CrossRef' && actShtName != 'New' && Row == 2 && Col == 5){
    SortCmd = actRng.getValue();
    
    if (SortCmd == 'Refresh Deck'){
      fcnRefreshDeck();
      actRng.setValue("Select Function");
    }

  }
}

