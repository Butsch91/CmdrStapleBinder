// StplCodeStapleList

// **********************************************
// function fcnRefreshAllStapleCount()
//
// 
//
// **********************************************

function fcnRefreshAllStapleCount() {
  var StapleSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Staple Binder");
  var MaxRows = StapleSht.getMaxRows();
  var Row;
  var Col;
  var CardCell;
  var CardCount;
  var CardCatTotal;
  var CardName;
  var ColNameType;
  var ColName;
  var ColColor;
  var ColFont;
  
  var ListSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DeckLists");
  var ListMaxRow = ListSht.getMaxRows();
  var ListMaxCol = ListSht.getMaxColumns();
  var ListDeckName;
  var ListDeckStapleNb;
  var ListCardName;
  var ListCardCount;
  
  var RefSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Staple CrossRef");
  var RefMaxCol = RefSht.getMaxColumns();
  var RefRow = 3;
  var RefCardName;
  var RefCardColor;
  var RefDeckList = new Array(12);
  var DeckNames = new Array(12);
  
  var ConfigSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  
  // HIGHLIGHTS THE CURRENT SHEET
  StapleSht.setTabColor('red');
  RefSht.setTabColor('red');
  
  // CLEARS THE STAPLE CROSS REFERENCE SHEET
  RefSht.getRange(2, 2, 160, 14).setValue('');
  
  // UPDATES THE DECK NAMES AND COUNTS IN THE CROSS REFERENCE SHEET
  DeckNames = ListSht.getRange(1, 2, 2, ListMaxCol - 1).getValues();
  RefSht.getRange(1, 5, 2, ListMaxCol - 1).setValues(DeckNames);
  
  // LOOPS THROUGH EVERY CARD IN THE STAPLE BINDER
  
  // COLUMN LOOP
  for (var Col = 1; Col <= 18; Col++) {
    
    // GETS COLUMN NAME AND COLOR
    ColName = StapleSht.getRange(1, Col).getValue();
    ColNameType = typeof ColName;
    ColColor = StapleSht.getRange(1, Col).getBackground();
    ColFont = StapleSht.getRange(1, Col).getFontColor();
  
    // GETS THE AMOUNT OF CARDS IN THE COLUMN
    if(Col > 1) CardCatTotal = StapleSht.getRange(1, Col - 1).getValue();
    
    ConfigSht.getRange(Col, 1).setValue(ColName);
    ConfigSht.getRange(Col, 2).setValue(CardCatTotal);
    
    // ONLY LOOPS IF COLUM NAME IS NOT A NUMBER, NUMBER MEANS COUNT VALUE
    if (typeof ColName != 'number' && ColName != "#" && CardCatTotal > 0) {
      
      RefCardColor = ColName;
     
      // ROW LOOP
      for (var Row = 2; Row <= 49; Row++) {
        
        // CLEARS THE DECKLIST ARRAY
        for (var j = 1; j <= 12; j++) {
          RefDeckList[j] = '';
        }
        
        // GETS THE CARD RANGE TO AVOID CALLING "getRange" 
        CardCell = StapleSht.getRange(Row, Col);
        
        // GETS THE CARD NAME IN THE STAPLE BINDER
        CardName = CardCell.getValue();
        if (CardName != '' && CardName != 'End') {
          
          // HIGHLIGHTS THE ACTUAL CARD CELL
          CardCell.setBackground(ColColor);
          CardCell.setFontColor(ColFont);
          CardCount++;
          
          ListCardName = '';
          ListCardCount = 0;
          
          // LOOPS THROUGH EVERY CARD IN THE DECKLIST TAB  : ListMaxCol
          for (var ListCol = 2; ListCol <= ListMaxCol; ListCol++) {
            
            ListDeckName = ListSht.getRange(1, ListCol).getValue();
            ListDeckStapleNb = ListSht.getRange(2, ListCol).getValue();
            
            for (var ListRow = 3; ListRow <= ListDeckStapleNb + 2; ListRow++) {
              ListCardName = ListSht.getRange(ListRow, ListCol).getValue();
              if (ListCardName == CardName) {
                ListCardCount++;
                RefDeckList[ListCol - 1] = ListDeckName;
              }
            }
          }
          
          // REFRESH THE STAPLE CROSS REFERENCE SHEET
          RefSht.getRange(RefRow, 2).setValue(CardName);
          RefSht.getRange(RefRow, 3).setValue(RefCardColor);
          RefSht.getRange(RefRow, 4).setValue(ListCardCount);
          
          // DISPLAYS ALL DECKS CONTAINING THE CARD
          var DeckNum = 1;
          for (var XCol = 5; XCol < 5 + 11; XCol++) {
            
            RefSht.getRange(RefRow, XCol).setValue(RefDeckList[DeckNum])
            DeckNum++;
          }
          RefRow++;
        }
        // RESETS THE ACTUAL CARD CELL
        CardCell.setBackground(null);
        CardCell.setFontColor(null);
        
        if(CardCount >= CardCatTotal) Row = 50;
      }
    }
  }
  // RESETS THE CURRENT SHEET TAB COLOR
  StapleSht.setTabColor(null);
  RefSht.setTabColor(null);
}


// **********************************************
// function fcnGenAllStapleRef()
//
// This function generates all staple cards references
// It goes through every card in every deck sheet
// and searches for other decks where the card is
// present and list it using the subGenerateFrmDeckRef function.
//
// This is the ultimate UPDATE ALL function
//
// **********************************************

function fcnGenAllStapleRef()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var NbSheets = ss.getNumSheets();
  var FirstDeck = 3;
  var LastDeck = NbSheets - 4;
  var DeckSht = '';
  var DeckName ='';
  var DeckMaxRow = '';
  var CardCntCell;
  var CardCell;
  var OrderNumCell;
  var CardName = '';
  var CardOrderNum = 0;
  
  var ConfigSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
    
  // UPDATES DECKLISTS HEADER
  fcnListHeader();
  
  // GENERATES THE STAPLES IN THE DECKLISTS TAB
  subGenerateDeckListTab();
  
  for (var Tab = FirstDeck; Tab <= LastDeck; Tab++){
    DeckSht = SpreadsheetApp.getActiveSpreadsheet().getSheets()[Tab];
    DeckName = DeckSht.getSheetName();
    DeckMaxRow = DeckSht.getMaxRows();
    
    // SAVES THE DECK TAB COLOR
    var DeckColor = DeckSht.getTabColor();
    
    // HIGHLIGHTS THE CURRENT SHEET
    DeckSht.setTabColor('red');
    
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
    
    // RESETS THE CURRENT SHEET HIGHLIGHT
    DeckSht.setTabColor(DeckColor);
    
    // ConfigSht.getRange(Tab,1).setValue(DeckName);
  }
}