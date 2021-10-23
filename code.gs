function onOpen() {
  var ui = SpreadsheetApp.getUi();


  ui.createMenu('TPE Options')
    .addItem('Move Completed Records', 'batchMoveCompleted')
    .addSeparator()
    .addItem('Clear logs', 'clearLogs')
    .addItem('Activate Form Submit', 'activateSubmit')
    .addItem('DEV Delete all Tabs', 'deleteTabs')
    .addToUi();

  activateSubmit();
  /*
    let ss = SpreadsheetApp.getActiveSpreadsheet(); //Utils.openSpreadSheet(this.FileUrl);   
    let lTab = Utils.createLogTab(ss);
    let sheet = ss.getSheets()[0];    //.getSheetByName(mainSheetName);
    ScriptApp.newTrigger("onFormSubmitProcessor")
      .forSpreadsheet(sheet)
      .onFormSubmit()
      .create();
  
      */

}


function activateSubmit()
{
  let ss = SpreadsheetApp.getActiveSpreadsheet(); //Utils.openSpreadSheet(this.FileUrl);   
    let lTab = Utils.createLogTab(ss);
    let sheet = ss.getActiveSheet();
    ScriptApp.newTrigger("onFormSubmitProcessor")
      .forSpreadsheet(ss)
      .onFormSubmit()
      .create();
  
}

function onFormSubmitProcessor(e) {
}

