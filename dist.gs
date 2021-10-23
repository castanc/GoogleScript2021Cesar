/* MIT LICENSE
Copyright <2021> <PwC>

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

*/


// Compiled using undefined undefined (TypeScript 4.4.3)
var exports = exports || {};
var module = module || { exports: exports };
Object.defineProperty(exports, "__esModule", { value: true });
exports.TPEDistribution = void 0;
//import { GSResponse } from "../models/GSResponse";
//import { Utils } from "../Utils";

var header = [];

class TPEDistribution {
  constructor(_fileIUrl) {
    this.FileUrl = _fileIUrl;
  }





  processNamedValues(user, namedValues) {
    let r = new GSResponse();
    let updateHeader = false;
    let processName = "";
    let subProcessColName = "Sub Process";
    let headerKey = "single";
    let startIndex = 2;
    let processIndex = 1;
    let rowsProcessed = 0;
    let excludeColumns = `Timestamp,${subProcessColName}`;
    let masterHeader = [];



    let data = {}
    let hashObject = {}

    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let lTab = Utils.createLogTab(ss);
    r.logToTab(lTab, "C01", JSON.stringify(namedValues));


    for (let prop in namedValues) {
      try {
        let val = Utils.getNonEmptyValue(namedValues[prop]);
        if (val != null) {
          data[prop] = val;
          if ( prop.toLowerCase()!="timestamp")
            hashObject[prop] = val;
        }
      }
      catch (ex) {
        Utils.logException(lTab, ex);
      }
    }

    r.logToTab(lTab, "DATA", `${JSON.stringify(data)}`);
    if (data["What brings you here today"][0].includes("multiple"))
      headerKey = "multiple";

    processName = data[subProcessColName][0];


    let headerTab = Utils.getCreateTab(ss, "DynamicHeaders");
    let headerRange = headerTab.getDataRange();
    let lastRow = headerRange.lastRow;
    let lastCol = headerRange.lastCol;
    let ixSingle = 0;
    let ixMultiple = 0;
    let ixHeader = 0;
    let rangeA1 = "";
    let existingHeader = true;
    let headerModified = false;
    headerTab.hideSheet();

    let tab = Utils.getCreateTab(ss, processName);
    var rangeTab = tab.getDataRange();
    var lastColumnTab = rangeTab.getLastColumn();
    var lastRowTab = rangeTab.getLastRow();
    updateHeader = (lastRowTab < 2);


    let predefHeaders = headerTab.getDataRange().getValues().filter(x => x[1] == processName);
    if (predefHeaders && predefHeaders.length == 0) {
      headerTab.appendRow(["single", processName, "empty"]);
      tab.appendRow(["single", processName, "empty"]);

      headerTab.appendRow(["multiple", processName, "empty"]);
      tab.appendRow(["multiple", processName, "empty"]);
      headerModified = true;

    }

    let headerGrid = Utils.getDataObjectFromTab(headerTab);

    let headerRow = headerGrid.filter(x => x.SubProcess == processName && x.ProcType == headerKey);
    ixHeader = headerRow[0].RowId;

    masterHeader = headerTab.getDataRange().getValues().filter(x => x[1] == processName && x[0].toLowerCase() == headerKey)[0];
    Utils.Validate(!masterHeader, r, lTab, 400, -4, `Master Header undefined for\t${processName}\t${headerKey}`);
    Utils.Validate(masterHeader && masterHeader.length == 0, r, lTab, 400, -4, `Master Header not found for\t${processName}\t${headerKey}`);

    if (masterHeader.indexOf("empty") == 2) {
      let dynHeader = `${headerKey}\t${processName}\tStatus\tProcessor\tTimestamp`.split("\t");
      for (let item in data) {
        if (excludeColumns.indexOf(item) < 0)
          dynHeader.push(item);
      }
      masterHeader = dynHeader;
      headerModified = true;
    }

    //Master Keys tab
    let mkTab = Utils.getCreateTab(ss, "TPEMasterKeys", "TimeStamp,User,StartRow,EndRow,Hash,SubProcess,TargetRow");
    mkTab.hideSheet();

    let md5Hash = MD5(JSON.stringify(hashObject));
    let rangeData = mkTab.getDataRange();
    let gridVerify = Utils.getDataObjectFromTab(mkTab).filter(x => x.Hash == md5Hash);
    Utils.Validate(gridVerify.length > 0, r, lTab, 400, -3, `Row already pocessed Hash: ${md5Hash}`);

    try {

      let dataText = "";
      let addedFields = 0;
      let row = new Array(4 + masterHeader.length - 1);
      row[0] = headerKey;
      row[1] = data[subProcessColName][0];  //process name
      row[2] = "";                     //status
      row[3] = "";                     //processor
      row[4] = data["Timestamp"][0];             //timestamp


      for (let prop in data) {
        if (excludeColumns.indexOf(prop) < 0) {
          let ix = masterHeader.indexOf(prop);
          if (ix == - 1) {
            headerModified = true;
            masterHeader.push(prop);
            ix = masterHeader.indexOf(prop);
          }
          if (ix > 4) {
            r.logToTab(lTab, "ADDED", `${ix}\t${prop}\t${data[prop][0]}`);
            row[ix] = data[prop][0];
          }
        }
      }
      //r.logToTab(lTab, "FINAL ROW", `${JSON.stringify(row)}`);
      tab.appendRow(row);
      mkTab.appendRow([Utils.getTimeStamp(), user, 0, 0, md5Hash, processName, tab.getLastRow()]);


      //todo: update master header
      if (headerModified) {
        rangeA1 = `A${ixHeader}:${Utils.getExcelColumnName(masterHeader.length)}${ixHeader}`;
        headerRange = headerTab.getRange(rangeA1);
        headerRange.setValues([masterHeader]);

        let ixHeaderTab = 1;
        if (headerKey == "multiple")
          ixHeaderTab = 2;
          
        rangeA1 = `A${ixHeaderTab}:${Utils.getExcelColumnName(masterHeader.length)}${ixHeaderTab}`;

        headerRange = tab.getRange(rangeA1);
        headerRange.setValues([masterHeader]);
        headerModified = false;
      }
      rowsProcessed++;
      r.result = 200;
      r.domainResult = 0;
      r.message = "Process completed succesfully";
      r.resultLink = "";
      r.logToTab(lTab, "PROCESSED OK", `${processName}\t${headerKey}\t\thash\t${md5Hash}`);

      //todo: this autofits column width, but not working
      //  let cols = tab.getDataRange().getValues()[0];
      //  for(let j=0;j<cols.length;j++)
      //   tab.autoResizeColoumn(j);
    }
    catch (ex) {
      Utils.logException(lTab, ex);
    }

    return r;
  }


  DistributeRow2(user, grid, lastColumn, firstRow, lastRow) {
    let r = new GSResponse();
    let updateHeader = false;
    let processName = "";
    let multiRequest = false;
    let mainTabName = mainSheetName; //"Form Responses 1"
    let subProcessColName = "Sub Process";
    let headerKey = "single";
    let startIndex = 2;
    let processIndex = 1;
    let rowsProcessed = 0;



    let ss = SpreadsheetApp.getActiveSpreadsheet(); //Utils.openSpreadSheet(this.FileUrl);

    //todo: calculate and save to global cache the minimum number of columns = 
    //min ProcessList.row.length

    let lTab = Utils.createLogTab(ss);
    Utils.Validate(grid[0].length < 2, r, lTab, 400, -1, "Invalid range provided, must be a whole row");

    let mainTab = ss.getSheetByName(mainTabName);
    Utils.Validate(!mainTab, r, lTab, 400, -2, `Spreadsheet ${ss.getName()} has no main tab ${mainTabName}`);


    let fullHeader = Utils.GetFullHeader(mainTab);
    Utils.Validate(!fullHeader, r, lTab, 400, -4, `Can't access main header for ${mainTabName}`);

    //Search index of "Sub Process" coilumn
    processIndex = -1;
    for (let j = 0; j < fullHeader.length; j++) {
      if (fullHeader[j] == subProcessColName && grid[0][j].toString().trim() != "") {

        processIndex = j;
        if (j > 10)
          headerKey = "mass";
        break;
      }
    }

    Utils.Validate(processIndex == -1, r, lTab, 400, -5, `Main sheet ${mainTabName} has nmo column for Sub process ${subProcessColName}`);






    //Master Keys tab
    let mkTab = Utils.getCreateTab(ss, "TPEMasterKeys", "TimeStamp,User,StartRow,EndRow,Hash,SubProcess,TargetRow");
    //let tabParameters = Utils.getCreateParameters(ss);
    //tabParameters.hideSheet();

    mkTab.hideSheet();
    let md5Hash = MD5Row(grid[0]);

    let rangeData = mkTab.getDataRange();
    let gridVerify = Utils.getDataObjectFromTab(mkTab).filter(x => x.Hash == md5Hash);
    Utils.Validate(gridVerify.length > 0, r, lTab, 400, -3, `Row already pocessed Hash: ${md5Hash}`);

    r.logToTab(lTab, "M0", "");
    r.logToTab(lTab, "M1", `Start row processing. Start Row:\t${firstRow}\tEnd row\t${lastRow}`);

    for (let i = 0; i < grid.length; i++) {

      r.logToTab(lTab, "M0", "");
      r.logToTab(lTab, "row", `processing row\t${i + firstRow}\thash\t${md5Hash}`);

      try {

        processName = grid[i][processIndex].trim();
        if (processName == "") {
          headerKey = "mass";
          processIndex = 65;
          processName = grid[i][processIndex];

          multiRequest = true;
          startIndex = 3;
        }

        let masterHeader = Utils.getCreateTab(ss, "ProcessList").getDataRange().getValues().filter(x =>
          x[0] == processName && x[1].toLowerCase() == headerKey)[0];


        Utils.Validate(!masterHeader, r, lTab, 400, -5, `Master header not found in ProcessList for\t${processName}\t${headerKey}`);
        Utils.Validate(masterHeader.length < 3, r, lTab, 400, -5, `Empty Master header for process\t${processName}\t${headerKey}`);
        Utils.Validate(masterHeader[3].trim() == "", r, lTab, 400, -5, `Empty Master header for process\t${processName}\t${headerKey}`);


        let tab = Utils.getCreateTab(ss, processName);
        var rangeTab = tab.getDataRange();
        var lastColumnTab = rangeTab.getLastColumn();
        var lastRowTab = rangeTab.getLastRow();

        //tab is new requires header
        updateHeader = (lastRowTab < 2);


        r.logToTab(lTab, "M4",
          `Tab info: Name:\t${processName}\tlastColumn:\t${lastColumnTab}\tlastRow:\t${lastRowTab}`)


        var lastMasterCol = 0;
        for (let i = 2; i < masterHeader.length; i++) {
          if (masterHeader[i].trim() == "") {
            lastMasterCol = i - 1;
            break;
          }
        }


        let header = [];
        if (updateHeader) {
          let newHeader = "Status\tProcessor\tTimeStamp\tSub Process";
          for (let i = 2; i < masterHeader.length; i++)
            newHeader = `${newHeader}\t${masterHeader[i]}`;

          header = newHeader.split("\t");
          tab.appendRow(header);
          updateHeader = false;
        }

        let ix = 0;
        let headerText = "";
        let dataText = "";
        let addedFields = 0;

        let row = new Array(masterHeader.length + 1);
        row[0] = "";                     //status
        row[1] = "";                     //processor
        row[2] = grid[i][0];             //timestamp
        row[3] = grid[i][processIndex];  //process name


        //Adding field	What brings you here today	1	3	C	I would like to submit a request for single resource	Target Col	5	E

        r.logToTab(lTab, "M3", `Fields Add\Header\tOrder\tSource Col#\tSourceColName\tData\tTargetIndex\tTarget ColName`);


        for (let j = 2; j < fullHeader.length; j++) {
          headerText = fullHeader[j];
          dataText = grid[i][j];
          if (!dataText) {
            dataText = "";
            //r.logToTab(lTab,"UNDEFINED", `Column j: ${j} ${headerText} Data UNDEFINED`);
          }
          if (dataText.toString().trim() != "") {
            ix = masterHeader.indexOf(headerText);
            if (ix > 1) {
              if (ix < row.length) {
                ix += 2;
                row[ix] = dataText;
                addedFields++;
                r.logToTab(lTab, "M3", `Adding field\t${headerText}\t${addedFields}\t${j + 1}\t${Utils.getExcelColumnName(j + 1)}\t${dataText}\t${ix + 1}\t${Utils.getExcelColumnName(ix + 1)}`);
              }
              else
                r.logToTab(lTab, "M3", `FIELD OUT\t${headerText}\t${addedFields}\t${j}\t${Utils.getExcelColumnName(j + 1)}\t${dataText}\t${ix}\t${Utils.getExcelColumnName(ix + 1)}`);

            }
          }

        }
        tab.appendRow(row);
        mkTab.appendRow([Utils.getTimeStamp(), user, firstRow, lastRow, md5Hash, processName, tab.getLastRow()]);
        rowsProcessed++;

        //todo: this autofits column width, but not working
        //  let cols = tab.getDataRange().getValues()[0];
        //  for(let j=0;j<cols.length;j++)
        //   tab.autoResizeColoumn(j);

        //Process only the first row
        break;
      }
      catch (ex) {
        r.logToTab(lTab, "EXCEPTION", ex.message);
        Utils.logException(lTab, ex);
        break;

      }
    }
    r.result = 200;
    r.domainResult = 0;
    r.message = "Process completed succesfully";
    r.resultLink = "";
    r.logToTab(lTab, "M99", `${rowsProcessed} Processed rows OK.`);
    return r;
  }





  StatusChange(user, value, rowNum, sheetName) {
    let r = new GSResponse();
    if (value.toUpperCase().startsWith("C")) {
      try {
        let ss = SpreadsheetApp.getActiveSpreadsheet(); //Utils.openSpreadSheet(this.FileUrl);   
        let lTab = Utils.createLogTab(ss);
        let sheet = ss.getSheetByName(sheetName);
        r.logToTab(lTab, "M1", `Status Change()\t${user}\t${value}\t${rowNum}\t${sheetName}`);

        let completedTab = Utils.getCreateTab(ss, "Completed");
        let row = sheet.getDataRange().getValues()[rowNum - 1];

        //todo: in multiuser env, might move the wrong row
        //let target = completedTab.getRange(completedTab.getLastRow() + 1, 1);
        completedTab.appendRow(row);
        r.logToTab(lTab, "StatusChange()", "M2", "Row added to Completed");
        sheet.deleteRow(rowNum);
        r.logToTab(lTab, "StatusChange()", "M3", `Row removed from\t${sheetName}\t${rowNum}`);

        r.logToTab(lTab, "M99", `Row from\t${sheetName}\t${rowNum}\tto Completed, by user\t ${user}`);


        //ss.getRange(rowNum, 1, 1, row.length).moveTo(target);
      }
      catch (ex) {
        Utils.logException(lTab, ex);
        Utils.Validate(true, r, lTab, 500, -2, `Exception at StatusCahnge()\t${ex.message}`);
      }


    }

    return result;


  }


  BatchStatusChange() {
    let r = new GSResponse();

    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let lTab = Utils.createLogTab(ss);
    r.logToTab(lTab, "M1", "Starting moving Completed records");
    let movedRecs = 0;
    let tabCompleted = Utils.getCreateTab(ss, "Completed");

    for (let i = 0; i < ss.getNumSheets(); i++) {
      let sheet = ss.getSheets()[i];
      if (sysTabs.indexOf(sheet.getName()) >= 0)
        continue;

      let grid = sheet.getDataRange().getValues();

      for (let j = grid.length - 1; j >= 1; j--) {
        try {
          if (grid[j][0].toUpperCase().startsWith("C")) {
            tabCompleted.appendRow(grid[j]);
            sheet.deleteRow(j + 1);
            r.logToTab(lTab, "M2", `Moving\t${sheet.getName()}\trow\t${j}`);
            movedRecs++;
          }
        }
        catch (ex) {
          Utils.logException(lTab, ex);
          //r.logToTab(lTab, "EXCEPTION", `Moving\t${sheet.getName()}\trow\t${j}`);
        }

      }
    }
    r.logToTab(lTab, "M99", `End moving Completed\t${movedRecs}\trecords`);

    Browser.msgBox(`Completed ${movedRecs}records were moved to Completed tab`)


    return r;
  }

}
exports.TPEDistribution = TPEDistribution;
