/* MIT LICENSE
Copyright <2021> <PwC>

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OOR OTHER DEALINGS IN THE SOFTWARE.

*/



// Compiled using undefined undefined (TypeScript 4.4.3)
var exports = exports || {};
var module = module || { exports: exports };
Object.defineProperty(exports, "__esModule", { value: true });
exports.Utils = void 0;
//import { FileInfo } from "./models/FileInfo";
//import { GSResponse } from "./models/GSResponse";
class Utils {
  static getCreateFolder(folderName) {
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = null;
    if (folders.hasNext())
      folder = folders.next();
    else
      folder = DriveApp.createFolder(folderName);
    return folder;
  }

  static getNonEmptyValue(arr) {

    var notEmpty = arr.filter(x => Utils.isNotEmpty(x));
    if (notEmpty.length > 0)
      return notEmpty;
    else
      return null;

  }

  static GetFullHeader(tab) {
    let header = tab.getDataRange().getValues()[0];

    return header;
  }

  static deleteTabs(url, start = 1) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    let totalSheets = ss.getNumSheets();
    let i = 0;
    while (i <= ss.getNumSheets() - 1) {
      let sh = ss.getSheets()[i];
      if (sh) {
        let sheetName = sh.getSheetName();

        if ("ProcessList,Form Responses 1".indexOf(sheetName) < 0)
          ss.deleteSheet(sh);
        else i++;
        totalSheets = ss.getNumSheets();
      }
    }
  }


  static getColInPropertyName(arr, propName) {
    let ix = -1;
    for (let i = 0; i < arr.length; i++) {
      if (propName.contains(arr[i])) {
        ix = i;
        break;
      }
    }
    return ix;
  }


  static getCreateParameters(ss) {
    let tabParameters = Utils.getCreateTab(ss, "Parameters", "Key,Values");
    let data = tabParameters.getDataRange().getValues();
    var hasData = data.length > 0;
    var hasStatus = false;
    if (hasData) {
      var statusData = data.filter(x => x[0] == "status");
      hasStatus = statusData.length > 0;
    }
    if (!hasStatus) {
      tabParameters.appendRow(["status", "Status1", "Status2", "Status3", "Completed"]);
    }
    return tabParameters;
  }

  static createLogTab(ss, tabName = "ProcessLog", columns = "TimeStamp,Type,Message") {
    var lTab = Utils.getCreateTab(ss, tabName, columns);
    lTab.appendRow([Utils.getTimeStamp(), ""]);
    lTab.hideSheet();
    //Utils.deleteAllRows(lTab);
    return lTab;
  }

  static logException(lTab, ex) {
    let vDebug = "";
    for (var prop in ex) {
      vDebug += "property: " + prop + " value: [" + ex[prop] + "]\n";
    }
    vDebug += "toString(): " + " value: [" + ex.toString() + "]";

    let row = [Utils.getTimeStamp(), "EXCEPTION", vDebug];
    lTab.appendRow(row);

    row = [Utils.getTimeStamp(), "STACK", ex.stack];
    lTab.appendRow(row);
  }

  static setDropdown(ss, targetRange, sourceRangange) {
    // Set the data validation for cell A1 to require a value from B1:B10.
    var cell = SpreadsheetApp.getActive().getRange('A1');
    var range = SpreadsheetApp.getActive().getRange('B1:B10');
    var rule = SpreadsheetApp.newDataValidation().requireValueInRange(range).build();
    cell.setDataValidation(rule);
  }


  static ValidateEmail(mail) {
    if (/^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*$/.test(mail))
      return true;
    return false;
  }
  static getTimeStamp(dt = null) {
    if (dt == null)
      dt = new Date();
    return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm-ss');
  }

  //https://cwestblog.com/2013/09/05/javascript-snippet-convert-number-to-column-name/
  static getExcelColumnName(num) {
    for (var ret = '', a = 1, b = 26; (num -= a) >= 0; a = b, b *= 26) {
      ret = String.fromCharCode(parseInt((num % b) / a) + 65) + ret;
    }
    return ret;
  }
  static getDocTextByName(fileName) {
    var text = "";
    try {
      let doc = this.openDoc(fileName);
      if (doc != null)
        text = doc.getBody().getText();
    }
    catch (ex) {
      Logger.log(`getDocTextByName() Exception. fileName: ${fileName} ${ex.message}`);
    }
    return text;
  }
  static RowIndexOf(arr, value) {
    let n = -1;
    for (var i = 0; i < arr.length; i++) {
      if (arr[i] == value) {
        n = i;
        break;
      }
    }
    return n;
  }
  static getNameFromEmail(email) {
    let name = email;
    let index = name.indexOf("@");
    if (index > 0) {
      name = name.substring(0, index);
      let parts = name.split('.');
      name = "";
      for (var i = 0; i < parts.length; i++) {
        parts[i] = parts[i].trim;
        if (parts[i].length > 0) {
          name = `${name}${parts[i].substring(0, 1).toUpperCase()}${parts[i].substring(1)} `;
        }
      }
    }
    return name;
  }
  //https://stackoverflow.com/questions/16840038/easiest-way-to-get-file-id-from-url-on-google-apps-script
  static getIdFromUrl(url) {
    return url.match(/[-\w]{25,}/);
  }
  static getFileInfo(name) {
    let fi = new FileInfo(null);
    let id = this.getIdFromUrl(name);
    //todo: test
    Logger.log("getFileInfo()", id);
    if (id != null && id.length > 0) {
      let file = DriveApp.getFileById(id[0]);
      if (file != undefined) {
        fi = new FileInfo(file);
      }
    }
    else {
      let fileInfos = Utils.getFilesByName(name);
      if (fileInfos.length > 0) {
        fi = fileInfos[0];
      }
    }
    return fi;
  }
  static openSpreadSheet(ssName) {
    let ss = null;
    try {
      if (ssName.toLowerCase().indexOf("http") >= 0) {
        ss = SpreadsheetApp.openByUrl(ssName);
      }
      else {
        let fileInfos = Utils.getFilesByName(ssName);
        if (fileInfos.length > 0)
          ss = SpreadsheetApp.openById(fileInfos[0].id);
        else
          ss = SpreadsheetApp.openById(ssName);
      }
    }
    catch (ex) {
      ss = null;
    }
    return ss;
  }
  static openDoc(docName) {
    let ss = null;
    try {
      if (docName.toLowerCase().indexOf("http") >= 0) {
        ss = DocumentApp.openByUrl(docName);
      }
      else {
        let fileInfos = Utils.getFilesByName(docName);
        if (fileInfos.length > 0)
          ss = DocumentApp.openById(fileInfos[0].id);
        else
          ss = DocumentApp.openById(docName);
      }
    }
    catch (ex) {
      ss = null;
    }
    return ss;
  }
  static getFilesByName(name) {
    let fileInfos = new Array();
    let files = DriveApp.getFilesByName(name);
    while (files.hasNext()) {
      let file = files.next();
      //if ( !file.isTrashed)
      {
        let fi = new FileInfo(file);
        fileInfos.push(fi);
      }
    }
    return fileInfos;
  }
  static getSpreadSheet(folder, fileName) {
    let spreadSheet = null;
    let file = Utils.getFileFromFolder(fileName, folder);
    if (file != null) {
      spreadSheet = SpreadsheetApp.openById(file.getId());
    }
    return spreadSheet;
  }
  static getCreateSpreadSheet(folder, fileName, tabNames = "") {
    let file = Utils.getFileFromFolder(fileName, folder);
    let tabs = tabNames.split(',');
    let spreadSheet = null;
    if (file == null) {
      spreadSheet = SpreadsheetApp.create(fileName);
      if (tabs.length > 0) {
        if (tabs[0].length > 0) {
          var sh = spreadSheet.getActiveSheet();
          sh.setName(tabs[0]);
        }
        for (var i = 1; i < tabs.length; i++) {
          if (tabs[i].length > 0) {
            let itemsSheet = spreadSheet.insertSheet();
            itemsSheet.setName(tabs[i]);
          }
        }
      }
      var copyFile = DriveApp.getFileById(spreadSheet.getId());
      folder.addFile(copyFile);
      DriveApp.getRootFolder().removeFile(copyFile);
      file = Utils.getFileFromFolder(fileName, folder);
    }
    spreadSheet = SpreadsheetApp.openById(file.getId());
    return spreadSheet;
  }
  static getFileByName(fileName) {
    var files = DriveApp.getFilesByName(fileName);
    while (files.hasNext()) {
      var file = files.next();
      return file;
      break;
    }
    return null;
  }
  static getFileFromRoot(name) {
    let files;
    files = DriveApp.getFilesByName(name);
    if (files.hasNext()) {
      return files.next();
    }
    return null;
  }
  static getFileFromFolder(name, folder) {
    let files;
    files = folder.getFilesByName(name);
    if (files.hasNext()) {
      return files.next();
    }
    return null;
  }
  static getData(ss, sheetName) {
    let sheet = ss.getSheetByName(sheetName);
    let grid = [];
    if (sheet) {
      var rangeData = sheet.getDataRange();
      var lastColumn = rangeData.getLastColumn();
      var lastRow = rangeData.getLastRow();
      grid = rangeData.getValues();
    }
    return grid;
  }

  static getSheetData(sheet, startRow = 0) {
    let grid = [];
    if (sheet) {
      var rangeData = sheet.getDataRange();
      var lastColumn = rangeData.getLastColumn();
      var lastRow = rangeData.getLastRow();
      grid = rangeData.getValues();
    }
    return grid;
  }


  static getDataObject(ss, sheetName) {
    let grid = Utils.getData(ss, sheetName);
    let objects = [];
    for (var i = 1; i < grid.length; i++) {
      let o = {};
      o["RowId"] = i+1;
      for (var j = 0; j < grid[i].length; j++) {
        o[grid[0][j]] = grid[i][j];
      }
      objects.push(o);
    }
    return objects;
  }

  static getDataObjectFromTab(tab) {
    var rangeData = tab.getDataRange();
    var grid = rangeData.getValues();
    let objects = [];
    for (var i = 0; i < grid.length; i++) {
      let o = {};
      o["RowId"] = i+1;
      o["ProcType"] = grid[i][0];
      o["SubProcess"] = grid[i][1];
      for (var j = 2; j < grid[i].length; j++) {
        o[grid[0][j]] = grid[i][j];
      }
      objects.push(o);
    }

    return objects;
  }


  static getNewArray(baseCols, arr, procType, startPos = 2, colSep = "\t") {
    let resultArr = baseCols.split(colSep);
    for (let j = startPos; j < arr.length; j++) {
      resultArr.push(arr[j]);

    }

    if (resultArr.length > 0)
      resultArr[0] = procType;
    return resultArr;
  }
  static getArrayPosContaining(arr, text, startPos = 2) {
    let ix = -1;
    ix = arr.indexOf(text);
    if (ix < 0) {
      for (let j = startPos; j < arr.length; j++) {
        if (arr[j].trim() != "") {
          if (text.includes(arr[j])) {
            ix = j;
            break;
          }
        }
      }
    }
    return ix;
  }

  static getCreateTab(ss, tabName, header = "") {
    var newTab = ss.getSheetByName(tabName);
    if (newTab == null) {
      newTab = ss.insertSheet();
      newTab.setName(tabName);
      var cols = header.split(",");
      if (cols.length > 0)
        newTab.appendRow(cols);
    }
    return newTab;
  }
  static getUrl(fileName, folder = null) {
    let files;
    if (folder == null)
      files = DriveApp.getFilesByName(fileName);
    else
      files = folder.getFilesByName(fileName);
    if (files.hasNext()) {
      let file = files.next();
      return file.getUrl();
    }
    return "";
  }
  static getHtmlFromArray(name, caption = "", list, required = false) {
    let onChange = "";
    let requiredText = "";
    if (required)
      requiredText = "required";
    name = name.trim();
    if (caption.length == 0)
      caption = "None";
    var options = `<option value="-1" selected>${caption}</option>`;
    for (var i = 0; i < list.length; i++) {
      options = options + `<option value="${i}">${list[i]}</option>`;
    }
    onChange = `onChange="onChange_${name}('${name}',this.options[this.selectedIndex].value)"`;
    return `<select id="SELECTID" name="SELECTID" ${onChange} ${required}>${options}</select>`;
  }
  static replace(text, value, newValue) {
    try {
      while (text.indexOf(value) >= 0)
        text = text.replace(value, newValue);
    }
    catch (ex) {
    }
    return text;
  }
  static extract(text, start, end) {
    let word = "";
    try {
      let index = text.indexOf(start);
      let index2 = 0;
      while (index >= 0) {
        index += start.length;
        index2 = text.indexOf(end, index);
        if (index2 > index) {
          word = text.substr(index, index2 - index);
        }
        index = text.indexOf(start, index2 + end.length);
      }
    }
    catch (ex) {
      //return GSLog.handleException(ex, "Utils.replace()");
      //return text;
    }
    return word;
  }
  static moveFiles(sourceFileId, targetFolderId) {
    try {
      let file = DriveApp.getFileById(sourceFileId);
      let folder = DriveApp.getFolderById(targetFolderId);
      file.moveTo(folder);
    }
    catch (ex) {
      Logger.log("Exception moving file.");
    }
  }
  static sendMail(to, subject, body) {
    let result = 0;
    let mails = to.split('\n');
    let mailsList = "";
    if (mails.length == 0)
      mailsList = to;
    else
      for (var i = 0; i < mails.length; i++) {
        mails[i] = mails[i].trim();
        if (mails[i].length > 0) {
          if (mailsList.length == 0)
            mailsList = mails[i];
          else
            mailsList = `${mailsList},${mails[i]}`;
        }
      }
    try {
      body = Utils.replace(body, "\\n", "</br>");
      MailApp.sendEmail({
        to: mailsList,
        subject: subject,
        htmlBody: body
      });
      result = 0;
    }
    catch (ex) {
      Utils.ex = ex;
      Logger.log(`Exception sending mail to [${mailsList}]\n${ex.message}\n${ex.stacktrace}`);
      result = -1;
    }
    return result;
  }
  static deleteFiles(fileName, folder = null) {
    let files;
    if (folder == null)
      files = DriveApp.getFilesByName(fileName);
    else
      files = folder.getFilesByName(fileName);
    while (files.hasNext()) {
      let file = files.next();
      file.setTrashed(true);
    }
  }
  static async removeFileByName(fileName) {
    var files = DriveApp.getFilesByName(fileName);
    if (files.hasNext()) {
      var file = files.next();
      file.setTrashed(true);
    }
  }
  static getTextFile(fileName, folder = null) {
    let file;
    if (folder == null) {
      let files = DriveApp.getFilesByName(fileName);
      if (files.hasNext())
        file = files.next();
    }
    else {
      file = Utils.getFileFromFolder(folder, fileName);
      if (file != null)
        return file.getBlob().getDataAsString();
      return "";
    }
  }
  static getTextFileFromFolder(folder, fileName) {
    let file = Utils.getFileFromFolder(folder, fileName);
    if (file != null)
      return file.getBlob().getDataAsString();
    return "";
  }
  static writeTextFile(fileName, text, folder = null) {
    var existing;
    if (folder == null)
      existing = DriveApp.getFilesByName(fileName);
    else
      existing = folder.getFilesByName(fileName);
    // Does file exist? if (existing.hasNext()) {
    var file = null;
    if (existing.hasNext()) {
      file = existing.next();
      file.setTrashed(true);
    }
    folder.createFile(fileName, text, MimeType.PLAIN_TEXT);
  }

  static deleteAllRows(tab) {
    for (let i = tab.getLastRow(); i > 1; i--) {
      tab.deleteRow(i);
    }

  }

  static isEmpty(text) {
    return (text == null || text == undefined || text.trim() == "");
  }

  static isNotEmpty(text) {
    return text != undefined && text.toString().trim().length > 0;
  }


  static gridIsEmpty(g) {

    let nonEmptyCount = 0;
    let empty = g.length == 0;
    if (!empty) {
      for (let i = 0; i < g.length; i++) {
        for (let j = 0; j < g[i].length; j++) {
          empty = Utils.isEmpty(g[i][j]);
          if (!empty) {
            nonEmptyCount++;
            break;
          }

          if (nonEmptyCount > 0)
            break;
        }
      }
      empty = nonEmptyCount == 0;
      return empty;
    }
  }

  static Validate(condition, r, lTab, error, domError, message) {
    if (condition) {
      r.logToTab(lTab, `Error ${error} ${domError}`, message);
      //Forces an abort calling a nonexistent function
      r.ForceExit();
    }
  }
}
exports.Utils = Utils;

