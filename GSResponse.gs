
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
exports.GSResponse = void 0;
//import { KeyValuePair } from "./KeyValuePair";
//import { Utils } from "../Utils";
class GSResponse {
    constructor() {
        this.domainResult = -1;
        this.html = new Array();
        this.script = new Array();
        this.error = new Array();
        this.logData = new Array();
        this.localRenders = new Array();
        this.messages = new Array();
        this.result = 200;
        this.logEnabled = true;
        this.message = "";
        this.resultLink = "";
        this.domainResult = -1;
        this.html = new Array();
        this.script = new Array();
        this.error = new Array();
        this.localRenders = new Array();
        this.messages = new Array();
        this.domainResult = 0;
        this.result = 200;
        this.message = "";
        this.lTab = {};
    }
    addHtml(key, value) {
        this.html.push(new KeyValuePair(key, value));
    }
    addScript(key, value) {
        this.script.push(new KeyValuePair(key, value));
    }
    addError(key, value) {
        this.error.push(new KeyValuePair(key, value));
    }
    addLocalRenders(key, value) {
        this.localRenders.push(new KeyValuePair(key, value));
    }
    addLog(text) {
        if (this.logEnabled)
            this.logData.push(`${Utils.getTimeStamp()} | ${text}`);
    }

    getLogs()
    {
      var text = "";
      for(var i =0;i<this.logData.length; i++)
      text = `${text}\n<br>${this.logData[i]}`;
      return text;
    }

    logToTab(tab,logType,logMessage)
    {
      //if ( !logEnabled)
      //  return;
      this.lTab = tab;
      
      if ( logMessage == undefined || !logMessage)
      {
        logMessage = "UNDEFINED"
      }
      var msgArr = logMessage.split("\t");

      var row = [Utils.getTimeStamp(),logType];
      for(var i=0;i<msgArr.length;i++)
        row.push(msgArr[i]);

      tab.appendRow(row);
    }

    logToTabExt(logType,logMessage)
    {
      var row = [Utils.getTimeStamp(),logType,logMessage];
    }


}
exports.GSResponse = GSResponse;
