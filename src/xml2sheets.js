// Copyright 2021 Taro TSUKAGOSHI
// 
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
// 
//     http://www.apache.org/licenses/LICENSE-2.0
// 
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

/* exported onOpen, showDialog, xmlParse */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('XML to Sheets')
    .addItem('Parse XML', 'showDialog')
    .addToUi();
}

function showDialog() {
  var html = HtmlService.createHtmlOutputFromFile('src/dialog');
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload XML');
}

function xmlParse(form) {
  var blob = form.xmlFile;
  Logger.log(`File "${blob.getName()}" retrieved.`); // log
  Logger.log(blob.getContentType()); // test
}