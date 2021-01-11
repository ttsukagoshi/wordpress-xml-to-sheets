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

/* exported onOpen, showDialog, parseWpXml */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('WordPress XML to Sheets')
    .addItem('Parse XML', 'showDialog')
    .addToUi();
}

function showDialog() {
  var html = HtmlService.createHtmlOutputFromFile('src/dialog');
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload WordPress XML');
}

function parseWpXml(data) {
  // Parse XML data
  var blob = Utilities.newBlob(data.bytes, data.mimeType, data.name);
  console.log(`File "${blob.getName()}" (Content type: ${blob.getContentType()}) retrieved.`); // log
  var document = XmlService.parse(blob.getDataAsString("UTF-8"));
  // Get contents
  var contents = { items: [] };
  var targetMetaContents = [
    'title',
    'link',
    'description',
    'pubDate',
    'language',
    'wxr_version',
    'generator'
  ];
  var targetItemContents = [
    'title',
    'link',
    'pubDate',
    'creator',
    'description',
    'post_id',
    'status',
    'post_type'
  ];
  var channelContents = document.getRootElement().getChild('channel').getChildren();
  channelContents.forEach(element => {
    let elementName = element.getName();
    if (targetMetaContents.includes(elementName)) {
      contents[elementName] = element.getValue();
    } else if (elementName === 'item') {
      let itemContents = element.getChildren().reduce((item, itemElement) => {
        let itemElementName = itemElement.getName();
        if (targetItemContents.includes(itemElementName)) {
          item[itemElementName] = itemElement.getValue();
        }
        return item;
      }, {});
      contents.items.push(itemContents);
    }
  });
  // Copy on spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var now = Utilities.formatDate(new Date(), tz, 'yyyyMMddHHmmss');
  var templateChannel = ss.getSheetByName('Template - Channel');
  var templateItems = ss.getSheetByName('Template - Items');
  // Website Meta Data (Channel)
  var channelSheet = ss.insertSheet(`Channel_${now}`, 0, { template: templateChannel });
  var channelSheetData = [[
    contents.title,
    contents.link,
    contents.description,
    Utilities.formatDate(new Date(contents.pubDate), tz, 'yyyy/MM/dd (E) HH:mm:ss Z'),
    contents.language,
    contents.wxr_version,
    contents.generator
  ]];
  channelSheet.getRange(4, 2, channelSheetData.length, channelSheetData[0].length)
    .setValues(channelSheetData);
  // Website Item (Page & Post) Data
  var itemSheet = ss.insertSheet(`Items_${now}`, 1, { template: templateItems });
  var itemSheetData = contents.items.map(item => [
    item.title,
    item.link,
    Utilities.formatDate(new Date(item.pubDate), tz, 'yyyy/MM/dd (E) HH:mm:ss Z'),
    item.creator,
    item.description,
    item.post_id,
    item.status,
    item.post_type
  ]);
  itemSheet.getRange(4, 2, itemSheetData.length, itemSheetData[0].length)
    .setValues(itemSheetData);
}