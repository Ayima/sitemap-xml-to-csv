/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Launch sidebar', 'showSidebar')
    .addItem('About', 'showAbout')
    .addToUi();
}

function showAbout() {
  var html = HtmlService.createHtmlOutputFromFile('about')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('About')
    .setWidth(250)
    .setHeight(450);
  SpreadsheetApp.getActive().show(html);
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('index')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('Sitemap XML to CSV')
    .setWidth(300);

  // Open sidebar
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Make a new sheet.
 */
function makeNewSheet(spreadsheet, newSheetName){
  var sitemapSheet = spreadsheet.getSheetByName(newSheetName);
  if (sitemapSheet) {
    sitemapSheet.clear();
    sitemapSheet.activate();
  } else {
    sitemapSheet =
        spreadsheet.insertSheet(newSheetName, spreadsheet.getNumSheets());
  }
  return sitemapSheet;
}

/**
 * Show msg to user and get input.
 */
function userInput(msg){
  var input = Browser.inputBox(msg, Browser.Buttons.OK_CANCEL);
  if (input == 'cancel') {
    return;
  }
  return input.trim();
}

/**
 * Get the content of a page.
 */
function makeHttpRequest(url){
  var page = UrlFetchApp.fetch(url);
  if (page == null)
    throw '';
  return page;
}

/**
 * Update the user on the current script progress.
 */
function writeLoadingMessage(sitemapSheet, msg){
  sitemapSheet.getRange(1, 1, 1, 1).setValues([[msg]]);
  SpreadsheetApp.flush();
}

/**
 * Write initial warning message.
 */
function showDeleteMessage(){
  var msg = [
    'Clear all sheets in current workbook?',
    'In order to download as much of the sitemap as possible, the entire workbook will be cleared.\n\n'+
    'Click "Yes" to delete all sheets in this workbook, click "No" to cancel.'
  ];
  var ui = SpreadsheetApp.getUi();
  result = ui.alert(msg[0], msg[1], ui.ButtonSet.YES_NO);
  return result;
}

/**
 * Write error message.
 */
function showFailMessage(err, url){
  var msg = ["Can't download sitemap"]
  // Logger.log(err);
  if (err === 'gzip'){
    msg.push('Sorry, the sitemap ' + url + ' is in gzip format, '
             + 'which is not currently supported for Add-ons.');
  } else if (err.toString().split('returned code').length === 2) {
    var statusCode = err.toString().split('returned code')[1].trim();
    msg.push('Sorry, there was a problem downloading the sitemap '
             + url + ' (status code ' + statusCode + ').');
  } else {
    msg.push('Sorry, there was a problem downloading the sitemap '
             + url + '.');
  }
  var ui = SpreadsheetApp.getUi();
  ui.alert(msg[0], msg[1], ui.ButtonSet.OK);
}

/**
 * Write cutoff warning message.
 */
function showCutoffMessage(numRows){
  var msg = [
    'Sitemap only partially downloaded',
    'Sorry, ' + numRows + ' rows were removed due to Add-ons maximum of 2 million cells per workbook.'
  ];
  var ui = SpreadsheetApp.getUi();
  ui.alert(msg[0], msg[1], ui.ButtonSet.OK);
}

/**
 * Check if the sitemap is an index.
 */
function isSitemapIndex(pageText){
  if (pageText.indexOf('<sitemap>') !== -1)
    return true;
  else
    return false;
}

/**
 * Get the page text. Raise error if gzip.
 */
function loadPageText(page){
  if (page.getBlob().getContentType() == 'application/x-gzip')
    // var pageText = unGzip(page.getContent());
    throw 'gzip';
  else
    var pageText = page.getContentText();
  return pageText;
}

/**
 * Parse the sitemap data from the on-page text.
 */
function getPageData(pageText, url){
  var doc = XmlService.parse(pageText.trim());  
  var rootElement = doc.getRootElement();
  var children = rootElement.getChildren();
  var data = [];
  children.forEach(function(child) {
    var subDict = {};
    var subChildren = child.getDescendants();
    var i = 1;
    subChildren.forEach(function(c) {
      if (typeof c.getName === 'function') {
        if ( Object.keys(subDict).indexOf(c.getName()) == -1 ) {
          try {
            if (c.getText())
              subDict[c.getName()] = c.getText();  
          }
          catch(err) {}
        } else {
          try {
            while (true) {
              var newCol = c.getName() + ' - ' + i;
              if ( Object.keys(subDict).indexOf(newCol) == -1 )
                break;
              i += 1;
            }
            if (c.getText())
              subDict[newCol] = c.getText();
          } catch(err) {
            // Logger.log('Error in new col for subdict');Logger.log(err);
          }
        }
      }
    });
    data.push([url, subDict]);
  });
  return data;
}

/**
 * Parse sitemap URLs from sitemap index page.
 */
function getSitemapUrls(pageText){
  var doc = XmlService.parse(pageText);
  var rootElement = doc.getRootElement();
  var locElements = getElementsByTagName(rootElement, 'loc');
  var urls = [];
  locElements.forEach(function(url){
    urls.push(url.getText());
  });
  return urls;
}

/**
 * Utility function.
 * Source: https://sites.google.com/site/scriptsexamples/learn-by-example/parsing-html
 */
function getElementsByTagName(element, tagName) {  
  var data = [];
  var descendants = element.getDescendants();  
  for(i in descendants) {
    var elt = descendants[i].asElement();     
    if( elt !=null && elt.getName()== tagName) data.push(elt);      
  }
  return data;
}

/**
 * Transform the data from [[url, {key:val, ...}], ...]
 * to table format: [[url, val, ...], ...]
 */
function buildTable(data){
  var allKeys = [];
  data.forEach(function(d){
    Object.keys(d[1]).forEach(function(key){
      if (allKeys.indexOf(key) == -1)
        allKeys.push(key);
    });
  });
  var headers = ['Sitemap URL']
  headers.push.apply(headers, allKeys);
  var table = [headers];
  data.forEach(function(d){
    var newRow = [d[0]];
    allKeys.forEach(function(key){
      if (Object.keys(d[1]).indexOf(key) !== -1)
        newRow.push(d[1][key]);
      else
        newRow.push('');
    });
    table.push(newRow);
  });
  return table;
}

/**
 * Write the table data to the spreadsheet,
 * keeping to less than 2M total cells in workbook
 */
function writeTable(spreadsheet, sheet, table){
  var padding = 10000;
  var maxNumberOfCells = 2000000 - padding;
  var maxLength = maxNumberOfCells / table[0].length;
  var i = 1;
  var numColumns = table[0].length;
  if (sheet.getMaxColumns() > numColumns)
    sheet.deleteColumns(numColumns+1, sheet.getMaxColumns()-numColumns);
  else if (sheet.getMaxColumns() < numColumns)
    sheet.insertColumns(1, numColumns-sheet.getMaxColumns());
  if (table.length > maxLength)
    var cutoff = table.splice(maxLength).length;
  else
    var cutoff = 0;
  sheet.getRange(1, 1, table.length, table[0].length).setValues(table);
  if (cutoff)
    showCutoffMessage(cutoff);
}

/**
 * Load user-input sitemap URL
 */
function getSitemapFromUrl(sitemapURL){
  var result = showDeleteMessage();
  if (result != ui.Button.YES)
    return;
  if ((sitemapURL == null) || (sitemapURL == ''))
    return;
  getSitemap(sitemapURL);
}

/**
 * Request sitemap data and convert to CSV.
 */
function getSitemap(url){
  
  // Create new sheet and delete all others
  var spreadsheet = SpreadsheetApp.getActive();
  var newSheetName = 'Sitemap (' + url.split('//').slice(-1)[0].split('/')[0] + ')';
  var sitemapSheet = makeNewSheet(spreadsheet, newSheetName);
  sitemapSheet.activate();
  var allSheets = spreadsheet.getSheets();
  allSheets.forEach(function(sheet){
    if (sheet.getName() !== sitemapSheet.getName())
      spreadsheet.deleteSheet(sheet);
  });
    
  // Request the sitemap resource at the given URL
  try {
    var page = makeHttpRequest(url);
  } catch(err) {
    showFailMessage(err, url, page);
    //writeFailMessage(sitemapSheet, err, url, page);
    return;
  }
  try {
    var pageText = loadPageText(page);
  } catch(err) {
    showFailMessage(err, url, page);
    return;
  }
  
  // Determine what the sitemap looks like and then get the data
  if (isSitemapIndex(pageText)) {
    var sitemapUrls = getSitemapUrls(pageText);
    var data = [];
    for (var i = 0; i < sitemapUrls.length; i++){
    // sitemapUrls.forEach(function(_url, i){
      var _url = sitemapUrls[i];
      var num = i + 1;
      var msg = 'Reading data from sitemap ';
      msg += num + ' / ' + sitemapUrls.length;
      msg += ': ' + _url;
      //showLoadingMessage(msg);
      writeLoadingMessage(sitemapSheet, msg); 
      try {
        var page = makeHttpRequest(_url);
      } catch(err) {
        showFailMessage(err, _url, page);
        //writeFailMessage(sitemapSheet, err, _url, page);
        return;
      }
      try {
        var pageText = loadPageText(page);
      } catch(err) {
        showFailMessage(err, _url, page);
        return;
      }
      var d = getPageData(pageText, _url);
      data.push.apply(data, d);
    };
  } else {
    var msg = 'Reading data from sitemap ' + url;
    //showLoadingMessage(msg);
    writeLoadingMessage(sitemapSheet, msg);
    data = getPageData(pageText, url);
  }
  
  var table = buildTable(data);
  writeTable(spreadsheet, sitemapSheet, table);

}
