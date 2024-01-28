const sheet = SpreadsheetApp.getActive();
const sheetTestEnvironment = sheet.getSheetByName('Email_Setup');
const sheetTable = sheet.getSheetByName('Tables');
const sheetTextMessage = sheet.getSheetByName('Text_Message');
const sheetCharts = sheet.getSheetByName('Charts');
const sheetLinks = sheet.getSheetByName('Links');
const sheetOverview = sheet.getSheetByName('Overview');
let aliases = GmailApp.getAliases();
let htmlMessage = [];
let htmlHyperLink = [];
let htmlTables = [];
let htmlCharts = [];


// retrieves the message in tab Text_Message
function getMessage() {
  var messageRange = sheetTextMessage.getRange(2, 3, sheetTextMessage.getLastRow()-1, 1);

  // Read message contents
  var data = messageRange.getDisplayValues();

  // Get text styling
  var fontColors = messageRange.getFontColors();
  var backgrounds = messageRange.getBackgrounds();
  var fontFamilies = messageRange.getFontFamilies();
  var fontSizes = messageRange.getFontSizes();
  var fontWeights = messageRange.getFontWeights();
  var horizontalAlignments = messageRange.getHorizontalAlignments();
  var verticalAlignments = messageRange.getVerticalAlignments();
  //Build Text
  for (row=0;row<data.length;row++) {
    for (col=0;col<data[row].length;col++) {
      var cellText = data[row][col];
      var html = [];

        if (cellText instanceof Date) {
          cellText = Utilities.formatDate(
                      cellText,
                      ss.getSpreadsheetTimeZone(),
                      'MMM/d EEE');
          }
        var style = 'style="'
                + 'color: ' + fontColors[row][col] +'; '
                + 'font-family: ' + fontFamilies[row][col] +'; '
                + 'font-size: ' + fontSizes[row][col] +'px; '
                + 'font-weight: ' + fontWeights[row][col] +'; '
                + 'background-color: ' + backgrounds[row][col] +'; '
                + 'text-align: ' + horizontalAlignments[row][col] +'; '
                + 'vertical-align: ' + verticalAlignments[row][col] +'; '
                + 'display: inline"';
        html.push('<p ' + style + '>'
                    +cellText
                    +'</p>');

    }
    htmlMessage.push(html.join(''));
  }
  return htmlMessage;
}


// retrieves the Hyperlink in tab Links
function getHyperLink () {
  var messageRange = sheetLinks.getRange(2, 3, sheetLinks.getLastRow()-1, 1);

  // Read message contents
  var data = messageRange.getDisplayValues();

  // Get text styling
  var fontColors = messageRange.getFontColors();
  var backgrounds = messageRange.getBackgrounds();
  var fontFamilies = messageRange.getFontFamilies();
  var fontSizes = messageRange.getFontSizes();
  var fontWeights = messageRange.getFontWeights();
  var horizontalAlignments = messageRange.getHorizontalAlignments();
  var verticalAlignments = messageRange.getVerticalAlignments();

  //Build Text
  for (row=0;row<data.length;row++) {
    for (col=0;col<data[row].length;col++) {
      var cellText = data[row][col];
      var html = [];
      var link = sheetLinks.getRange(row+2, 3, 1, 1).getRichTextValue().getLinkUrl();

        if (cellText instanceof Date) {
          cellText = Utilities.formatDate(
                      cellText,
                      ss.getSpreadsheetTimeZone(),
                      'MMM/d EEE');
          }
        var style = 'style="'
                + 'color: ' + fontColors[row][col] +'; '
                + 'font-family: ' + fontFamilies[row][col] +'; '
                + 'font-size: ' + fontSizes[row][col] +'px; '
                + 'font-weight: ' + fontWeights[row][col] +'; '
                + 'background-color: ' + backgrounds[row][col] +'; '
                + 'text-align: ' + horizontalAlignments[row][col] +'; '
                + 'vertical-align: ' + verticalAlignments[row][col] +'; '
                + 'display: inline"';
        html.push('<p ' + style + '>'
                    +'<a href="'
                    +link
                    +'">'
                    +cellText
                    +'</p>');

    }
    htmlHyperLink.push(html.join(''));
  }
  return htmlHyperLink;
}


// retrieves the Table in tab Tables
function getTable(){
  var tableNumber = sheetTable.getRange("C1").getValue();
  var tableRange = sheetTable.getRange(3, 3, tableNumber, 1);
  var tableRangeValues = tableRange.getValues();

  for (var i=0; i<tableNumber; i++)  {
    var range = sheetTable.getRange(tableRangeValues[i][0]);
    startRow = range.getRow();
    startCol = range.getColumn();
    lastRow = range.getLastRow();
    lastCol = range.getLastColumn();
    
    // Read table contents
    var data = range.getDisplayValues();

    // Get css style attributes from range
    var fontColors = range.getFontColors();
    var backgrounds = range.getBackgrounds();
    var fontFamilies = range.getFontFamilies();
    var fontSizes = range.getFontSizes();
    var fontWeights = range.getFontWeights();
    var horizontalAlignments = range.getHorizontalAlignments();
    var verticalAlignments = range.getVerticalAlignments();

    var check = (lastRow-startRow)>0;
    if (check===true) {
    // Get column widths in pixels
    var colWidths = [];
    for (var col=startCol; col<=lastCol; col++) { 
      colWidths.push(sheetTable.getColumnWidth(col));
    }

    // Get row heights in pixels
    var rowHeights = [];
    for (var row=startRow; row<=lastRow; row++) { 
      rowHeights.push(sheetTable.getRowHeight(row));
    }
          // Build HTML Table, with inline styling for each cell
    var tableFormat = 'style="border:1px solid black;border-collapse:collapse;text-align:center" border=1 cellpadding=5';
    var html = ['<table '+tableFormat+'>'];

    // Column widths appear outside of table rows
    for (col=0;col<colWidths.length;col++) {
      html.push('<col width="'+colWidths[col]+'">')
    }

    // Populate rows
    for (row=0;row<data.length;row++) {
      html.push('<tr height="'+rowHeights[row]+'">');
      for (col=0;col<data[row].length;col++) {

        // Get formatted data
        var cellText = data[row][col];
        if (cellText instanceof Date) {
          cellText = Utilities.formatDate(
                      cellText,
                      ss.getSpreadsheetTimeZone(),
                      'MMM/d EEE');
        }
        var style = 'style="'
                  + 'color: ' + fontColors[row][col]+'; '
                  + 'font-family: ' + fontFamilies[row][col]+'; '
                  + 'font-size: ' + fontSizes[row][col]+'; '
                  + 'font-weight: ' + fontWeights[row][col]+'; '
                  + 'background-color: ' + backgrounds[row][col]+'; '
                  + 'text-align: ' + horizontalAlignments[row][col]+'; '
                  + 'vertical-align: ' + verticalAlignments[row][col]+'; '
                  +'"';
        html.push('<td ' + style + '>'
                  +cellText
                  +'</td>');
      }
      html.push('</tr>');
    }
    html.push('</table>');
    } else {
      var html = ['<br></br>'];
    }
    
    htmlTables.push(html.join(''))
  }
  return htmlTables;
}


// retrieves the Chart in tab Charts
function getChart(){
  var chartNumber = sheetCharts.getRange("C1").getValue();
  var chartRange = sheetCharts.getRange(3, 2, chartNumber, 1);
  var charts = sheetCharts.getCharts();

  for (var i=0; i<chartNumber; i++)  {
    for (var j=0; j<chartNumber; j++) {
      if (charts[i].getOptions().get('title') === chartRange.getValues()[j][0]) {
        htmlCharts.push(charts[i]);
      }
    }
  }
  return htmlCharts;
}


// checks if any triggers exists
function checkIfTriggerExists() {
  var ui = SpreadsheetApp.getUi();
  var triggers = ScriptApp.getProjectTriggers();
  var triggerName = [];

  for (var i = 0; i < triggers.length; i++) {
    triggerName.push(triggers[i].getHandlerFunction());
  }

  ui.alert('There are ' + triggers.length + ' triggers running (one per weekday).' + '\r\n' + triggerName.join("\r\n"));

}


// deletes all existing triggers
function deleteAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}
