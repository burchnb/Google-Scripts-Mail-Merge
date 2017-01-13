var finalMsg="";
var finalCustRows=[{}];

function getObj($a){
  var req=$a;
  var sheet=SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet=sheet.getSheetByName("List");  //Change "List" to the name of your customer list sheet.
  var rows=mainSheet.getDataRange();
  var numRows=rows.getLastRow();
  var numCols=rows.getLastColumn();
  var ui=SpreadsheetApp.getUi();
  switch(req){
    case "sheet":return sheet;break;
    case "mainSheet":return mainSheet;break;
    case "rows":return rows;break;
    case "numRows":return numRows;break;
    case "numCols":return numCols;break;
    case "ui":return ui;break;
    case "nameColumn":return "B";break;
    default:return "unknown";
  }
}

function chooseMessage(){
  var sheets=getObj("sheet").getSheets();
  var emailSheets=[{}];
  var inc=0;
  for(var i=0;i<sheets.length;i++){
    if(sheets[i].getName().indexOf("(eml)")>-1){
      emailSheets[inc]=sheets[i].getName();
      inc++;
    }
  }
  var message="Which message do you want to send?\nEnter the number that corresponds to the message you want to send.\n\n";
  var num=0;
  for(var i=0;i<emailSheets.length;i++){
    num=i+1;
    message=message+num+".  -"+emailSheets[i]+"\n";
  }
  var input=getObj("ui").prompt(message,getObj("ui").ButtonSet.OK_CANCEL);
  if(input.getSelectedButton()==getObj("ui").Button.CANCEL) return; else Logger.log("input is "+input.getResponseText());
  var ans=input.getResponseText();
  if(isNaN(ans) || ans>emailSheets.length || ans<1){
    getObj("ui").alert("Your input is invalid.  Please enter the number that corresponds to the message you want to send. ");chooseMessage();
  }
  else{
    var confirmInput=getObj("ui").alert("Confirm Selection","You have selected this message:\n\n"+ans+".  -"+emailSheets[ans-1]+"\n\nIs this correct?",getObj("ui").ButtonSet.YES_NO);
    confirmInput==getObj("ui").Button.YES ? Logger.log("continuing...") : Logger.log("Cancelled...") && chooseMessage();
  }
  return emailSheets[ans-1];
}

function myFunction() {
  var customersToEmail=[{}];
  var messageToEmail="";
  var errorCheck=[{}];
  var input=getObj("ui").prompt("Which Row Numbers?",getObj("ui").ButtonSet.OK_CANCEL);
  if(input.getSelectedButton()==getObj("ui").Button.CANCEL) return; else Logger.log("input is "+input.getResponseText());
  var inf=input.getResponseText();
  errorCheck=checkInput(inf);
  Logger.log("errorCheck length is "+errorCheck.length+" and errorcheck0 is "+errorCheck[0]);
  if(errorCheck[0] != "nothing"){
     var errorMsg="";
     for(var i=0;i<errorCheck.length;i++){
       errorMsg=errorMsg+errorCheck[i]+"\n";
     }
    getObj("ui").alert("The following are invalid input(s):\n"+errorMsg+"\nPlease try again.");
    myFunction();
  }
  else{
    if(inf.indexOf(",")>-1)
      customersToEmail=splitCommas(inf);
    else if(inf.indexOf(":")>-1)
      customersToEmail=splitColon(inf);
    else customersToEmail=normalGet(inf);
      messageToEmail=displayResults(customersToEmail);
  }
  Logger.log("Message to email - "+messageToEmail);
  Logger.log("Customer(s) to email");
  for(var i=0;i<customersToEmail.length;i++){
    Logger.log(customersToEmail[i]);
  }
  finalMsg=messageToEmail;
  var inc=0;
  for(var i=0;i<customersToEmail.length;i++){
    Logger.log("i%2 is "+i%2);
    if(i%2==0){
    finalCustRows[inc]=customersToEmail[i];
      Logger.log(finalCustRows[inc]);
      inc++;
    }
  }
  sendEmails();
}

function checkInput($a){
  var input=$a;
  var arr=[{}];
  var errors=[{}];
  arr=input.split(/[":",]+/);
  var inc=0;
  for(var i=0;i<arr.length;i++){
    arr[i]<2||arr[i]>getObj("numRows") ? errors[inc]=arr[i] && inc++ : errors[inc]="nothing";
  }
  return errors;
}

function displayResults($a){
  var rowAndName=$a;
  var confirm="";
  var message="";
  for(var i=0;i<rowAndName.length;i++){
    confirm=confirm+"Row: "+rowAndName[i]+"......";
    i++;
    confirm=confirm+"   Name: "+rowAndName[i]+"\n";
  }
  var confirmInput=getObj("ui").alert("Confirm Selection",confirm+"\n\nIs this correct?",getObj("ui").ButtonSet.YES_NO);
  if(confirmInput==getObj("ui").Button.YES){
    Logger.log("continuing...");
    message=chooseMessage();
  }
  else{
    Logger.log("Cancelled...");
    myFunction();
  }
  return message;
}

function colNumToLetter(){
  var dividend=getObj("numCols");
  var columnName = "";
  var modulo;
  while(dividend>0){
    modulo=(dividend-1) % 26;
    columnName=String.fromCharCode(65+modulo).toString()+columnName;
    dividend=parseInt((dividend-modulo) / 26);
  }
  return columnName;
}

function normalGet($a) {
  var inf=$a;
  var rowAndName=[{}];
  Logger.log(getObj("nameColumn")+inf);
  rowAndName[0]=inf;
  rowAndName[1]=getObj("mainSheet").getRange(getObj("nameColumn")+inf).getDisplayValue();
  return rowAndName;
}

function splitCommas($a) {
  //Logger.log("splitCommas");
  var inf=$a;
  var arr=[{}];
  var rowAndName=[{}];
  arr=inf.split(",");
  var inc=0;
  for(var i=0;i<arr.length;i++){
    if(arr[i].indexOf(":")>-1){
      var arr2=[{}];
      arr2=arr[i].split(":");
      var k=arr2[0]-1;
      do{
        k++;
        rowAndName[inc]=k.toString();
        inc++;
        rowAndName[inc]=getObj("mainSheet").getRange(getObj("nameColumn")+k).getDisplayValue();
        inc++;
      }while(k<arr2[1]);
      //rowAndName=splitColon(arr[i]);
    }
    else{
      rowAndName[inc]=arr[i].toString();
      inc++;
      rowAndName[inc]=getObj("mainSheet").getRange(getObj("nameColumn")+arr[i]).getDisplayValue();
      inc++;
    }
  }
  return rowAndName;
}

function splitColon($a) {
  //Logger.log("splitColon");
  var inf=$a;
  var arr=[{}];
  var rowAndName=[{}];
  arr=inf.split(":");
  var inc=0;
  var k=arr[0]-1;
  do{
    k++;
    rowAndName[inc]=k.toString();
    inc++;
    rowAndName[inc]=getObj("mainSheet").getRange(getObj("nameColumn")+k).getDisplayValue();
    inc++;
  }while(k<arr[1]);
  return rowAndName;
}

/**************************************
// COPIED
//
//
//
//
**************************************/
function sendEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = getObj("mainSheet");
  var dataRange = dataSheet.getRange(2, 1, dataSheet.getMaxRows() - 1, getObj("numCols"));

  var templateSheet = ss.getSheetByName(finalMsg);
  var emailTemplate = templateSheet.getRange("A1").getValue();

  // Create one JavaScript object per row of data.
  var objects = getRowsData(dataSheet, dataRange);

  // For every row object, create a personalized email from a template and send
  // it to the appropriate person.
  for (var i = 0; i < objects.length; ++i) {
    // Get a row object
    var rowData = objects[i];

    // Generate a personalized email.
    // Given a template string, replace markers (for instance ${"First Name"}) with
    // the corresponding value in a row object (for instance rowData.firstName).
    var emailText = fillInTemplateFromObject(emailTemplate, rowData);
    var emailSubject = "Lunch Heroes";

    MailApp.sendEmail(rowData.emailAddress, emailSubject, emailText);
  }
}


// Replaces markers in a template string with values define in a JavaScript data object.
// Arguments:
//   - template: string containing markers, for instance ${"Column name"}
//   - data: JavaScript object with values to that will replace markers. For instance
//           data.columnName will replace marker ${"Column name"}
// Returns a string without markers. If no data is found to replace a marker, it is
// simply removed.
function fillInTemplateFromObject(template, data) {
  var email = template;
  // Search for all the variables to be replaced, for instance ${"Column name"}
  var templateVars = template.match(/\$\{\"[^\"]+\"\}/g);

  // Replace variables from the template with the actual values from the data object.
  // If no value is available, replace with the empty string.
  for (var i = 0; i < templateVars.length; ++i) {
    // normalizeHeader ignores ${"} so we can call it directly here.
    var variableData = data[normalizeHeader(templateVars[i])];
    email = email.replace(templateVars[i], variableData || "");
  }

  return email;
}





//////////////////////////////////////////////////////////////////////////////////////////
//
// The code below is reused from the 'Reading Spreadsheet data using JavaScript Objects'
// tutorial.
//
//////////////////////////////////////////////////////////////////////////////////////////

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), normalizeHeaders(headers));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  var inc=0;
  for (var i = 0; i < data.length; ++i) {
    if(i+2==finalCustRows[inc]){
      Logger.log("i = "+i+"and fcr = "+finalCustRows[inc]);
      inc++;
      var object = {};
      var hasData = false;
      for (var j = 0; j < data[i].length; ++j) {
        Logger.log(data[i][j]);
        var cellData = data[i][j];
        if (isCellEmpty(cellData)) {
          continue;
        }
        object[keys[j]] = cellData;
        hasData = true;
      }
      if (hasData) {
        objects.push(object);
      }
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  Logger.log("headers length is "+headers.length);
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
    Logger.log("arr "+keys[i]);
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  Logger.log(key);
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the doMerge() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Start Merge",
    functionName : "myFunction"
  }];
  spreadsheet.addMenu("Merge", entries);
};
