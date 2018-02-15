//12:21PM, 2/8/18
/* This is the discreetFunctions code sheet. It's for functions that take in and put out data, like small processors. It's not for the main code flow. */

function apiCall(data) {
  Logger.log('Attempting an API call to ' + data.url + '.');
  if (data.apikey) { //if an api key is present, set the headers and use it
    Logger.log('API key provided, adding headers... (API key provided)')
    var APIheaders = {'X-Cisco-Meraki-API-Key': data.apikey}; //sets headers
    if (data.method) {
      Logger.log('Using specified method ' + data.method + '. (API key provided)'); //if a method is provided use it
      var options = {'contentType':'application/json', 'method':data.method, 'headers':APIheaders};
    } else {
      Logger.log('No method provided. Using GET. (API key provided)'); //if no method, use get
      var options = {'contentType':'application/json', 'method':'get', 'headers':APIheaders};
    }
  } else {
    Logger.log('No API key provided, skipping headers... (no API key)');
    if (data.method && data.payload) { //if a payload is provided, use it
      Logger.log('Using specified method ' + data.method + '. (no API key)');
      var options = {'contentType':'application/json', 'method':data.method, 'payload':data.payload};
    } else if (data.method && !data.payload) { //if a method and no payload is provided
      Logger.log('Using specified method ' + data.method + '. (no API key)');
      var options = {'contentType':'application/json', 'method':data.method};
    } else {
      Logger.log('No method provided. Using GET. (no API key)');
      var options = {'contentType':'application/json', 'method':'get'};
    }
  }
  var response = UrlFetchApp.fetch(data.url, options); //actual api call
  Logger.log('API call succeeded. Parsing responses.');
  var stringResponse = response.getContentText();
  var jsonResponse = JSON.parse(stringResponse); //parses response as json
  Logger.log('Completed API call to ' + data.url + '.');
  return {'jsonResponse':jsonResponse, 'stringResponse':stringResponse};
}

/*function apiCallPut(url, apikey) {
  Logger.log('Attempting an API call to ' + url + '.');
  var APIheaders = {'X-Cisco-Meraki-API-Key': apikey}; //sets headers
  var options = {'contentType':'application/json', 'method':'put', 'headers':APIheaders};
  var response = UrlFetchApp.fetch(url, options); //actual api call
  Logger.log('API call succeeded. Parsing responses.');
  var stringResponse = response.getContentText();
  var jsonResponse = JSON.parse(stringResponse); //parses response as json
  Logger.log('Completed API call to ' + url + '.');
  return;
  return {'jsonResponse':jsonResponse, 'stringResponse':stringResponse};
} //The only difference between the top and bottom functions is that apiCallPut is a PUT request whereas apiCall is a GET request.

function apiCallPost(url, payload) {
  Logger.log('Attempting an API call to ' + url + '.');
  var options = {'contentType':'application/json', 'method':'post', 'payload':JSON.stringify(payload), 'muteHttpExceptions':true};
  var response = UrlFetchApp.fetch(url, options); //actual api call
  Logger.log('API call succeeded. Parsing responses.');
  var stringResponse = response.getContentText();
  Logger.log(stringResponse);
  var jsonResponse = JSON.parse(stringResponse); //parses response as json
  Logger.log('Completed API call to ' + url + '.');
  return {'jsonResponse':jsonResponse, 'stringResponse':stringResponse};
} //This function takes in a payload and not an API key. It's for reporting errors to the mismatch API.
*/
/* Using the apiCall function:
the apiCall function is made so that you don't have to continuously copy and paste the code. Just set your API Key as a variable, set your method: GET, PUT, POST, etc. and
set your URL depending on what you need, and you're all set. It will return an object, from which you can get the response as a string and as JSON.
Just ask for result.stringResponse or result.jsonResponse.
Below is a function, that when called, logs the response as JSON and as a string. Uncomment it if you're interested.

function testAPICall() {

  var apikey = 'myapikey'; //set your API key. this can be set at the beginning of the document.

  //As you can see, once the API key is set, just set the URL and call the function. So, for each call, we only need two lines of code.
  var url = 'http://httpbin.org/get'; //set the API url.
  var method = 'get'; //set the method. this is not necessary because it defaults to get if no method is provided.
  var result = apiCall(url, method, apikey); //uses apiCall to make the actual api call, sending the URL and API key with the memberwise initalizer.

  //The below lines are not necessary, they just log the responses so you can see them.
  Logger.log(result.stringResponse); //Logs the response as a string.
  Logger.log(result.jsonResponse); //Logs the response as JSON.

} */

function getUserInfo() {
  try {
  //find user organization
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User data');

  var range = sheet.getRange('A2'); //grabs API Key
  var apikey = range.getDisplayValue();

  var range = sheet.getRange('B2'); //grabs Organization ID
  var organizationId = range.getDisplayValue();

  var range = sheet.getRange('C2'); //grabs Network ID
  var networkId = range.getDisplayValue();

  var range = sheet.getRange('D2'); //grabs security appliance serial
  var securityApplianceSerial = range.getDisplayValue();

  var range = sheet.getRange('E2'); //grabs timespan to list clients
  var clientTimespan = range.getDisplayValue();

  var range = sheet.getRange('F2'); //grabs client dashboard link
  var clientsURL = range.getDisplayValue();

  return {'apikey':apikey,'organizationId':organizationId,'networkId':networkId,'securityApplianceSerial':securityApplianceSerial,'clientTimespan':clientTimespan,'clientsURL':clientsURL};
  } catch(e) {
    var payload = {
       "id":"vGWK3gnQozAAjuCkU9ni7jH93yCutPRfsnU6HtaAn66gq4ekRtwGk9zTTYXgbbAk",
       "function":"getUserInfo",
       "fileName":e.fileName,
       "lineNumber":e.lineNumber,
       "message":e.message,
    };
    apiCall({'url':'https://api.mismatch.io/analytics/error', 'payload':payload, 'method':'post'});
    SpreadsheetApp.getUi().alert('I\'m sorry, something didn\'t work right. ' + 'I\'ve reported this to the developers. Here\'s the full error: ' + e.message);
  }
}
/* Using the getUserInfo function:
Grabs the user's API key, organization ID and network ID. Doesn't require any variables. */


function verifyInfoWithUser(dataToVerify, errorIfNotVerified) {
  try {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Is this correct?', dataToVerify , ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) {
    ui.alert(errorIfNotVerified);
    return;
  }
  Logger.log(dataToVerify + ' has been verified.');
  } catch(e) {
    var payload = {
       "id":"vGWK3gnQozAAjuCkU9ni7jH93yCutPRfsnU6HtaAn66gq4ekRtwGk9zTTYXgbbAk",
       "function":"connectToMeraki",
       "fileName":e.fileName,
       "lineNumber":e.lineNumber,
       "message":e.message,
    };
    apiCall({'url':'https://api.mismatch.io/analytics/error', 'payload':payload, 'method':'post'});
    SpreadsheetApp.getUi().alert('I\'m sorry, something didn\'t work right. ' + 'I\'ve reported this to the developers. Here\'s the full error: ' + e.message);
  }
}
/* Using the verifyInfoWithUser function:
This function takes in some dataToVerify and an errorIfNotVerified. It will prompt the user, and ask if dataToVerify is correct. If the user responds with
anything but a yes, it will throw up errorIfNotVerified. */

/* DEPRECATED DEPRECATED DEPRECATED
function resetSheet() {
  var sheet = SpreadsheetApp.getActiveSheet();

  switchSheets('Results');
  sheet.clearContents();

  var cell = sheet.getRange('A1');
  cell.setValue('Client name');
  var cell = sheet.getRange('B1');
  cell.setValue('Mac Address');
  var cell = sheet.getRange('C1');
  cell.setValue('Lan IP');
  var cell = sheet.getRange('D1');
  cell.setValue('Usage down/up in GB')
  var cell = sheet.getRange('E1');
  cell.setValue('Meraki dashboard URL')

}

 Using the resetSheet function:
This function is called directly from the menu, so nothing comes in and nothing is returned. It can also be called from inside of the code. */

function switchSheets(sheetName) {
 try {
 var newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
 if (newSheet == null || newSheet == undefined) {
   SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
 }
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  newSheet.activate();
  return newSheet;
  //newSheet.activate();
 } catch(e) {
    var payload = {
       "id":"vGWK3gnQozAAjuCkU9ni7jH93yCutPRfsnU6HtaAn66gq4ekRtwGk9zTTYXgbbAk",
       "function":"connectToMeraki",
       "fileName":e.fileName,
       "lineNumber":e.lineNumber,
       "message":e.message,
    };
   apiCall({'url':'https://api.mismatch.io/analytics/error', 'payload':payload, 'method':'post'});
    SpreadsheetApp.getUi().alert('I\'m sorry, something didn\'t work right. ' + 'I\'ve reported this to the developers. Here\'s the full error: ' + e.message);
  }
}

/* Using the switchSheets function:
This function makes a new sheet or switches to a sheet with a name you pass in via memberwise initalizer. If there's no sheet with the name you specify, it will
create a new sheet. It will return the current sheet object, so use a statement like:
var sheet = switchSheets('sheetName');
to make sure that you get the active sheet object. then, you'll be able to do operations like:
sheet.clear();
*/

function getApprovedClients() {
	try {
	var indexingSheetUrls = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Approved clients').getRange('A2:A' + SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Approved clients').getLastRow()).getValues(); //get the URLs of all the sheets to index for the approved clients list
  var indexingSheetNames = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Approved clients').getRange('B2:B' + SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Approved clients').getLastRow()).getValues(); //get the sheet names of all the sheets to index for the approved clients list
  var indexingSheetFirstCells = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Approved clients').getRange('C2:C' + SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Approved clients').getLastRow()).getValues(); //get the first cells of all the sheets to index for the approved clients list
  var approvedClients = [];

  for (var i = 0; i < indexingSheetUrls.length; i++) {
	var spreadSheet = SpreadsheetApp.openByUrl(indexingSheetUrls[i].join()); //open the i-st spreadsheet
    var sheet = spreadSheet.getSheetByName(indexingSheetNames[i].join()); //open the sheet inside of aforementioned spreadsheet
    approvedClients.push(sheet.getRange(indexingSheetFirstCells[i].join() + ':' + indexingSheetFirstCells[i].join().slice(0,1) + spreadSheet.getSheetByName(indexingSheetNames[i]).getLastRow()).getValues()); //add all of the mac addresses on that sheet to the approved clients variable
  }
  return approvedClients; //return our final product
} catch(e) { //I feel like this script might have a high chance for error, so I better add proper error reporting.
  var payload = {
     "id":"vGWK3gnQozAAjuCkU9ni7jH93yCutPRfsnU6HtaAn66gq4ekRtwGk9zTTYXgbbAk",
     "function":"getApprovedClients",
     "fileName":e.fileName,
     "lineNumber":e.lineNumber,
     "message":e.message,
  };
  apiCall({'url':'https://api.mismatch.io/analytics/error', 'payload':payload, 'method':'post'});
  SpreadsheetApp.getUi().alert('I\'m sorry, something didn\'t work right. ' + 'I\'ve reported this to the developers. Here\'s the full error: ' + e.message);
}
}

function initializeSpreadsheet() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Do you want to initalize this sheet?', 'That will completely erase every sheet on this document.', ui.ButtonSet.OK_CANCEL);
  if (response != ui.Button.OK) {
   ui.alert('Cancelling.');
   return;
  }
  Logger.log(SpreadsheetApp.getActiveSpreadsheet().getSheets());
  if (SpreadsheetApp.getActiveSpreadsheet().getSheetName() == 'Results' || SpreadsheetApp.getActiveSpreadsheet().getSheetName() == 'Approved clients' || SpreadsheetApp.getActiveSpreadsheet().getSheetName() == 'User data' || SpreadsheetApp.getActiveSpreadsheet().getSheetName() == 'Advanced output') {
    var response = ui.alert('You\'re already set up!', 'If you\'d like to re-initialize the sheet, press OK below.', ui.ButtonSet.OK_CANCEL);
    if (response != ui.Button.OK) {
      ui.alert('Cancelling.');
      return;
    } else {
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results').activate();
      var currentSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
      for (var i = 0; i < (currentSheets.length -= 1); i++)
        SpreadsheetApp.getActiveSpreadsheet().deleteSheet(currentSheets[i])
    }
  }


  var readingSpreadSheet = SpreadsheetApp.openById('1STQQAHvW9Re4vmFnHRnu6PVjX5TfdUFUVaH8jDba5LE');
  var initSheet = SpreadsheetApp.getActiveSpreadsheet();
  initSheet.getActiveSheet().setName('Results').getRange(readingSpreadSheet.getSheetByName('Results').getDataRange().getA1Notation()).setValues(readingSpreadSheet.getSheetByName('Results').getDataRange().getValues());
  initSheet.insertSheet('Approved clients').getRange(readingSpreadSheet.getSheetByName('Approved clients').getDataRange().getA1Notation()).setValues(readingSpreadSheet.getSheetByName('Approved clients').getDataRange().getValues());
  initSheet.getSheetByName('Approved clients').getRange('A2:D').setBackground('yellow');
  initSheet.insertSheet('User data').getRange(readingSpreadSheet.getSheetByName('User data').getDataRange().getA1Notation()).setValues(readingSpreadSheet.getSheetByName('User data').getDataRange().getValues());
  initSheet.getSheetByName('User data').getRange('A2:F2').setBackground('yellow');
  initSheet.insertSheet('Advanced output').getRange(readingSpreadSheet.getSheetByName('Advanced output').getDataRange().getA1Notation()).setValues(readingSpreadSheet.getSheetByName('Advanced output').getDataRange().getValues());
  initSheet.getSheetByName('Results').activate();
}
