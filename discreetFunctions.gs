//7:10PM, 3/22/18
/* This is the discreetFunctions code sheet. It's for functions that take in and put out data, like small processors. It's not for the main code flow. */

function apiCall(url, apikey) {
  Logger.log('Attempting an API call to ' + url + '.');
  var APIheaders = {'X-Cisco-Meraki-API-Key': apikey}; //sets headers
  var options = {'contentType':'application/json', 'method':'GET', 'headers':APIheaders, 'muteHttpExceptions':true};
  var response = UrlFetchApp.fetch(url, options); //actual api call
  if (response.getResponseCode() != 200) {
    return SpreadsheetApp.getUi().alert('Something wasn\'t right.', 'I tried to contact Meraki, but I got a ' + response.getResponseCode() + ' response code. If you got a 404, either your security appliance serial number, API key, organization ID or network ID isn\'t correct.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
  Logger.log('API call succeeded. Parsing responses.');
  var stringResponse = response.getContentText();
  var jsonResponse = JSON.parse(stringResponse); //parses response as json
  Logger.log('Completed API call to ' + url + '.');
  return {'jsonResponse':jsonResponse, 'stringResponse':stringResponse};
  
}

function apiCallPut(url, apikey) {
  Logger.log('Attempting an API call to ' + url + '.');
  var APIheaders = {'X-Cisco-Meraki-API-Key': apikey}; //sets headers
  var options = {'contentType':'application/json', 'method':'put', 'headers':APIheaders, 'muteHttpExceptions':true};
  var response = UrlFetchApp.fetch(url, options); //actual api call
  if (response.getResponseCode() != 200) {
    return SpreadsheetApp.getUi().alert('Something wasn\'t right.', 'I tried to contact Meraki, but I got a ' + response.getResponseCode() + ' response code. If you got a 404, either your security appliance serial number, API key, organization ID or network ID isn\'t correct.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
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
  if (response.getResponseCode() != 200) {
    return SpreadsheetApp.getUi().alert('Something wasn\'t right.', 'I tried to contact Meraki, but I got a ' + response.getResponseCode() + ' response code. If you got a 404, either your security appliance serial number, API key, organization ID or network ID isn\'t correct.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
  Logger.log('API call succeeded. Parsing responses.');
  var stringResponse = response.getContentText();
  Logger.log(stringResponse);
  var jsonResponse = JSON.parse(stringResponse); //parses response as json
  Logger.log('Completed API call to ' + url + '.');
  return {'jsonResponse':jsonResponse, 'stringResponse':stringResponse};
} //This function takes in a payload and not an API key. It's for reporting errors to the mismatch API.

/* Using the apiCall function:
the apiCall function is made so that you don't have to continuously copy and paste the code. Just set your API Key as a variable, and set your URL depending on what you need,
and you're all set. It will return an object, from which you can get the response as a string and as JSON. Just ask for result.stringResponse or result.jsonResponse.
Below is a function, that when called, logs the response as JSON and as a string. Uncomment it if you're interested.

function testAPICall() {

  var apikey = 'myapikey'; //set your API key. this can be set at the beginning of the document.

  //As you can see, once the API key is set, just set the URL and call the function. So, for each call, we only need two lines of code.
  var url = 'http://httpbin.org/get'; //set the API url.
  var result = apiCall(url, apikey); //uses apiCall to make the actual api call, sending the URL and API key with the memberwise initalizer.

  //The below lines are not necessary, they just log the responses so you can see them.
  Logger.log(result.stringResponse); //Logs the response as a string.
  Logger.log(result.jsonResponse); //Logs the response as JSON.

} */

function getUserInfo() {
  try {
    Logger.log('begin get user info');
  var ui = SpreadsheetApp.getUi();
  //find user organization
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User data');
  if (!sheet) return ui.alert('Can\'t find User data sheet.', 'Please make sure you initialised your sheet. Try going to Add-ons, MerakiBlocki, Advanced, Initialize spreadsheet.', ui.ButtonSet.OK);
  if (!sheet.getRange('A2')) {
  	return ui.alert('Your API key is missing.', 'Please check your User data sheet. If there\'s no User data sheet, try initializing your sheet from Add-ons, MerakiBlocki, Advanced, Initialize spreadsheet', ui.ButtonSet.OK);
  } else {
	 var range = sheet.getRange('A2'); //grabs API Key
	 var apikey = range.getDisplayValue();
  }

  var range = sheet.getRange('B2'); //grabs Organization ID
  var organizationId = range.getDisplayValue();
  if (!organizationId) return ui.alert('Your Organization ID is missing.', 'Please check your User data sheet. If there\'s no User data sheet, try initializing your sheet from Add-ons, MerakiBlocki, Advanced, Initialize spreadsheet', ui.ButtonSet.OK);

  var range = sheet.getRange('C2'); //grabs Network ID
  var networkId = range.getDisplayValue();
  if (!networkId) return ui.alert('Your Network ID is missing.', 'Please check your User data sheet. If there\'s no User data sheet, try initializing your sheet from Add-ons, MerakiBlocki, Advanced, Initialize spreadsheet', ui.ButtonSet.OK);

  var range = sheet.getRange('D2'); //grabs security appliance serial
  var securityApplianceSerial = range.getDisplayValue();
  if (!securityApplianceSerial) return ui.alert('Your Security Appliance serial number is missing.', 'Please check your User data sheet. If there\'s no User data sheet, try initializing your sheet from Add-ons, MerakiBlocki, Advanced, Initialize spreadsheet', ui.ButtonSet.OK);

  var range = sheet.getRange('E2'); //grabs timespan to list clients
  var clientTimespan = range.getDisplayValue();
  if (!clientTimespan) return ui.alert('Your client timespan is missing.', 'Please check your User data sheet. If there\'s no User data sheet, try initializing your sheet from Add-ons, MerakiBlocki, Advanced, Initialize spreadsheet', ui.ButtonSet.OK);
  
  var range = sheet.getRange('F2'); //grabs client dashboard link
  var clientsURL = range.getDisplayValue();
  if (!clientsURL) return ui.alert('Your Meraki Dashboard link is missing.', 'Please check your User data sheet. If there\'s no User data sheet, try initializing your sheet from Add-ons, MerakiBlocki, Advanced, Initialize spreadsheet', ui.ButtonSet.OK);

  var range = sheet.getRange('G2'); //grabs user license Key
    Logger.log('log1');
  var licenseKey = range.getDisplayValue();
    Logger.log('log2');
  if (!licenseKey) return ui.alert('Your license key is missing.', 'Please check your User data sheet. If there\'s no User data sheet, try initializing your sheet from Add-ons, MerakiBlocki, Advanced, Initialize spreadsheet', ui.ButtonSet.OK);
  
  var range = sheet.getRange('H2'); //grabs user license email
  var licenseEmail = range.getDisplayValue();
  if (!licenseEmail) return ui.alert('Your license email is missing.', 'Please check your User data sheet. If there\'s no User data sheet, try initializing your sheet from Add-ons, MerakiBlocki, Advanced, Initialize spreadsheet', ui.ButtonSet.OK);
  
  var verificationResponse = apiCall('https://api.mismatch.io/licensing/verify?licenseKey=' + licenseKey + '&app=merakiApp&email=' + licenseEmail, 'noAPIKeyNeeded').jsonResponse;
  var response = verificationResponse[0];
  if (!response) {
      Logger.log(verificationResponse.licenseType);
      Logger.log(verificationResponse.email);
      var licenseValidity = false;
      var licenseMaxClients = -2;
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User data').getRange("G3").setValue('Invalid license.');
      return ui.alert('Invalid license.', verificationResponse.message, ui.ButtonSet.OK);
  } else {
    if (response.licenseType == 'basic' && response.email == licenseEmail) {
      var licenseValidity = true;
      var licenseMaxClients = response.licenseMaxClients;
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User data').getRange("G3").setValue('Valid - ' + response.licenseType + ' license. Thank you for playing fair.');
      Logger.log('BASIC LICENSE');
    } else if (response.licenseType == 'pro' && response.email == licenseEmail) {
      var licenseValidity = true;
      var licenseMaxClients = response.licenseMaxClients;
      Logger.log('PRO LICENSE');
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User data').getRange("G3").setValue('Valid - ' + response.licenseType + ' license. Thank you for playing fair.');
    } else if (response.licenseType == 'unlimited' && response.email == licenseEmail) {
      var licenseValidity = true;
      var licenseMaxClients = response.licenseMaxClients;
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User data').getRange("G3").setValue('Valid - ' + response.licenseType + ' license. Thank you for playing fair.');
      Logger.log('UNLIMITED LICENSE');
    } else if (response.licenseType == 'expired') {
      var licenseValidity = false;
      var licenseMaxClients = -1;
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User data').getRange("G3").setValue('Expired license.');
      Logger.log('EXPIRED LICENSE');
    } else {
      //Logger.log(response.licenseType);
      //Logger.log(response.email);
      var licenseValidity = false;
      var licenseMaxClients = -2;
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User data').getRange("G3").setValue('Invalid license.');
      ui.alert('Something\'s wrong with your license.', 'It wasn\'t specifically expired, so it\'s very possible that your license and email don\'t match. Or it\'s possible your key doesn\'t exist. Check them and try again.', ui.ButtonSet.OK);
    }
  }
    Logger.log('end get user info');
    return {'apikey':apikey,'organizationId':organizationId,'networkId':networkId,'securityApplianceSerial':securityApplianceSerial,'clientTimespan':clientTimespan,'clientsURL':clientsURL,'licenseValidity':licenseValidity,'licenseMaxClients':licenseMaxClients,'licenseType':response.licenseType};
  } catch(e) {
    var payload = {
       "id":"vGWK3gnQozAAjuCkU9ni7jH93yCutPRfsnU6HtaAn66gq4ekRtwGk9zTTYXgbbAk",
       "function":"getUserInfo",
       "fileName":e.fileName,
       "lineNumber":e.lineNumber,
       "message":e.message,
    };
    apiCallPost('https://api.mismatch.io/analytics/error', payload);
    SpreadsheetApp.getUi().alert('I\'m sorry, something didn\'t work right. ' + 'I\'ve reported this to the developers. Here\'s the full error: ' + e.message);
  }
}
/* Using the getUserInfo function:
Grabs user data from User info sheet. Doesn't require any memberwise inits. */


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
    apiCallPost('https://api.mismatch.io/analytics/error', payload);
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
    apiCallPost('https://api.mismatch.io/analytics/error', payload);
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
      Logger.log('starting approved clients scan');
	var indexingSheetUrls = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Approved clients').getRange('A2:A' + SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Approved clients').getLastRow()).getValues(); //get the URLs of all the sheets to index for the approved clients list
  var indexingSheetNames = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Approved clients').getRange('B2:B' + SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Approved clients').getLastRow()).getValues(); //get the sheet names of all the sheets to index for the approved clients list
  var indexingSheetFirstCells = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Approved clients').getRange('C2:C' + SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Approved clients').getLastRow()).getValues(); //get the first cells of all the sheets to index for the approved clients list
  var approvedClients = [];
  var approvedTest = [];

  for (var i = 0; i < indexingSheetUrls.length; i++) {
	var spreadSheet = SpreadsheetApp.openByUrl(indexingSheetUrls[i].join()); //open the i-st spreadsheet
    var sheet = spreadSheet.getSheetByName(indexingSheetNames[i].join()); //open the sheet inside of aforementioned spreadsheet
    approvedClients.push(sheet.getRange(indexingSheetFirstCells[i].join() + ':' + indexingSheetFirstCells[i].join().slice(0,1) + spreadSheet.getSheetByName(indexingSheetNames[i]).getLastRow()).getValues()); //add all of the mac addresses on that sheet to the approved clients variable
  }
      Logger.log('approved clients scan done.');
  return approvedClients; //retunr our final product
} catch(e) { //I feel like this script might have a high chance for error, so I better add proper error reporting.
  var payload = {
     "id":"vGWK3gnQozAAjuCkU9ni7jH93yCutPRfsnU6HtaAn66gq4ekRtwGk9zTTYXgbbAk",
     "function":"getApprovedClients",
     "fileName":e.fileName,
     "lineNumber":e.lineNumber,
     "message":e.message,
  };
  apiCallPost('https://api.mismatch.io/analytics/error', payload);
  SpreadsheetApp.getUi().alert('I\'m sorry, something didn\'t work right. ' + 'I\'ve reported this to the developers. Here\'s the full error: ' + e.message);
}
}
//[[ [MAC], [MAC], [MAC] ]]

function initializeSpreadsheet() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Do you want to initalize this sheet?', 'That will completely erase every sheet on this document.', ui.ButtonSet.OK_CANCEL);
  if (response != ui.Button.OK) {
   return;
  }
  Logger.log(SpreadsheetApp.getActiveSpreadsheet().getSheets());
  if (SpreadsheetApp.getActiveSpreadsheet().getSheetName() == 'Results' || SpreadsheetApp.getActiveSpreadsheet().getSheetName() == 'Approved clients' || SpreadsheetApp.getActiveSpreadsheet().getSheetName() == 'User data' || SpreadsheetApp.getActiveSpreadsheet().getSheetName() == 'Advanced output') {
    var response = ui.alert('You\'re already set up!', 'If you\'d like to re-initialize the sheet, press OK below.', ui.ButtonSet.OK_CANCEL);
    if (response != ui.Button.OK) {
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
