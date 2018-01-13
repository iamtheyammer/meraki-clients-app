//9:20PM 1/12/18
/* This is the discreetFunctions code sheet. It's for functions that take in and put out data, like small processors. It's not for the main code flow. */

function apiCall(url, apikey) {
  var APIheaders = {'X-Cisco-Meraki-API-Key': apikey};
  var options = {'contentType':'application/json', 'method':'GET', 'headers':APIheaders};
  var response = UrlFetchApp.fetch(url, options);
  var stringResponse = response.getContentText();
  var jsonResponse = JSON.parse(stringResponse);
  return {'jsonResponse':jsonResponse, 'stringResponse':stringResponse};
}

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

function getUserInfo(apikey) {
  //find user organization
  var api = apiCall('https://api.meraki.com/api/v0/organizations/', apikey);
  temp = api.stringResponse;
  var userOrganization = temp.slice(7, 25);
  //find user name
  var userName = api.jsonResponse[0].name;
  
  return {'userOrganization':userOrganization,'userName':userName}; 
}
/* Using the getUserInfo function:
It simply just grabs the name and the organization of the user. Doesn't require any variables. */


function verifyInfoWithUser(dataToVerify, errorIfNotVerified) {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Is this correct?', dataToVerify , ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) {
    ui.alert(errorIfNotVerified);
    return;
  }
  Logger.log(dataToVerify + ' has been verified.');
}
/* Using the verifyInfoWithUser function:
This function takes in some dataToVerify and an errorIfNotVerified. It will prompt the user, and ask if dataToVerify is correct. If the user responds with 
anything but a yes, it will throw up errorIfNotVerified. */
//https://api.meraki.com/api/v0/devices/VRT-2207617868457/clients?timespan=86400 http://httpbin.org/get

function resetSheet() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  sheet.clearContents();
  
  var cell = sheet.getRange('A1');
  cell.setValue('Device Serial');
  var cell = sheet.getRange('B1');
  cell.setValue('Mac Address');
  var cell = sheet.getRange('C1');
  cell.setValue('Device Model');
  var cell = sheet.getRange('D1');
  cell.setValue('Lan IP');
  var cell = sheet.getRange('F1');
  cell.setValue('Networks:')

}

/* Using the resetSheet function:
This function is called directly from the menu, so nothing comes in and nothing is returned. It can also be called from inside of the code. */

function switchSheets(sheetName) {
 var newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
 if (newSheet == null || newSheet == undefined) {
   SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
 }
  Logger.log(newSheet)
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  newSheet.activate();
  return newSheet;
  //newSheet.activate();
}

/* Using the switchSheets function:
This function makes a new sheet or switches to a sheet with a name you pass in via memberwise initalizer. If there's no sheet with the name you specify, it will
create a new sheet. It will return the current sheet object, so use a statement like:
var sheet = switchSheets('sheetName');
to make sure that you get the active sheet object. then, you'll be able to do operations like:
sheet.clear();
*/

function completelyClearSheet() {
 SpreadsheetApp.getActiveSheet().clear();
}
/* Using the completelyClearSheet() function:
You'll probably never need to use this because you can just sheet.clear();. It's just for the advanced functions menu. */