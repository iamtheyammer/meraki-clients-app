//9:12PM, 1/31/18
function printOrganizations() {

  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var cell;
  var range;
  
  apiCallPut('https://api.mismatch.io/analytics?id=vGWK3gnQozAAjuCkU9ni7jH93yCutPRfsnU6HtaAn66gq4ekRtwGk9zTTYXgbbAk&function=printOrganizations', 'noApiKeyNeeded'); //analytics
  
  var userData = getUserInfo();
  
  var apikey = userData.apikey;
  if (apikey.length <= 20) {ui.alert('Your API key is missing or too short.'); return;} //get (^) and verify api key
  
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Advanced output").activate(); //visually switch to advanced output
  sheet.clear();
  var userOrganizations = apiCall('https://api.meraki.com/api/v0/organizations/', apikey); //make api call
  
  range = sheet.getRange("A1:B1"); //print heading
  cell = sheet.setActiveRange(range);
  cell.setValues([['Name', 'Organization ID']]);
  
  var numberOfOrganizations = userOrganizations.jsonResponse.length;
  
  for (var i = 0; i < numberOfOrganizations; i++) { //print result
    range = sheet.getRange("A" + (i+2) + ":B" + (i+2));
    cell = sheet.setActiveRange(range);
    cell.setValues([[userOrganizations.jsonResponse[i].name, userOrganizations.jsonResponse[i].id]]);
  }
}
  
function printNetworks() {
  
  
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var cell;
  var range;
  
  apiCallPut('https://api.mismatch.io/analytics?id=vGWK3gnQozAAjuCkU9ni7jH93yCutPRfsnU6HtaAn66gq4ekRtwGk9zTTYXgbbAk&function=printNetworks', 'noApiKeyNeeded'); //analytics
  
  var userData = getUserInfo();
  
  var apikey = userData.apikey;
  if (apikey.length <= 20) {ui.alert('Your API key is missing or too short.'); return;} //get (^) and verify api key
  
  var organizationId = userData.organizationId;
  if (organizationId.length <= 1) {ui.alert('Your Organization ID is missing or too short.'); return;} //get (^) and verify organization id
  
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Advanced output").activate(); //visually switch to advanced output
  sheet.clear();
  
  var networkList = apiCall('https://api.meraki.com/api/v0/organizations/' + organizationId + '/networks', apikey); //make api call
  var numberOfNetworks = networkList.jsonResponse.length;
  
  range = sheet.getRange("A1:B1"); //print heading
  cell = sheet.setActiveRange(range);
  cell.setValues([['Name', 'Network ID']]);
  
  for (var i = 0; i < numberOfNetworks; i++) { //print result
    range = sheet.getRange("A" + (i+2) + ":B" + (i+2));
    cell = sheet.setActiveRange(range);
    cell.setValues([[networkList.jsonResponse[i].name, networkList.jsonResponse[i].id]]);
  }
}

function customAPICall() {
 
  var ui = SpreadsheetApp.getUi();
  
  apiCallPut('https://api.mismatch.io/analytics?id=vGWK3gnQozAAjuCkU9ni7jH93yCutPRfsnU6HtaAn66gq4ekRtwGk9zTTYXgbbAk&function=customAPICall', 'noApiKeyNeeded'); //analytics
  
  var userData = getUserInfo();
  var apikey = userData.apikey;
  if (apikey.length <= 20) {ui.alert('Your API key is missing or too short.'); return;} //get (^) and verify api key
  
  var response = ui.prompt('What is the URL you want to fetch?', 'Enter the entire URL, including https:// and the domain.', ui.ButtonSet.OK_CANCEL); //ask user for url
  if (response.getSelectedButton() !== ui.Button.OK) {
   ui.alert('The user chose to close the dialog.'); 
   return;
  }
   var url = response.getResponseText();

  var sheet = switchSheets('Advanced output');
  var apiResult = apiCall(url, apikey); //make api call
  
  sheet.clear(); //print result
  range = sheet.getRange("A1");
  cell = sheet.setActiveRange(range);
  cell.setValue(['Printing response from ' + url])
  range = sheet.getRange("A2")
  cell = sheet.setActiveRange(range)
  cell.setValue([apiResult.stringResponse]);
}

function customAPICallPut() {
 
  var ui = SpreadsheetApp.getUi();
  
  apiCallPut('https://api.mismatch.io/analytics?id=vGWK3gnQozAAjuCkU9ni7jH93yCutPRfsnU6HtaAn66gq4ekRtwGk9zTTYXgbbAk&function=customAPICallPut', 'noApiKeyNeeded'); //analytics
  
  var userData = getUserInfo();
  var apikey = userData.apikey;
  if (apikey.length <= 20) {ui.alert('Your API key is missing or too short.'); return;} //get (^) and verify api key
  
  var response = ui.prompt('What is the URL you want to fetch?', 'Enter the entire URL, including https:// and the domain.', ui.ButtonSet.OK_CANCEL); //ask user for url
  if (response.getSelectedButton() !== ui.Button.OK) {
   ui.alert('The user chose to close the dialog.'); 
   return;
  }
   var url = response.getResponseText();

  var sheet = switchSheets('Advanced output'); //visually switch to advanced output
  var apiResult = apiCallPut(url, apikey); //make api call
  
  sheet.clear(); //print result
  range = sheet.getRange("A1");
  cell = sheet.setActiveRange(range);
  cell.setValue(['Printing response from ' + url])
  range = sheet.getRange("A2")
  cell = sheet.setActiveRange(range)
  cell.setValue([apiResult.stringResponse]);
}

function unblockClients() {
 
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var cell;
  var range;
  
  apiCallPut('https://api.mismatch.io/analytics?id=vGWK3gnQozAAjuCkU9ni7jH93yCutPRfsnU6HtaAn66gq4ekRtwGk9zTTYXgbbAk&function=unblockClients', 'noApiKeyNeeded'); //analytics
  
  var userData = getUserInfo();
  var apikey = userData.apikey;
  if (apikey.length <= 20) {ui.alert('Your API key is missing or too short.'); return;} //get (^) and verify api key
  
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results").activate(); //visually switch to results sheet
  var unknownClients = sheet.getRange('B2:B' + sheet.getLastRow()).getValues(); //grab mac addresses from there
  
  var response = ui.alert('Are you sure you want to unblock all clients listed on this sheet?', 'You can press no below to remove clients you don\'t want to unblock.' , ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) {
    ui.alert('Cancelling.');
    return;
  }
  
  sheet.getRange('F1').setValue('Device policy'); 
  
  for (i = 0; i < unknownClients.length; i++) {
  Logger.log('Attempting to block ' + unknownClients[i] + 'from the network...');
  //var unknownClientURI = encodeURIComponent(unknownClients[i]); //encodes mac address but not used, works as is
  var response = apiCallPut('https://n126.meraki.com/api/v0/networks/' + userData.networkId + '/clients/' + unknownClients[i] + '/policy?timespan=2592000&devicePolicy=normal', apikey); //make api call 
  range = sheet.getRange("F" + (i+2) + ":F" + (i+2)); //print that it's done
  cell = sheet.setActiveRange(range);
  cell.setValue([['Device policy set to Normal']]);
  Logger.log('Successfully allowed ' + unknownClients[i] + ' onto the network.');
  Logger.log(response);
  Utilities.sleep(400); //wait 400 milliseconds to comply with meraki's 5 calls/second limit
  }
}

function completelyClearSheet() {
 SpreadsheetApp.getActiveSheet().clear();
}