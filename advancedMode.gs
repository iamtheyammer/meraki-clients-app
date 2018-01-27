//9:33PM, 1/26/18
function printOrganizations() {

  
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var cell;
  var range;
  var userData = getUserInfo();
  
  var apikey = userData.apikey;
  if (apikey.length <= 20) {ui.alert('Your API key is missing or too short.'); return;}
  
  switchSheets("Advanced output");
  sheet.clear();
  var userOrganizations = apiCall('https://api.meraki.com/api/v0/organizations/', apikey);
  
  range = sheet.getRange("A1:B1");
  cell = sheet.setActiveRange(range);
  cell.setValues([['Name', 'Organization ID']]);
  
  var numberOfOrganizations = userOrganizations.jsonResponse.length;
  
  for (var i = 0; i < numberOfOrganizations; i++) {
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
  var userData = getUserInfo();
  
  var apikey = userData.apikey;
  if (apikey.length <= 20) {ui.alert('Your API key is missing or too short.'); return;}
  
  var organizationId = userData.organizationId;
  if (organizationId.length <= 1) {ui.alert('Your Organization ID is missing or too short.'); return;}
  
  switchSheets('Advanced output');
  sheet.clear();
  
  var networkList = apiCall('https://api.meraki.com/api/v0/organizations/' + organizationId + '/networks', apikey);
  var numberOfNetworks = networkList.jsonResponse.length;
  
  range = sheet.getRange("A1:B1");
  cell = sheet.setActiveRange(range);
  cell.setValues([['Name', 'Network ID']]);
  
  for (var i = 0; i < numberOfNetworks; i++) {
    range = sheet.getRange("A" + (i+2) + ":B" + (i+2));
    cell = sheet.setActiveRange(range);
    cell.setValues([[networkList.jsonResponse[i].name, networkList.jsonResponse[i].id]]);
  }
}

function customAPICall() {
 
  var ui = SpreadsheetApp.getUi();
  var userData = getUserInfo();
  
  var apikey = userData.apikey;
  if (apikey.length <= 20) {ui.alert('Your API key is missing or too short.'); return;}
  
  var response = ui.prompt('What is the URL you want to fetch?', 'Enter the entire URL, including https:// and the domain.', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) {
   ui.alert('The user chose to close the dialog.'); 
   return;
  }
   var url = response.getResponseText();

  var sheet = switchSheets('Advanced output');
  var apiResult = apiCall(url, apikey);
  
  sheet.clear();
  range = sheet.getRange("A1");
  cell = sheet.setActiveRange(range);
  cell.setValue(['Printing response from ' + url])
  range = sheet.getRange("A2")
  cell = sheet.setActiveRange(range)
  cell.setValue([apiResult.stringResponse]);
}