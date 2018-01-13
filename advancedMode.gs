//9:20PM 1/12/18
function printOrganizations() {

  switchSheets("Advanced output");
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var cell;
  var range;
  
  
  var response = ui.prompt('What is your API key?', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) {
   ui.alert('The user chose to close the dialog.'); 
   return;
  }
  var apikey = response.getResponseText();
  

  var userOrganizations = apiCall('https://api.meraki.com/api/v0/organizations/', apikey);
  
  range = sheet.getRange("A1");
  cell = sheet.setActiveRange(range);
  cell.setValue(['Printing User Organizations:'])
  range = sheet.getRange("A2")
  cell = sheet.setActiveRange(range)
  cell.setValue([userOrganizations.stringResponse])
}
  
function printNetworks() {
  
  switchSheets('Advanced output');
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var cell;
  
  verifyInfoWithUser('You\'ll need your API key and your organization ID. Make sure they\'re available.', 'Maybe they\'re not available. Try again when you\'re ready.');
  
  var response = ui.prompt('What is your API key?', ui.ButtonSet.OK_CANCEL);
  var apikey = response.getResponseText();
  
  var response = ui.prompt('What is your organization ID?', 'Don\'t know? Add-ons -> Advanced -> Print organizations', ui.ButtonSet.OK_CANCEL);
  var merakiOrganizationId = response.getResponseText();
  
  var networkList = apiCall('https://api.meraki.com/api/v0/organizations/' + merakiOrganizationId + '/networks' , apikey);
  var numberOfNetworks = networkList.jsonResponse.length;
  
  for (var i = 0; i < numberOfNetworks; i++) {
    range = sheet.getRange("D" + (i+2) + ":E" + (i+2));
    cell = sheet.setActiveRange(range);
    cell.setValues([[networkList.jsonResponse[i].name, networkList.jsonResponse[i].id]]);
  }
}

function customAPICall() {
 
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('What is your API key?', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) {
   ui.alert('The user chose to close the dialog.'); 
   return;
  }
  var apikey = response.getResponseText();
  
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