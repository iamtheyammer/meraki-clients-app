//12:42AM, 7/4/19
function printOrganizations() {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var ui = SpreadsheetApp.getUi();
    var cell;
    var range;
    
    var userData = getUserInfo(true); //true to mute warnings
    Logger.log(userData);
    if (userData == 'OK' || userData == 'CLOSE') return;

    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Advanced output").activate(); //visually switch to advanced output
    sheet.clear();
    
    var apikey = userData.apikey;
    if (apikey.length <= 20) {ui.alert('Your API key is missing or too short.'); return;} //get (^) and verify api key

    var userOrganizations = apiCall('https://api.meraki.com/api/v0/organizations/', apikey); //make api call
    if (userOrganizations == 'OK' || userOrganizations == 'CLOSE') return;
   
    if (userOrganizations.stringResponse.indexOf('Live Demo')) {
      range = sheet.getRange("A2:B2"); //print heading
      cell = sheet.setActiveRange(range);
      cell.setValues([['Name', 'Organization ID']]);
      sheet.getRange('A1').setValue('Using a live demo organization? Some features may not be available. Also, the live demo org IDs don\'t tend to format well, so here\'s the JSON right from Meraki:');
      sheet.getRange('F2').setValue(userOrganizations.stringResponse);
      var numberOfOrganizations = userOrganizations.jsonResponse.length;

      for (var i = 0; i < numberOfOrganizations; i++) { //print result
        range = sheet.getRange("A" + (i+3) + ":B" + (i+3));
        cell = sheet.setActiveRange(range);
        cell.setValues([[userOrganizations.jsonResponse[i].name, userOrganizations.jsonResponse[i].id]]);
      }
    } else {
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
  } catch(e) {
    SpreadsheetApp.getUi().alert('I\'m sorry, something didn\'t work right. Here\'s the full error: ' + e.message);

  }
}

function printNetworks() {
  try {


  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var cell;
  var range;

  var userData = getUserInfo(true); //true to mute warnings
  if (userData == 'OK' || userData == 'CLOSE') return;

  var apikey = userData.apikey;
  if (apikey.length <= 20) {ui.alert('Your API key is missing or too short.'); return;} //get (^) and verify api key

  var organizationId = userData.organizationId;
  if (organizationId.length <= 1) {ui.alert('Your Organization ID is missing or too short.'); return;} //get (^) and verify organization id

  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Advanced output").activate(); //visually switch to advanced output
  sheet.clear();

  var networkList = apiCall('https://api.meraki.com/api/v0/organizations/' + organizationId + '/networks', apikey); //make api call
  if (networkList == 'OK' || networkList == 'CLOSE') return;
  var numberOfNetworks = networkList.jsonResponse.length;

  range = sheet.getRange("A1:B1"); //print heading
  cell = sheet.setActiveRange(range);
  cell.setValues([['Name', 'Network ID']]);

  for (var i = 0; i < numberOfNetworks; i++) { //print result
    range = sheet.getRange("A" + (i+2) + ":B" + (i+2));
    cell = sheet.setActiveRange(range);
    cell.setValues([[networkList.jsonResponse[i].name, networkList.jsonResponse[i].id]]);
  }
  } catch(e) {
    SpreadsheetApp.getUi().alert('I\'m sorry, something didn\'t work right. Here\'s the full error: ' + e.message);
  }
}

function listDevices() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var userData;
  var spreadsheetUrl;
  var response = ui.alert('This will erase the current sheet.', 'I\'m going to erase the current sheet and list all of your devices. I recommend doing this on a seperate spreadsheet and adding that spreadsheet to your Approved clients sheet, but that\'s your choice. Press cancel or the X in the top right to cancel.', ui.ButtonSet.OK_CANCEL);
  //if (response != ui.Button.OK) return;
  Logger.log(sheet.getSheetId());
  if (!sheet.getRange('H1').getDisplayValue()) {
    //logAndUpdateCell('First run on this sheet. Asking for URL...', 'A1', sheet.getSheetName());
    var response = ui.prompt('What is your main spreadsheet\'s URL?', 'Please paste in the URL of the spreadsheet with your API key and other data. If it\'s this spreadsheet, leave the box blank.', ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() != ui.Button.OK) return;
    if (response.getResponseText().length > 1) {
      spreadsheetUrl = response.getResponseText();
      try {
        //logAndUpdateCell('Getting data from User data sheet (different spreadsheet) and checking license...', 'A1', sheet.getSheetName());
        userData = getUserInfo(true, SpreadsheetApp.openByUrl(spreadsheetUrl).getSheetByName('User data'));
      } catch(e) {
        return ui.alert('Your main sheet URL is invalid.', 'Try opening it in your browser or re-copying it. It\'s also possible you don\'t have access to it.', ui.ButtonSet.OK);
      }
      Logger.log(response.getResponseText());
    } else {
      //logAndUpdateCell('Getting data from User data sheet (this spreadsheet) and checking license...', 'A1', sheet.getSheetName());
      var userData = getUserInfo();
    }
  } else {
    spreadsheetUrl = sheet.getRange('H1').getDisplayValue();
    try {
        userData = getUserInfo(true, SpreadsheetApp.openByUrl(spreadsheetUrl).getSheetByName('User data'));
      } catch(e) {
        return ui.alert('Your main sheet URL is invalid.', 'Try removing the data in H6.', ui.ButtonSet.OK);
    }
  }
  
  sheet.clear();
  if (userData == 'OK' || userData == 'CLOSE' || userData == ui.Button.CLOSE || userData == ui.Button.CANCEL) return;
  try {
    var apikey = userData.apikey; //set our api key from above data
  } catch(e) {
    return ui.alert('Something isn\'t right.', 'I failed to find data that I need from your User data sheet. Reach out to support for help with this issue.', ui.ButtonSet.OK);
  }
  if (apikey.length <= 20) {ui.alert('Your API key is missing or too short.'); return;} //check the API key is longer than 20 characters
  //logAndUpdateCell('Getting devices list and finishing up...', 'A1', sheet.getSheetName());
  var merakiDevices = apiCall('https://api.meraki.com/api/v0/networks/' + userData.networkId + '/devices', apikey).jsonResponse;
  
  sheet.getRange('A1:H1').setValues([['If you\'d like to add this spreadsheet to your Approved clients sheet, copy and paste in the row below.', '', '', '', '', 'Main spreadsheet URL:', '', spreadsheetUrl]]);
  sheet.getRange('A2:D2').setBackground('yellow').setValues([[SpreadsheetApp.getActiveSpreadsheet().getUrl(), SpreadsheetApp.getActiveSheet().getName(), 'C4', 'Meraki devices']]);
  sheet.getRange('A3:H3').setValues([['Name', 'Model', 'MAC Address', 'Serial Number', 'Address', 'Latitude', 'Longitude', 'Tags']]);
  var merakiDevicesPrint = [];
  for (var i = 0; i < merakiDevices.length; i++) {
    if (merakiDevices[i].tags) {
      merakiDevicesPrint.push([merakiDevices[i].name, merakiDevices[i].model, merakiDevices[i].mac, merakiDevices[i].serial, merakiDevices[i].address, merakiDevices[i].lat, merakiDevices[i].lng, merakiDevices[i].tags]);
    } else {
      merakiDevicesPrint.push([merakiDevices[i].name, merakiDevices[i].model, merakiDevices[i].mac, merakiDevices[i].serial, merakiDevices[i].address, merakiDevices[i].lat, merakiDevices[i].lng, '']); //empty space is because no tags
    }
  }
  sheet.getRange(4, 1, (merakiDevices.length), 8).setValues(merakiDevicesPrint); //get a range large enough for our data and paste the data in
}
