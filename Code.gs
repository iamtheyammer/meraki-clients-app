//8:45AM, 1/27/18
function onOpen(e) { //The 'e' there tells the system that this doesn't work in certain authentication modes. Something to look into, but not a priority.
  
  var ui = SpreadsheetApp.getUi();
  SpreadsheetApp.getUi().createMenu('MerakiApp') //Tells the UI to add a space to put items under the mTools add-ons menu in docs
      .addItem('Start', 'connectToMeraki') //Adds 'Start', the visible text. 'websiter' is the function we're calling.
      .addItem('Block remaining clients', 'blockUnknownClients')
      .addSeparator()
      .addSubMenu(ui.createMenu('Advanced')
          .addItem('Completely clear sheet', 'completelyClearSheet')
          .addItem('Print organizations', 'printOrganizations')
          .addItem('Print networks', 'printNetworks')
          .addSubMenu(ui.createMenu('Custom API call')
          .addItem('Custom GET request', 'customAPICall')
          .addItem('Custom PUT request', 'customAPICallPut')))
      .addToUi(); //Completes the add call.
}

function connectToMeraki() {

  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var cell;
  var range;
  var userData = getUserInfo();
  var apikey = userData.apikey;
  if (apikey.length <= 20) {ui.alert('Your API key is missing or too short.'); return;}
  
  var merakiOrganizationId = userData.organizationId;
  var merakiClientsURL;
  
  switchSheets('Results');
  
  var clientList = apiCall('https://api.meraki.com/api/v0/devices/' + userData.securityApplianceSerial + '/clients?timespan=' + userData.clientTimespan, apikey);
  Logger.log('got the device infos:');
  var numberOfClients = clientList.jsonResponse.length;
  
 /* for (var i = 0; i < numberOfClients; i++) {
    merakiClientsURL = userData.clientsURL + '#q=' + encodeURIComponent(clientList.jsonResponse[i].mac);
    range = sheet.getRange("A" + (i+2) + ":E" + (i+2));
    cell = sheet.setActiveRange(range);
    cell.setValues([[clientList.jsonResponse[i].description, clientList.jsonResponse[i].mac, clientList.jsonResponse[i].ip, clientList.jsonResponse[i].usage.recv/1000000 + '/' + clientList.jsonResponse[i].usage.sent/1000000, merakiClientsURL]]); 
  }
  This for loop prints out all client data.*/
  
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results");
  var currentClients = clientList; //gets the clients that are currently connected.
  Logger.log('CURRENT CLIENTS:');
  Logger.log(currentClients);
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Approved clients");
  var approvedClients = sheet.getRange('A2:A' + sheet.getLastRow()).getValues(); //gets the clients that are approved to connect.
  Logger.log('APPROVED CLIENTS:');
  Logger.log(approvedClients);
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results");
  var unknownClients = new Array();
  
  for(i in currentClients.jsonResponse){
    var row = currentClients.jsonResponse[i].mac;
    Logger.log(row);
    var duplicate = false;
    for(j in approvedClients){
      if(row == approvedClients[j]){
        duplicate = true;
      }
    }
    if(!duplicate){
      unknownClients.push(row);
    }
  }
  Logger.log("UNKNOWN CLIENTS:");
  Logger.log(unknownClients);
  
  sheet.clear();
  var cell = sheet.getRange('A1');
  cell.setValue('Mac address');
  var cell = sheet.getRange('B1');
  cell.setValue('Meraki dashboard URL');
  var cell = sheet.getRange('C1');
  cell.setValue('These are all of the clients that aren\'t on your approved clients list.');
  
  for (var i = 0; i < unknownClients.length; i++) {
    merakiClientsURL = userData.clientsURL + '#q=' + encodeURIComponent(unknownClients[i]);
    range = sheet.getRange("A" + (i+2) + ":B" + (i+2));
    cell = sheet.setActiveRange(range);
    cell.setValues([[unknownClients[i], merakiClientsURL]]);
  }

}

function blockUnknownClients() {
 
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var cell;
  var range;
  var userData = getUserInfo();
  var apikey = userData.apikey;
  if (apikey.length <= 20) {ui.alert('Your API key is missing or too short.'); return;}
  
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results");
  var unknownClients = sheet.getRange('A2:A' + sheet.getLastRow()).getValues();
  
  var response = ui.alert('Are you sure you want to block all clients listed on this sheet?', 'You can press no below to remove clients you don\'t want to block.' , ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) {
    ui.alert('Cancelling.');
    return;
  }
  
  for (i = 0; i < unknownClients.length; i++) {
  Logger.log('Attempting to block ' + unknownClients[i] + 'from the network...');
  var unknownClientURI = encodeURIComponent(unknownClients[i]);
  var response = apiCallPut('https://n126.meraki.com/api/v0/networks/' + userData.networkId + '/clients/' + unknownClients[i] + '/policy?timespan=2592000&devicePolicy=blocked', apikey)
  range = sheet.getRange("B" + (i+2) + ":B" + (i+2));
  cell = sheet.setActiveRange(range);
  cell.setValue([['Blocked']]);
  Logger.log('Successfully blocked ' + unknownClients[i] + 'from the network.');
  Logger.log(response);
  }
}