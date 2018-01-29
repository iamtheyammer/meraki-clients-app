//10:35AM, 1/29/18
function onOpen(e) { //The 'e' there tells the system that this doesn't work in certain authentication modes. Something to look into, but not a priority.
  var ui = SpreadsheetApp.getUi();
  SpreadsheetApp.getUi().createAddonMenu() //Tells the UI to add a space to put items under the mTools add-ons menu in docs
      .addItem('Start', 'connectToMeraki') 
      .addItem('Block remaining clients', 'blockUnknownClients')
      .addItem('Approve remaining clients', 'approveUnknownClients')
      .addSeparator()
      .addSubMenu(ui.createMenu('Advanced')
          .addItem('Completely clear sheet', 'completelyClearSheet')
          .addItem('Print organizations', 'printOrganizations')
          .addItem('Print networks', 'printNetworks')
          .addItem('Unblock clients on Results sheet', 'unblockClients')
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
  var unknownClients = new Array(); //this is the array that will hold the MAC addresses for clients we haven't approved
  var unknownClientsPrint = []; //this is the array that will be printed to the Results sheet
  var unknownClientsLineNum = new Array(); //this is the array that will hold the line numbers of the clients that aren't approved so we can get more info about each unknown client without another API call
  
  var numberOfUnknownDevices = 0;
  for(i in currentClients.jsonResponse){
    var row = currentClients.jsonResponse[i].mac;
    var duplicate = false;
    for(j in approvedClients){
      if(row == approvedClients[j]){
        duplicate = true;
      }
    }
    if(!duplicate){
      unknownClients.push(row);
      unknownClientsLineNum.push(i);
    }
  }
  Logger.log("UNKNOWN CLIENTS:");
  Logger.log(unknownClients);
  
  sheet.clear();
  sheet.getRange('A1').setValue('Description');
  sheet.getRange('B1').setValue('MAC address');
  sheet.getRange('C1').setValue('LAN IP');
  sheet.getRange('C1').setValue('Data down/up in MB');
  sheet.getRange('C1').setValue('Meraki dashboard URL');
  //sheet.getRange(2, 1, unknownClients.length, 2).setValues(newData); //gets a selection. starts on row 2, column 1, with a length of the number of unknown clients, 2 wide
   
  /*for (var i = 0; i < unknownClients.length; i++) {
    merakiClientsURL = userData.clientsURL + '#q=' + encodeURIComponent(unknownClients[i]);
    unknownClientsPrint.push([unknownClients[i],merakiClientsURL]);

  }*/
  
  for (var i = 0; i < unknownClientsLineNum.length; i++) {
    merakiClientsURL = userData.clientsURL + '#q=' + encodeURIComponent(unknownClients[i]);
    unknownClientsPrint.push([clientList.jsonResponse[unknownClientsLineNum[i]].description, clientList.jsonResponse[unknownClientsLineNum[i]].mac, clientList.jsonResponse[unknownClientsLineNum[i]].ip, clientList.jsonResponse[unknownClientsLineNum[i]].usage.recv/1000 + '/' + clientList.jsonResponse[unknownClientsLineNum[i]].usage.sent/1000, merakiClientsURL]); 
  }
  
  
  Logger.log('UNKNOWN CLIENTS LENGTH:');
  Logger.log(unknownClients.length);
  Logger.log('UNKNOWN CLIENTS PRINT:');
  Logger.log(unknownClientsPrint);
  Logger.log('UNKNOWN CLIENTS LINENUM:');
  Logger.log(unknownClientsLineNum);
  sheet.getRange(2, 1, unknownClients.length, 5).setValues(unknownClientsPrint);
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
  var response = apiCallPut('https://n126.meraki.com/api/v0/networks/' + userData.networkId + '/clients/' + unknownClients[i] + '/policy?timespan=2592000&devicePolicy=blocked', apikey);
  range = sheet.getRange("C" + (i+2) + ":C" + (i+2));
  cell = sheet.setActiveRange(range);
  cell.setValue([['Blocked']]);
  Logger.log('Successfully blocked ' + unknownClients[i] + 'from the network.');
  Logger.log(response);
  Utilities.sleep(400);
  }
}

function approveUnknownClients() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var cell;
  var range;
  var userData = getUserInfo();
  
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results");
  var unknownClients = sheet.getRange('A2:A' + sheet.getLastRow()).getValues();
  
  var response = ui.alert('Are you sure you want to approve all clients listed on this sheet?', 'This will add all MAC addresses on this sheet to your approved devices list.' , ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) {
    ui.alert('Cancelling.');
    return;
  }
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Approved clients");
  
  for (i = 0; i < unknownClients.length; i++) {
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Approved clients");
  sheet.appendRow([unknownClients[i].join()]);
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results");
  range = sheet.getRange("C" + (i+2) + ":C" + (i+2));
  cell = sheet.setActiveRange(range);
  cell.setValue([['Added to allowed list.']]);
  }
}