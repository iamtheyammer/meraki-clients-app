//9:33PM, 1/26/18
function onOpen(e) { //The 'e' there tells the system that this doesn't work in certain authentication modes. Something to look into, but not a priority.
  
  var ui = SpreadsheetApp.getUi();
  SpreadsheetApp.getUi().createMenu('MerakiApp') //Tells the UI to add a space to put items under the mTools add-ons menu in docs
      .addItem('Start', 'connectToMeraki') //Adds 'Start', the visible text. 'websiter' is the function we're calling.
      .addItem('Reset sheet', 'resetSheet')
      .addSeparator()
      .addSubMenu(ui.createMenu('Advanced')
          .addItem('Completely clear sheet', 'completelyClearSheet')
          .addItem('Print organizations', 'printOrganizations')
          .addItem('Print networks', 'printNetworks')
          .addItem('Custom API call', 'customAPICall'))
      .addToUi(); //Completes the add call.
}

function removeDuplicates() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var data = sheet.getDataRange().getValues();
  var newData = new Array();
  for(i in data){
    var row = data[i];
    var duplicate = false;
    for(j in newData){
      if(row.join() == newData[j].join()){
        duplicate = true;
      }
    }
    if(!duplicate){
      newData.push(row);
    }
  }
  sheet.clearContents();
  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
  Logger.log(newData);
}

function myFunction() {
 var sheet = SpreadsheetApp.getActiveSheet();
 var ui = SpreadsheetApp.getUi();
 var cell;
 var range;

  switchSheets('Results');
  var currentClients = sheet.getRange('B2:B' + sheet.getLastRow()).getValues(); //gets the clients that are currently connected.
  Logger.log(currentClients);
  switchSheets('Approved clients');
  var approvedClients = sheet.getRange('A2:A' + sheet.getLastRow()).getValues(); //gets the clients that are approved to connect.
  Logger.log(approvedClients);
  switchSheets('Results');
  var unknownClients = new Array();
  
  for(i in currentClients){
    var row = currentClients[i];
    var duplicate = false;
    for(j in unknownClients){
      if(row.join() == unknownClients[j].join()){
        duplicate = true;
      }
    }
    if(!duplicate){
      unknownClients.push(row);
    }
  }
  Logger.log(unknownClients);
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
  
  for (var i = 0; i < numberOfClients; i++) {
    merakiClientsURL = userData.clientsURL + '#q=' + encodeURIComponent(clientList.jsonResponse[i].mac);
    range = sheet.getRange("A" + (i+2) + ":E" + (i+2));
    cell = sheet.setActiveRange(range);
    cell.setValues([[clientList.jsonResponse[i].description, clientList.jsonResponse[i].mac, clientList.jsonResponse[i].ip, clientList.jsonResponse[i].usage.recv/1000000 + '/' + clientList.jsonResponse[i].usage.sent/1000000, merakiClientsURL]]); 
  }
  

}
  
/*  var merakiOrganizationId = firstCall.slice(7, 25);
  var response = UrlFetchApp.fetch('https://api.meraki.com/api/v0/organizations/' + merakiOrganizationId + '/networks', options);

 
   var merakiNetworkId = "L_641762946900302972";
  var response = UrlFetchApp.fetch('https://api.meraki.com/api/v0/organizations/' + merakiOrganizationId + '/networks/' + merakiNetworkId + '/devices', options);

 
  var merakiDeviceSerial = "VRT-2207617868457";
  var response = UrlFetchApp.fetch('https://api.meraki.com/api/v0/organizations/' + merakiOrganizationId + '/networks/' + merakiNetworkId + '/devices/' + merakiDeviceSerial + '/clients', options);
 
  */

//curl -L -H 'X-Cisco-Meraki-API-Key: 38058ca4c95b21ae6b4c568e19d280bc9bc5495d' -X GET -H 'Content-Type: application/json' 'https://api.meraki.com/api/v0/devices/02:02:00:47:62:a9/clients?timespan=86400'