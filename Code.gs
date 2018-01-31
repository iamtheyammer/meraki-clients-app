//11:00PM, 1/30/18
function onOpen(e) { //The 'e' there tells the system that this doesn't work in certain authentication modes. Something to look into, but not a priority.
  var ui = SpreadsheetApp.getUi();
  SpreadsheetApp.getUi().createAddonMenu() //Tells the UI to add a space to put items under the add-ons menu in docs
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
  var userData = getUserInfo(); //grab the user's data: see discreetFunctions
  var apikey = userData.apikey; //set our api key from above data
  if (apikey.length <= 20) {ui.alert('Your API key is missing or too short.'); return;} //check the API key is longer than 20 characters
  
  var merakiOrganizationId = userData.organizationId;
  var merakiClientsURL;
  
  var clientList = apiCall('https://api.meraki.com/api/v0/devices/' + userData.securityApplianceSerial + '/clients?timespan=' + userData.clientTimespan, apikey); //grab all clients connected to security appliance
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
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Approved clients");
  var approvedClients = sheet.getRange('A2:A' + sheet.getLastRow()).getValues(); //gets the clients that are approved to connect.
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results");
  var unknownClients = new Array(); //this is the array that will hold the MAC addresses for clients we haven't approved
  var unknownClientsPrint = []; //this is the array that will be printed to the Results sheet
  var unknownClientsLineNum = new Array(); //this is the array that will hold the line numbers of the clients that aren't approved so we can get more info about each unknown client without another API call
  
  var numberOfUnknownDevices = 0; //assume that there are no unknown devices
  for(i in currentClients.jsonResponse){
    var row = currentClients.jsonResponse[i].mac; //set the row to a mac address
    var duplicate = false; //assume every row is not a duplicate
    for(j in approvedClients){
      if(row == approvedClients[j]){ //if the row matches an entry on the approved clients list.
        duplicate = true; //mark it as a duplicate
      }
    }
    if(!duplicate){ //if it's a duplicate,
      unknownClients.push(row); //add it to unknownClients, and
      unknownClientsLineNum.push(i); //add the line number to unknownClientsLineNum
    }
  }
  
  sheet.clear(); //reset the sheet and set the headings
  sheet.getRange('A1').setValue('Description');
  sheet.getRange('B1').setValue('MAC address');
  sheet.getRange('C1').setValue('LAN IP');
  sheet.getRange('D1').setValue('Data up/down in MB');
  sheet.getRange('E1').setValue('Meraki dashboard URL');
  //sheet.getRange(2, 1, unknownClients.length, 2).setValues(newData); //gets a selection. starts on row 2, column 1, with a length of the number of unknown clients, 2 wide
   
  /*for (var i = 0; i < unknownClients.length; i++) {
    merakiClientsURL = userData.clientsURL + '#q=' + encodeURIComponent(unknownClients[i]);
    unknownClientsPrint.push([unknownClients[i],merakiClientsURL]);

  }*/
  
  for (var i = 0; i < unknownClientsLineNum.length; i++) {
    merakiClientsURL = userData.clientsURL + '#q=' + encodeURIComponent(unknownClients[i]); //set up the URLs: encode the mac address so it's readable by meraki
    unknownClientsPrint.push([clientList.jsonResponse[unknownClientsLineNum[i]].description, clientList.jsonResponse[unknownClientsLineNum[i]].mac, clientList.jsonResponse[unknownClientsLineNum[i]].ip, clientList.jsonResponse[unknownClientsLineNum[i]].usage.recv/1000 + '/' + clientList.jsonResponse[unknownClientsLineNum[i]].usage.sent/1000, merakiClientsURL]); 
  }
  
  sheet.getRange(2, 1, unknownClients.length, 5).setValues(unknownClientsPrint); //get a range large enough for our data and paste the data in
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
  var unknownClients = sheet.getRange('B2:B' + sheet.getLastRow()).getValues(); //grab the mac addresses from the results sheet
  sheet.getRange('F1').setValue('Device policy');
  
  var response = ui.alert('Are you sure you want to block all clients listed on this sheet?', 'You can press no below to remove clients you don\'t want to block.' , ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) { 
    ui.alert('Cancelling.');
    return;
  }
  
  for (i = 0; i < unknownClients.length; i++) {
  //var unknownClientURI = encodeURIComponent(unknownClients[i]); //would encode the mac address but works ok as is
  var response = apiCallPut('https://n126.meraki.com/api/v0/networks/' + userData.networkId + '/clients/' + unknownClients[i] + '/policy?timespan=2592000&devicePolicy=blocked', apikey); //call the api to block the client
  range = sheet.getRange("F" + (i+2) + ":F" + (i+2)); //get the cell to print to
  cell = sheet.setActiveRange(range); //set the cell as active
  cell.setValue([['Blocked']]); //put data in the cell
  Utilities.sleep(400); //wait 400 milliseconds to comply with meraki's 5 calls/second limit
  }
}

function approveUnknownClients() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  var userData = getUserInfo();
  
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results").activate(); //visually switch sheets to Results
  var unknownClients = sheet.getRange('B2:B' + sheet.getLastRow()).getValues(); //grab the mac addresses
  
  var response = ui.alert('Are you sure you want to approve all clients listed on this sheet?', 'This will add all MAC addresses on this sheet to your approved devices list.' , ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) {
    ui.alert('Cancelling.');
    return;
  }
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Approved clients");
  
  for (i = 0; i < unknownClients.length; i++) {
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Approved clients"); //switch to approved clients
  sheet.appendRow([unknownClients[i].join()]); //turn the mac addresses into strings
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results").getRange("F" + (i+2) + ":F" + (i+2)); //select cell to write to
  var cell = sheet.activate(); //activate it
  cell.setValue([['Added to allowed list.']]); //write to it
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Approved clients');
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
  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData); //above code (whole paragraph) finds and removes duplicates in the approved clients list
}