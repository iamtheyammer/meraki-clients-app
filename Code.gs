//9:43AM, 2/17/18
function onInstall(e) {
 onOpen(e); 
 initializeSpreadsheet();
}

function onOpen(e) { //The 'e' there tells the system that this doesn't work in certain authentication modes. Something to look into, but not a priority.
  var ui = SpreadsheetApp.getUi();
  SpreadsheetApp.getUi().createAddonMenu() //Tells the UI to add a space to put items under the add-ons menu in docs
      .addItem('Start', 'connectToMeraki') 
      .addItem('Block remaining clients', 'blockUnknownClients')
      .addItem('Approve remaining clients', 'approveUnknownClients')
      .addSeparator()
      .addSubMenu(ui.createMenu('Advanced')
          .addItem('Completely clear sheet', 'completelyClearSheet')
          .addItem('Get Spreadsheet ID', 'getSpreadsheetId')
          .addItem('Print organizations', 'printOrganizations')
          .addItem('Print networks', 'printNetworks')
          .addItem('Unblock clients on Results sheet', 'unblockClients')
          .addItem('Initialize spreadsheet', 'initializeSpreadsheet')
                  .addSubMenu(ui.createMenu('Custom API call')
                              .addItem('Custom GET request', 'customAPICall')
                              .addItem('Custom PUT request', 'customAPICallPut')))
      .addToUi(); //Completes the add call.
}

function connectToMeraki() {
  try {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results'); //get sheet
  sheet.clear(); //clear it to draw focus to the status indicator in A1
  sheet.activate();
  var ui = SpreadsheetApp.getUi();
  logAndUpdateCell('Getting user data and checking license...', 'A1');
  var userData = getUserInfo(); //grab the user's data: see discreetFunctions
  var apikey = userData.apikey; //set our api key from above data
  if (apikey.length <= 20) {ui.alert('Your API key is missing or too short.'); return;} //check the API key is longer than 20 characters
  
    if (userData.licenseMaxClients == -1) { //make sure the license is valid
      ui.alert('Your license is expired.', 'Please get a new license at merakiblocki.com and use that.', ui.ButtonSet.OK);
      return;
    } else if (userData.licenseMaxClients == -2) {
      return; //they've already been told...
    }
    
  logAndUpdateCell('Reporting analytics...', 'A1');
  apiCallPut('https://api.mismatch.io/analytics?id=vGWK3gnQozAAjuCkU9ni7jH93yCutPRfsnU6HtaAn66gq4ekRtwGk9zTTYXgbbAk&function=connectToMeraki', 'noApiKeyNeeded'); //analytics
    
  var merakiOrganizationId = userData.organizationId;
  var merakiClientsURL;
  logAndUpdateCell('Getting clients from Meraki...', 'A1');
  var currentClients = apiCall('https://api.meraki.com/api/v0/devices/' + userData.securityApplianceSerial + '/clients?timespan=' + userData.clientTimespan, apikey); //grab all clients connected to security appliance
  var numberOfClients = currentClients.jsonResponse.length;

  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results");
  //var currentClients = currentClients; //gets the clients that are currently connected.
  logAndUpdateCell('Getting approved clients... (this can take a while)', 'A1');
  var approvedClientsResponse = getApprovedClients();
  var approvedClients = JSON.stringify(approvedClientsResponse);
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results");
  var unknownClients = new Array(); //this is the array that will hold the MAC addresses for clients we haven't approved
  var unknownClientsPrint = []; //this is the array that will be printed to the Results sheet
  var unknownClientsLineNum = new Array(); //this is the array that will hold the line numbers of the clients that aren't approved so we can get more info about each unknown client without another API call
  var numberOfUnknownDevices = 0; //assume that there are no unknown devices
  if (userData.licenseMaxClients != 0) var remainingLicenseClients = userData.licenseMaxClients; //ignore if unlimited license
  logAndUpdateCell('Checking for unapproved devices... (this can also take a while)', 'A1');
    
  for(i in currentClients.jsonResponse){
    var row = currentClients.jsonResponse[i].mac; //set the row to a mac address
    var duplicate = false; //assume every row is not a duplicate
    if (remainingLicenseClients < 1) {var notAllClientsPrinted = true; break;}
    if (userData.licenseMaxClients != 0) remainingLicenseClients -= 1; //for licensing: if it's unlimited, skip. 
    for(j in approvedClientsResponse){
      if(row == approvedClientsResponse[j]){ //if the row matches an entry on the approved clients list.
        duplicate = true; //mark it as a duplicate
      }
    }
    if(!duplicate){ //if it's not a duplicate,
      unknownClients.push(row); //add it to unknownClients, and
      unknownClientsLineNum.push(i); //add the line number to unknownClientsLineNum
    }
  }
  
  logAndUpdateCell('Getting ready to print result...', 'A1');
  for (var i = 0; i < unknownClientsLineNum.length; i++) {
    merakiClientsURL = userData.clientsURL + '#q=' + encodeURIComponent(unknownClients[i]); //set up the URLs: encode the mac address so it's readable by meraki
    unknownClientsPrint.push([currentClients.jsonResponse[unknownClientsLineNum[i]].description, currentClients.jsonResponse[unknownClientsLineNum[i]].mac, currentClients.jsonResponse[unknownClientsLineNum[i]].ip, currentClients.jsonResponse[unknownClientsLineNum[i]].usage.recv/1000 + '/' + currentClients.jsonResponse[unknownClientsLineNum[i]].usage.sent/1000, merakiClientsURL]); 
  }
    
  logAndUpdateCell('Printing result.', 'A1');
  sheet.clear(); //reset the sheet and set the headings
    if (unknownClientsPrint.length >= 1) { //if there are  unknown clients
      sheet.getRange('A1:E1').setValues([['Description', 'MAC address', 'LAN IP', 'Data up/down in MB', 'Meraki dashboard URL']]);
      if (notAllClientsPrinted == true) {
        unknownClientsPrint.push(['Not all clients were scanned due to an insufficient license.', '', '', '', '']);
        unknownClientsPrint.push(['Upgrade your license at', '', '=HYPERLINK(\"https://merakiblocki.com/#pricing\",\"merakiblocki.com\")', '', '']);
        sheet.getRange(2, 1, (unknownClients.length+=2), 5).setValues(unknownClientsPrint); //get a range large enough for our data and paste the data in
      } else {
        unknownClientsPrint.push(['All clients scanned. Thank you for playing fair.', '', '', '', '']);
        sheet.getRange(2, 1, (unknownClients.length+=1), 5).setValues(unknownClientsPrint); //get a range large enough for our data and paste the data in
      }
      sheet.activate();
	  //print out the unknown clients
      
    } else { //otherwise,
      if (userData.licenseType == 'trial') {
        sheet.getRange('A1:A4').setValues([['It\'s very possible that since you\'re using a trial license, no unauthorized clients were found because you can only scan ' + userData.licenseMaxClients + ' clients.'],
                                           ['Right now, there are ' + approvedClientsResponse.length + ' clients on all of your approved clients sheets. Maybe you want to pair that down?'],
                                           ['Or, it\'s very possible that you have no unapproved clients.'],
                                           ['If so, congratulations!']]);
      } else {
      sheet.activate();
      sheet.getRange('A1:A4').setValues([['Congratulations!'],
                                         ['You have no unapproved devices!'],
                                         [''],
                                         ['(You might want to check your client timespan, which is currently ' + userData.clientTimespan + ' seconds.)']]);
      }
    }
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

function blockUnknownClients() {
  try {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var cell;
  var range;
  var userData = getUserInfo();
  
  apiCallPut('https://api.mismatch.io/analytics?id=vGWK3gnQozAAjuCkU9ni7jH93yCutPRfsnU6HtaAn66gq4ekRtwGk9zTTYXgbbAk&function=blockUnknownClients', 'noApiKeyNeeded'); //analytics
  
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
    
  } catch(e) {
    var payload = {
       "id":"vGWK3gnQozAAjuCkU9ni7jH93yCutPRfsnU6HtaAn66gq4ekRtwGk9zTTYXgbbAk",
       "function":"blockUnknownClients",
       "fileName":e.fileName,
       "lineNumber":e.lineNumber,
       "message":e.message,
    };
    apiCallPost('https://api.mismatch.io/analytics/error', payload);
    SpreadsheetApp.getUi().alert('I\'m sorry, something didn\'t work right. ' + 'I\'ve reported this to the developers. Here\'s the full error: ' + e.message); 
  }
}

function approveUnknownClients() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  var userData = getUserInfo();
  try {
  apiCallPut('https://api.mismatch.io/analytics?id=vGWK3gnQozAAjuCkU9ni7jH93yCutPRfsnU6HtaAn66gq4ekRtwGk9zTTYXgbbAk&function=approveUnknownClients', 'noApiKeyNeeded'); //analytics
  
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results").activate(); //visually switch sheets to Results
  var unknownClients = sheet.getRange('B2:B' + sheet.getLastRow()).getValues(); //grab the mac addresses
  
  var response = ui.alert('Are you sure you want to approve all clients listed on this sheet?', 'This will add all MAC addresses on this sheet to your approved devices list.' , ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) {
    ui.alert('Cancelling.');
    return;
    
  }
   
  var sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Approved clients').getRange('A2').getValues(); //grab sheet to write to
  var sheetName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Approved clients').getRange('B2').getValues();
  //var sheetCell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Approved clients').getRange('C2').getValues();
  var writingSheet = SpreadsheetApp.openByUrl(sheetUrl).getSheetByName(sheetName);
  for (i = 0; i < unknownClients.length; i++) {
  sheet = writingSheet;
  sheet.appendRow([unknownClients[i].join()]); //turn the mac addresses into strings
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results").getRange("F" + (i+2) + ":F" + (i+2)); //select cell to write to
  var cell = sheet.activate(); //activate it
  cell.setValue([['Added to allowed list.']]); //write to it
  }
  
  sheet = writingSheet;
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
  
  } catch(e) {
    var payload = {
       "id":"vGWK3gnQozAAjuCkU9ni7jH93yCutPRfsnU6HtaAn66gq4ekRtwGk9zTTYXgbbAk",
       "function":"approveUnknownClients",
       "fileName":e.fileName,
       "lineNumber":e.lineNumber,
       "message":e.message,
    };
    apiCallPost('https://api.mismatch.io/analytics/error', payload);
    SpreadsheetApp.getUi().alert('I\'m sorry, something didn\'t work right. ' + 'I\'ve reported this to the developers. Here\'s the full error: ' + e.message); 
  }
}