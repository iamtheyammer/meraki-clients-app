//11:22PM, 4/5/18
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
      .addSeparator()
      .addItem('Make a wish', 'makeAWish')
      .addToUi(); //Completes the add call.
}

function connectToMeraki() {
  try {
    Logger.log('beginning main');
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  sheet.clear();
  logAndUpdateCell('Getting data from User data sheet and checking license...', 'A1', 'Results');
  var userData = getUserInfo(); //grab the user's data: see discreetFunctions
  if (userData == 'OK' || userData == 'CLOSE') return;
  var apikey = userData.apikey; //set our api key from above data
  if (apikey.length <= 20) {ui.alert('Your API key is missing or too short.'); return;} //check the API key is longer than 20 characters
   
    if (userData.licenseMaxClients == -1) {
      ui.alert('Your license is expired.', 'Please get a new license at merakiblocki.com and use that.', ui.ButtonSet.OK);
      return;
    } else if (userData.licenseMaxClients == -2) {
      return; //they've already been told...
    }

  logAndUpdateCell('Reporting analytics...', 'A1', 'Results');
  apiCallPut('https://api.mismatch.io/analytics?id=vGWK3gnQozAAjuCkU9ni7jH93yCutPRfsnU6HtaAn66gq4ekRtwGk9zTTYXgbbAk&function=connectToMeraki', 'noApiKeyNeeded'); //analytics
  
  var merakiOrganizationId = userData.organizationId;
  var merakiClientsURL;
  logAndUpdateCell('Getting clients from Meraki...', 'A1', 'Results');
  var currentClients = apiCall('https://api.meraki.com/api/v0/devices/' + userData.securityApplianceSerial + '/clients?timespan=' + userData.clientTimespan, apikey); //grab all clients connected to security appliance
  if (currentClients == 'OK' || currentClients == 'CLOSE') return;
  var numberOfClients = currentClients.jsonResponse.length;

 /* for (var i = 0; i < numberOfClients; i++) {
    merakiClientsURL = userData.clientsURL + '#q=' + encodeURIComponent(currentClients.jsonResponse[i].mac);
    range = sheet.getRange("A" + (i+2) + ":E" + (i+2));
    cell = sheet.setActiveRange(range);
    cell.setValues([[currentClients.jsonResponse[i].description, currentClients.jsonResponse[i].mac, currentClients.jsonResponse[i].ip, currentClients.jsonResponse[i].usage.recv/1000000 + '/' + currentClients.jsonResponse[i].usage.sent/1000000, merakiClientsURL]]);
  }
  This for loop prints out all client data.*/

  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results");
  var currentClients = currentClients; //gets the clients that are currently connected.
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Approved clients");
  logAndUpdateCell('Getting approved clients...', 'A1', 'Results');
  var approvedClientsResponse = getApprovedClients();
  var approvedClients = JSON.stringify(approvedClientsResponse);
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results");
  var unknownClients = new Array(); //this is the array that will hold the MAC addresses for clients we haven't approved
  var unknownClientsPrint = []; //this is the array that will be printed to the Results sheet
  var unknownClientsLineNum = new Array(); //this is the array that will hold the line numbers of the clients that aren't approved so we can get more info about each unknown client without another API call

  var numberOfUnknownDevices = 0; //assume that there are no unknown devices
  if (userData.licenseMaxClients != 0) var remainingLicenseClients = userData.licenseMaxClients; //ignore if unlimited license
  
  logAndUpdateCell('Checking for unapproved clients...', 'A1', 'Results');
  for(i in currentClients.jsonResponse){
    var row = currentClients.jsonResponse[i].mac; //set the row to a mac address
    var duplicate = false; //assume every row is not a duplicate
    if (remainingLicenseClients < 1) {var notAllClientsPrinted = true; break;} else {remainingLicenseClients -= 1;}
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

  sheet.clear(); //reset the sheet and set the headings
  sheet.getRange('A1:E1').setValues([['Description', 'MAC address', 'LAN IP', 'Data down/up in MB', 'Meraki dashboard URL']]);
  /*sheet.getRange('A1').setValue('Description');
  sheet.getRange('B1').setValue('MAC address');
  sheet.getRange('C1').setValue('LAN IP');
  sheet.getRange('D1').setValue('Data down/up in MB');
  sheet.getRange('E1').setValue('Meraki dashboard URL');*/
  //sheet.getRange(2, 1, unknownClients.length, 2).setValues(newData); //gets a selection. starts on row 2, column 1, with a length of the number of unknown clients, 2 wide

  /*for (var i = 0; i < unknownClients.length; i++) {
    merakiClientsURL = userData.clientsURL + '#q=' + encodeURIComponent(unknownClients[i]);
    unknownClientsPrint.push([unknownClients[i],merakiClientsURL]);

  }*/

      /*if (userData.licenseValidity) {
      switch (userData.licenseMaxClients) {
        case 0:
          var numberOfUnknownClients = unknownClientsLineNum.length;
          break;
        case 1500:
          if (unknownClientsLineNum.length >= userData.licenseMaxClients) {
            var numberOfUnknownClients = unknownClientsLineNum.length;
          } else {
            var numberOfUnknownClients = userData.licenseMaxClients;
          }
          break;
        case 100000:
          if (unknownClientsLineNum.length >= userData.licenseMaxClients) {
            var numberOfUnknownClients = unknownClientsLineNum.length;
          } else {
            var numberOfUnknownClients = userData.licenseMaxClients;
          }
          break;
        case -1:
          ui.alert('Your license is expired.', 'Please get a new license at merakiblocki.com and use that.', ui.ButtonSet.OK);
          break;
       default:
          ui.alert('Internal error', 'Since this is my fault and not yours, here\'s a free pass.', ui.ButtonSet.OK);
          var numberOfUnknownClients = unknownClientsLineNum.length;
          break;
      }
    //} else {
      //ui.alert('Your license is invalid.', 'Your email most likely does not match your key. Your email is the email used to purchase the key originally. If you know everything is right, go to support at merakiblocki.com.', ui.ButtonSet.OK);
      //return;
    //}
       */

  logAndUpdateCell('Getting ready to display data...', 'A1', 'Results');
  for (var i = 0; i < unknownClientsLineNum.length; i++) {
    merakiClientsURL = userData.clientsURL + '#q=' + encodeURIComponent(unknownClients[i]); //set up the URLs: encode the mac address so it's readable by meraki
    unknownClientsPrint.push([currentClients.jsonResponse[unknownClientsLineNum[i]].description, currentClients.jsonResponse[unknownClientsLineNum[i]].mac, currentClients.jsonResponse[unknownClientsLineNum[i]].ip, currentClients.jsonResponse[unknownClientsLineNum[i]].usage.recv/1000 + '/' + currentClients.jsonResponse[unknownClientsLineNum[i]].usage.sent/1000, merakiClientsURL]);
  }
    if (unknownClientsPrint.length >= 1) { //if there are  unknown clients
      if (notAllClientsPrinted == true) {
      unknownClientsPrint.push(['Not all clients were scanned due to an insufficient license.', '', '', '', '']);
      unknownClientsPrint.push(['Upgrade your license at', '', '=HYPERLINK(\"https://merakiblocki.com/#pricing\",\"merakiblocki.com\")', '', '']);
      sheet.getRange(2, 1, (unknownClients.length+=2), 5).setValues(unknownClientsPrint); //get a range large enough for our data and paste the data in
      } else {
        unknownClientsPrint.push(['All clients scanned. Thank you for playing fair.', '', '', '', '']);
        sheet.getRange(2, 1, (unknownClients.length+=1), 5).setValues(unknownClientsPrint); //get a range large enough for our data and paste the data in
      }
      sheet.activate()
	  //print out the unknown clients
    } else { //otherwise,
      sheet.activate();
      ui.alert('Congratulations!', 'You don\'t have any un-approved devices! (if you think you do, you might want to check the \'Client timespan\' setting in the User data sheet. you can also check all of the sheets that make up your approved devices list and check those as well.)', ui.ButtonSet.OK); //congratulate the user that their network is in pristine perfectness
    }
    Logger.log('main function done');
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
  if (userData == 'OK' || userData == 'CLOSE') return;

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
  if (response == 'OK' || response == 'CLOSE') return;
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
  if (userData == 'OK' || userData == 'CLOSE') return;
  try {
  apiCallPut('https://api.mismatch.io/analytics?id=vGWK3gnQozAAjuCkU9ni7jH93yCutPRfsnU6HtaAn66gq4ekRtwGk9zTTYXgbbAk&function=approveUnknownClients', 'noApiKeyNeeded'); //analytics

  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results").activate(); //visually switch sheets to Results
  var unknownClients = sheet.getRange('B2:B' + sheet.getLastRow()).getValues(); //grab the mac addresses

  var response = ui.alert('Are you sure you want to approve all clients listed on this sheet?', 'This will add all MAC addresses on this sheet to your approved devices list.' , ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) {
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

function makeAWish() {
  SpreadsheetApp.getUi().alert('Make a wish', 'Help us help you! Click below to make your wish and we\'ll get back to you within 72 hours:\n http://jira.mismatch.io/servicedesk/customer/portal/1 \n (you may have to copy and paste the URL into your address bar)', SpreadsheetApp.getUi().ButtonSet.OK)
}
