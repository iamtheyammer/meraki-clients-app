//9:30PM 1/12/18
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

function connectToMeraki() {

  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var cell;
  var range;
  
  var response = ui.prompt('What is your API key?', ui.ButtonSet.OK_CANCEL);
  var apikey = response.getResponseText();
  
  var userInfo = getUserInfo(apikey) //got the user's organization and name
  
  //below, we verify that the the user's organization is correct
  var merakiOrganizationId = userInfo.userOrganization;
  Logger.log(merakiOrganizationId);
  
  verifyInfoWithUser(userInfo.userName,'Oh well. Let us know that: \'the user told us that the organization name that we provided was not correct. this probably means that they have more than one organization and wanted to use that.\'') 
  Logger.log('organization verified.')
  
  var deviceList = apiCall('https://api.meraki.com/api/v0/organizations/' + merakiOrganizationId + '/networks/' + 'L_641762946900302975' +  '/devices', apikey);
  Logger.log('got the device infos:');
  var numberOfDevices = deviceList.jsonResponse.length;
  
  for (var i = 0; i < numberOfDevices; i++) {
    range = sheet.getRange("A" + (i+2) + ":D" + (i+2));
    cell = sheet.setActiveRange(range);
    cell.setValues([[deviceList.jsonResponse[i].serial, deviceList.jsonResponse[i].mac, deviceList.jsonResponse[i].model, deviceList.jsonResponse[i].lanIp]]);
    
  }
   

  
  
/*  var merakiOrganizationId = firstCall.slice(7, 25);
  var response = UrlFetchApp.fetch('https://api.meraki.com/api/v0/organizations/' + merakiOrganizationId + '/networks', options);

 
   var merakiNetworkId = "L_641762946900302972";
  var response = UrlFetchApp.fetch('https://api.meraki.com/api/v0/organizations/' + merakiOrganizationId + '/networks/' + merakiNetworkId + '/devices', options);

 
  var merakiDeviceSerial = "VRT-2207617868457";
  var response = UrlFetchApp.fetch('https://api.meraki.com/api/v0/organizations/' + merakiOrganizationId + '/networks/' + merakiNetworkId + '/devices/' + merakiDeviceSerial + '/clients', options);
 
  */
}
//curl -L -H 'X-Cisco-Meraki-API-Key: 38058ca4c95b21ae6b4c568e19d280bc9bc5495d' -X GET -H 'Content-Type: application/json' 'https://api.meraki.com/api/v0/devices/02:02:00:47:62:a9/clients?timespan=86400'