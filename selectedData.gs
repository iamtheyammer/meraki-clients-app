//10:30AM, 6/2/18
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
   .setTitle('MerakiBlocki - Individual Actions');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function selectedMacPolicy(action) {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getSheetName() != 'Results') if (ui.alert('This isn\'t the results sheet.', 'Are you sure you want to proceed?', ui.ButtonSet.OK_CANCEL) != ui.Button.OK) return;
  var userData = getUserInfo();
  if (userData == 'OK' || userData == 'CLOSE') return;
  var apikey = userData.apikey;
  if (apikey.length <= 20) {ui.alert('Your API key is missing or too short.'); return;}
  /*if (userData.licenseType == 'basic') {
    Logger.log(userData.licenseType);
    return ui.alert('Insufficient license', 'Your current license does not support taking action on selections. Please upgrade at merakiblocki.com and try again.', ui.ButtonSet.OK);
  } //Currently commented out because taking action on selection is not currently a licenseType-specific action.
  */
  
  switch (action) {
    case 'view':
      var updateData = 'Checking policy...';
      break;
    case 'normalize':
      var intention = 'normal';
      var updateData = 'Normalizing...';
      break;
    case 'block': 
      var intention = 'blocked';
      var updateData = 'Blocking...';
      break;
    case 'whitelist':
      var intention = 'whitelisted';
      var updateData = 'Whitelisting...';
      break;
    default:
      return ui.alert('Unidentified intention', 'Intention must be passed into this variable: view,normalize,block,whitelist .', ui.ButtonSet.OK);
  }
  Logger.log('Intention: ' + intention);
  var ranges = SpreadsheetApp.getSelection().getActiveRangeList().getRanges();
  //from here
  for (var i = 0; i < ranges.length; i++) {
    //i is the number of selections
    for (var j = ranges[i].getRow(), k = 0; k < ranges[i].getValues().length; j++, k++) {
      //j is the row number; iterating
      //k number for the place of the value in the range; iterating (in an array [1,2,3,4], on the first iteration, k would be 0, then 1, then 2, then 3)
      sheet.getRange('F' + j).setValue(updateData);
      if (action == 'view') {
        var response = apiCall('https://' + userData.shard + '.meraki.com/api/v0/networks/' + userData.networkId + '/clients/' + ranges[i].getValues()[k] + '/policy?timespan=2592000', apikey);
      } else {
        Logger.log('apiCallPut');
        var response = apiCallPut('https://' + userData.shard + '.meraki.com/api/v0/networks/' + userData.networkId + '/clients/' + ranges[i].getValues()[k] + '/policy?timespan=2592000&devicePolicy=' + intention, apikey);
      }
      if (response == 'OK' || response == 'CLOSE') {
        sheet.getRange('F' + j).setValue('Failed to' + updateData + '.');
        Utilities.sleep(220); //to comply with Meraki's 5 calls/sec limit.
      /*} else if (response.jsonResponse.type == 'Group policy') {
        var groupPolicies = apiCall('https://' + userData.shard + '.meraki.com/api/v0/networks/' + userData.networkId + '/groupPolicies', apikey).jsonResponse;
        Logger.log(groupPolicies.filter(groupPolicies.id => groupPolicies.id === response.jsonResponse.id));
        var value = 'Group policy ID ' + response.jsonResponse.groupPolicyId + ' - ' + apiCall('https://' + userData.shard + '.meraki.com/api/v0/networks/' + userData.networkId + '/groupPolicies', apikey).jsonResponse.filter(response.jsonResponse.groupPolicyId).name;
        sheet.getRange('F' + j).setValue('OK');
    */} else {
        sheet.getRange('F' + j).setValue(response.jsonResponse.type);
        Utilities.sleep(220); //to comply with Meraki's 5 calls/sec limit.
      }
    }
  }
  //to here is code with an interesting story. read it here: https://gist.github.com/iamtheyammer/38cf6ffb1a059ce8718269a283b47f9a
}
