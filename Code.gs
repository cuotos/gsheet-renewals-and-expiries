var emailAddress = PropertiesService.getScriptProperties().getProperty("destEmail");
var sheetAddress = PropertiesService.getScriptProperties().getProperty("srcSheet");
var daysRemainingWarningThreshold = 21;
var daysRemainingCritialThreshold = 7;
var emailSubject = "Membership Expiry Warning";
var emailBody = "Memberships expiring in the next " + daysRemainingWarningThreshold + " days\n\n";
var sendToFile = false;

var monthNames = [ "January", "February", "March", "April", "May", "June",
                  "July", "August", "September", "October", "November", "December" ];

function main(){
  var birthdaysSoon = getExpiringEntries(0, daysRemainingWarningThreshold, true );
  var soonToExpireMemberships = getExpiringEntries(daysRemainingCritialThreshold, daysRemainingWarningThreshold );
  var criticalMemberships = getExpiringEntries(0, daysRemainingCritialThreshold);
  var expiredMemberships = getExpiringEntries(-1000, 0)
  
  // These currently mutate the global email body object, nasty.
  generateExpiryTypeEmailBlock(birthdaysSoon, 'Birthdays:');
  generateExpiryTypeEmailBlock(criticalMemberships, 'Critical:');
  generateExpiryTypeEmailBlock(soonToExpireMemberships, 'Expiring Soon:');
  generateExpiryTypeEmailBlock(expiredMemberships, 'Passed:');
    
  emailBody += "\n";
  emailBody += sheetAddress;
    
  if ( getExpiringEntries(daysRemainingCritialThreshold).length > 0 ){
    emailSubject = "CRITICAL: " + emailSubject;
  }
  
  if ( (soonToExpireMemberships.length + criticalMemberships.length + expiredMemberships.length) > 0 ){
    sendEmail(emailSubject, emailBody);
  }
}

function generateExpiryTypeEmailBlock(membershipList, heading){
    if ( membershipList.length > 0 ) {
    send = true;
    emailBody += heading + "\n";
    for(i = 0; i < membershipList.length; i++ ){
      var membership = membershipList[i];
      emailBody += generateEmailLine(membership);
    }
    emailBody += "\n";
  }
}

function generateEmailLine(membership){
  var title = membership[0];
  var expires = membership[3];
  var expiresDateString = expires.getDate() + " " + monthNames[expires.getMonth()] + " " + expires.getYear();
  var lastYearsCost = membership[4];
  var comment = membership[5]
    
  return title + " expires " + expiresDateString + " " + lastYearsCost + " " + comment + "\n";  
}

function getExpiringEntries(minIncludedDate, maxIncludedDate, birthdays){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
  var allData = SpreadsheetApp.getActiveSheet().getRange("A2:G").getValues();
  var expiring = ArrayLib.filterByDate(allData, 3, getOffsetDate(minIncludedDate), getOffsetDate(maxIncludedDate));
  expiring = ArrayLib.sort(expiring, 3, true);
  var filtered = [];
  
  for (i = 0; i < expiring.length; i++){
    var isABirthday = expiring[i][0].match(/birthday/i)
    
    if (expiring[i][6] == ''){
      if (birthdays) {
        if (isABirthday) {
          filtered.push(expiring[i]);
        }
      } else if (isABirthday == null) {
        filtered.push(expiring[i]);
      }
    }
  }
  return filtered;
}

function sendEmail(subject, message) {
  if (sendToFile) {
    Logger.log(subject + ": " + message);
  } else {
    MailApp.sendEmail(emailAddress, subject, message);
  }
}

function getOffsetDate(offset){
  var today = new Date().getTime();
  var offsetInMilliseconds = (1000 * 60 * 60 * 24) * offset;
  
  return new Date(today + offsetInMilliseconds);
} 

//Sorting Function
function onOpen(){
 var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sort')
    .addItem('Sort by Name', 'sortByName')
    .addItem('Sort by Date', 'sortByDate')
    .addToUi();
}

function sortByDate(){
  sortTable(4); 
}

function sortByName(){
  sortTable(1); 
}

function sortTable(column){
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange("A2:Z");
  range.sort([{column: column, ascending: true}, {column: 1, ascending: true}]);
}
