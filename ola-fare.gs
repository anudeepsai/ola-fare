// DO WHAT THE FUCK YOU WANT TO PUBLIC LICENSE 
//                    Version 2, December 2004 

// Copyright (C) 2017 Anudeep Sai N <saianudeepsai at gmail.com> 

// Everyone is permitted to copy and distribute verbatim or modified 
// copies of this license document, and changing it is allowed as long 
// as the name is changed. 

//            DO WHAT THE FUCK YOU WANT TO PUBLIC LICENSE 
//   TERMS AND CONDITIONS FOR COPYING, DISTRIBUTION AND MODIFICATION 

//  0. You just DO WHAT THE FUCK YOU WANT TO.


/**
 * Default entry point for the web app using Google App script.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index');
}

function calculateOlaExpenditure() {
    var ssNew = SpreadsheetApp.create("Ola Fare Expenditure");
    var mySheet = ssNew.getSheets()[0];
    Logger.log(ssNew.getId());
    var range = mySheet.getRange("A1:B1");
    range.setBackground("black");
    range.setFontColor("white");


    mySheet.appendRow(['Date', 'Amount in INR'])
    var totalFare = 0;
    SpreadsheetApp.openById(ssNew.getId());
    var emails = GmailApp.search('from:noreply@olacabs.com  to:me ola invoice');
    for (var i = 0; i < emails.length; i++) {
        var messages = emails[i].getMessages();
        for (var j = 0; j < messages.length; j++) {
            var messageBody = messages[j].getPlainBody();
            var fareAmount = messageBody.match(/₹[0-9,.]*/)
            var dateOfBooking = messageBody.match(/(([0-9])|([0-2][0-9])|([3][0-1]))\s(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(,\s|\s)\d{4}/);
            if (dateOfBooking && fareAmount) {
                mySheet.appendRow([dateOfBooking[0], fareAmount[0].substring(1)]);
                totalFare += parseInt(fareAmount[0].substring(1));
            }
        }
    }
    mySheet.appendRow(["Total", totalFare]);
    return totalFare;
    
}
