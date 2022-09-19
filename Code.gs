function myFunction() {
  var responseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");
  var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RawData");
  var calendar = CalendarApp.getCalendarById('REDACTED@group.calendar.google.com');

  var startRow = 2;  // First row of data to process - 2 exempts my header row
  var numRows = responseSheet.getLastRow();   // Number of rows to process
  var numColumns = responseSheet.getLastColumn();
 
  var responseDataRange = responseSheet.getRange(startRow, 1, numRows-1, numColumns);
  var dataDataRange = dataSheet.getRange(startRow, 1, numRows-1, numColumns);
  var responseData = responseDataRange.getValues();
  var dataData = dataDataRange.getValues();
  var complete = "Done";

// NOTE: Starts at 0 in Javascript :)
for (var i = 0; i < responseData.length; ++i) {
    var row = responseData[i];
    var timeDate = dataData[i];
    var name = row[2]; //Item Name
    var facilitators = row[6]; // List of RAs
    var programLocation = row[5];
    var shoppingList = row[7];
    var flyer = null;
    flyer = row[8];
    var rDate = new Date((timeDate[2] + " " + timeDate[3])); //remind date
    var eventID = row[9]; //event marked Done
    
    // Parse RA emails from facilitators string
    var raNames = [];
    var raEmails = [];
    
    var facilitatorsArray = facilitators.split(", ");
    
    //add default CCs here
    var toEmail = "";
    
    for (var k = 0; k < facilitatorsArray.length; k++) {
        for (var j = 0; j < raNames.length; j++) {
            if (facilitatorsArray[k].localeCompare(raNames[j]) == 0) {
            toEmail += raEmails[j] + ", "; 
            }
        }
    }
    //clips off final space and comma
    toEmail = toEmail.substring(0, toEmail.length-2);
      
   // Rest of code is for new programs only!!
    if (eventID != complete) {
      var currentCell = responseSheet.getRange(startRow + i, numColumns);
/*
      // Old updating event code, incomplete
      var ourEvent = {
      'description' : facilitators + '\r' + '\r' + flyer,
      'location' : programLocation,
      'guests' : toEmail,
      'sendUpdates' : true,
      //'id' : randId[x]
      };
      */

      var newEvent = calendar.createEvent(name, rDate, rDate, {
        description: facilitators + '\r' + '\r' + flyer,
        location: programLocation,
        guests: toEmail,
        sendInvites: true
      });

      // Record event ID, unused
      //responseSheet.getRange(startRow + i, 11).setValue(newEvent.getId.toString);
      
      var message = "This is an automated email.\n\nA new program has been submitted through the form. RA(s) " + facilitators.toString() + " plan to facilitate: " + name.toString();
      
      if (flyer != null) {
        message = message + "\n\nThe following flyer is also awaiting approval: " + flyer.toString();
      }
      if (shoppingList != null) {
        message = message + "\n\nThe following shopping list is included:\n" + shoppingList.toString();
      }
      
      message = message + "\n\nThank you!";
    
      // Send Mitch + RAs the email
        MailApp.sendEmail('REDACTED@uncc.edu', '[Holshouser Program Form] New Submission', message, {
          cc: toEmail
        });
      
      currentCell.setValue(complete);
    }
  }



}
