function myFunction() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SecretSanta");
  var range = sheet.getRange(2,2,sheet.getLastRow()-1,sheet.getLastColumn()-1)
  var data = range.getValues();
  var leng = data.length;

  for (var i=0; i<leng;i++) {
    var row = data[i];
    var line = parseInt(i)+2;
    if(row[4] != "Sent") {

      if(i !=leng-1) {
        index = i+1
      }
      else {
          index = 0
      }
      var ui = HtmlService.createTemplateFromFile('ui');
      ui.row = row; 
      ui.data = data; 
      ui.index = index;
      var htmlOutput = ui.evaluate();
      var message = htmlOutput.getContent();
      MailApp.sendEmail( {to:row[0],subject:"Secret Santa!",htmlBody:message});
      sheet.getRange(line,6).setValue("Sent");
    }
  }
}
