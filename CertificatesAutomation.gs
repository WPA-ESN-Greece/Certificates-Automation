function generateCertificates() {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const DataSheet = ss.getSheetByName('Data')
  const settingsSheet = ss.getSheetByName('Settings')
  var lastRow = DataSheet.getLastRow()-1
  const rawData = DataSheet.getRange(2,1,lastRow,4).getValues()

  var googleDocTemplateID = settingsSheet.getRange('Settings!H6').getValue();
  var outputFolderID = settingsSheet.getRange('Settings!H3').getValue();
  var outputFolder = DriveApp.getFolderById(outputFolderID)

//Basic Information for Certificates from Settings sheet. 
  var eventName = settingsSheet.getRange('Settings!B3').getValue()
  var eventStart = settingsSheet.getRange('Settings!B6').getValue()
  var eventEnd = settingsSheet.getRange('Settings!B9').getValue()
  var eventMonth = settingsSheet.getRange('Settings!B12').getValue()
  var eventYear = settingsSheet.getRange('Settings!B15').getValue()
  var eventPlace = settingsSheet.getRange('Settings!B18').getValue()
  var relatedPosition = settingsSheet.getRange('Settings!D3').getValue()
  var fullNameRP = settingsSheet.getRange('Settings!D6').getValue()
  var fullNameHead = settingsSheet.getRange('Settings!D9').getValue()
  var extraTitle = settingsSheet.getRange('Settings!D8').getValue()


//Data Loop
  rawData.forEach(function (row, index) {
  
  var participant = row[3] //full name per row
  var act = row[2]        //act per row
  var tempDoc = DriveApp.getFileById(googleDocTemplateID).makeCopy('tempCertficate',DriveApp.getFolderById(outputFolderID))
  var tempDocID = DocumentApp.openById(tempDoc.getId())
  var body = DocumentApp.openById(tempDoc.getId()).getBody()
  
  //Text Replacing
    body.replaceText('{{Full Name}}', participant)
    body.replaceText('{{act}}', act)

    body.replaceText('{{event name full}}',eventName)
    body.replaceText('{{starting date}}',eventStart)
    body.replaceText('{{end date}}',eventEnd)
    body.replaceText('{{month}}',eventMonth)
    body.replaceText('{{year}}',eventYear)
    body.replaceText('{{place}}',eventPlace)
    body.replaceText('{{Full Name Related Position}}',fullNameRP)
    body.replaceText('{{Related Position}}',relatedPosition)
    body.replaceText('{{Head}}',extraTitle)
    body.replaceText('{{Full Name Head}}',fullNameHead)

tempDocID.saveAndClose()
 
  //PDF Creation
  var blobPDF = tempDoc.getAs(MimeType.PDF)
  var pdf = outputFolder.createFile(blobPDF).setName('Certficate: ' + participant + " - " + eventName)

  DriveApp.getFileById(tempDoc.getId()).setTrashed(true)

  var pdfUrl = pdf.getUrl()
  DataSheet.getRange(index+2,5).setValue(pdfUrl)

  // Email PDF as attachements.
      var message = {
        to: emailAddress,
        subject: `${eventName} Certficate of ${act}`,
        body: `Hello ${firstName} ${lastName},\n\nThis is your letter of invitation to Local Boards Meeting 2023. 
        \nYou can provide this document to your Erasmus Office for a Reimbursement. \n\nBest regards,\nThe LBM OC`,
        name: "LBM 23",
        attachments: [pdf]
      }
  
});
}