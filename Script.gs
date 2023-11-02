const ss = SpreadsheetApp.getActiveSpreadsheet()
const DataSheet = ss.getSheetByName('Data')
const SettingsSheet = ss.getSheetByName('Settings')

const spredsheetParentFolder = getParentFolder(ss.getId())

const Certficate_Template_Slide_URL = "https://docs.google.com/presentation/d/1TJi1S0y7cZGvmkRpac5ZCmI5M3WSR0kypSarVw310x0/edit?usp=sharing"

// Event Details
  const Event_Name = SettingsSheet.getRange("B3").getValue()
  const Event_StartDate = SettingsSheet.getRange("B6").getValue()
  const Event_EndDate = SettingsSheet.getRange("B9").getValue()
  const Event_Month = SettingsSheet.getRange("B12").getValue()
  const Event_Year = SettingsSheet.getRange("B15").getValue()
  const Event_Place = SettingsSheet.getRange("B18").getValue()

  const Event_Participating_Act = SettingsSheet.getRange("D3").getValue()
  const Event_Appreciation_Act = SettingsSheet.getRange("D4").getValue()

  let Event_Participating_Act_Text = SettingsSheet.getRange("E3").getValue()
  let Event_Appreciation_Act_Text = SettingsSheet.getRange("E4").getValue().toString()
// 

function getYourTemplateCopy()
{
  // Creates your copy of the Template. 
  let your_Certificate_Template_URL = DriveApp.getFileById(extractDocumentIdFromUrl(Certficate_Template_Slide_URL)).makeCopy(Event_Name + " Certificate Template", spredsheetParentFolder).getUrl()

  SettingsSheet.getRange("D15").setValue(your_Certificate_Template_URL)
}

function generateParticipantsCertificates()
{ 

  let certificatesFolder = createNewFolder(spredsheetParentFolder.getId(), Event_Name + " Certificates")
  let your_Certificate_Template_ID = extractDocumentIdFromUrl(SettingsSheet.getRange("D15").getValue())

  SettingsSheet.getRange("D21").setValue(certificatesFolder.getUrl())
  
  var rowRange = DataSheet.getLastRow()-1
  var lastColumn = DataSheet.getLastColumn()

  let userData = DataSheet.getRange(2, 1, rowRange, lastColumn).getValues()

  Logger.log(userData)

  userData.forEach(
    function(row, index)
    {
      if (row[0] || row[2])
      {
    
        // Mapping user data.
        let firstName = row[0]
        let lastName = row[1]
        let fullName = row[2]
        let emailAddress = row[3]
        let esnCountry = row[4]
        let esnLocalSection = row[5]
        let act = row[6]
        let teamMember = row[7]

        // Creates a full name if there wasn't one.
        if (!(fullName) && (firstName && lastName))
        {
          fullName = firstName + " " + lastName
        }

        if (teamMember)
        {
          Event_Appreciation_Act_Text = Event_Appreciation_Act_Text.replace("{{as member text}}", "as a member of the " + teamMember)
        }

        // Creates a folder for each ESN Country. 
        let Current_FolderID = createNewFolder(certificatesFolder.getId(), esnCountry)

        // Creates Cerificate and Replaces text for each participant. 
        let Current_Folder = DriveApp.getFolderById(Current_FolderID.getId())

        let currentCertificateID = DriveApp.getFileById(your_Certificate_Template_ID).makeCopy(index+1 + "." + fullName + " " + Event_Name + " Certificate", Current_Folder).getId()

        let currentCertificate_Slide = SlidesApp.openById(currentCertificateID)

        currentCertificate_Slide.replaceAllText("{{act}}", act)
        currentCertificate_Slide.replaceAllText("{{Full Name}}", fullName)
        currentCertificate_Slide.replaceAllText("{{ESN Section}}", esnLocalSection)

        if (act === Event_Participating_Act)
        {
          currentCertificate_Slide.replaceAllText("{{text}}", Event_Participating_Act_Text)
        }
        else 
        {
          currentCertificate_Slide.replaceAllText("{{text}}", Event_Appreciation_Act_Text)
        }

        currentCertificate_Slide.saveAndClose()

        let cerificate_PDF = DriveApp.getFileById(currentCertificate_Slide.getId()).getBlob().getAs(MimeType.PDF)
        let PDF = Current_Folder.createFile(cerificate_PDF).setName(fullName + " - " + Event_Name + " Certificate")
        let pdfUrl = PDF.getUrl()
        DataSheet.getRange(index +2, 9).setValue(pdfUrl) 

        DriveApp.getFileById(currentCertificate_Slide.getId()).setTrashed(true)

        //DriveApp.getFileById()

        if (emailAddress)
        {
          emailsCertificate(emailAddress, Event_Name, firstName, lastName, PDF, index+2) 
        }
      }
    })
}




function emailsCertificate(_emailAddress,_eventName,_firstName,_lastName,_pdf,_index) 
{
  // Email PDF as attachements.
      var message = {
        to: _emailAddress,
        subject: `${_eventName} Certficate for ${_eventName}`,
        body: `Hello ${_firstName} ${_lastName},\n
        This is your certificate for ${_eventName}.`,
        name: _eventName,
        attachments: [_pdf]
      }

  MailApp.sendEmail(message)

  DataSheet.getRange(_index, 4).setBackground('#b6d7a8')
}
