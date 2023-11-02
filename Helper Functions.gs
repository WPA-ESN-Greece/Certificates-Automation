// Gets the ID of a google doc file (Doc, spredsheet, presentation, form), folder or script from its URL.
function extractDocumentIdFromUrl(url) 
{
  var parts = url.split('/')
  //Logger.log(parts[4])

  if (parts[4] == "d")
  {
    var idIndex = parts.indexOf('d') + 1
    //Logger.log(parts = url.split('/'))

    if (idIndex > 0 && idIndex < parts.length) 
    {
      //Logger.log(parts[idIndex])
      return parts[idIndex]
    } 
    else 
    {
      // If the URL doesn't contain the expected parts
      Logger.log("Invalid URL")
      return "Invalid URL"
    }
  }

  if (parts[4] == "folders" || parts[4] == "projects" )
  {
    var idIndex = 5
    //Logger.log(parts = url.split('/'))

    if (idIndex > 0 && idIndex < parts.length) 
    {
      //Logger.log(parts[idIndex])
      return parts[idIndex]
    }
    else 
    {
      // If the URL doesn't contain the expected parts
      Logger.log("Invalid URL")
      return "Invalid URL";
    }
  }

  else
  {
    Logger.log("Unknown type of URL")
    return "Unknown type of URL"
  }
}

// Gets the parent folder of the file.
function getParentFolder(fileID)
{
  var parentFolderID = DriveApp.getFileById(fileID).getParents().next().getId() // File Parent folder
  var destinationFolder = DriveApp.getFolderById(parentFolderID)

  return destinationFolder
}
 

//Create new folder function while it checks if it already exists.
function createNewFolder(parentFolderID, newFolderName)
{
  var parentFolderName = DriveApp.getFolderById(parentFolderID).getName()
  var folder = DriveApp.getFolderById(parentFolderID).getFolders()

  while(folder.hasNext()) 
  {
    var folderN = folder.next()

    if(folderN.getName() == newFolderName)
    {
      Logger.log("[" + newFolderName + "] alredy exists in " + parentFolderName)

      return folderN 
    }
  }

  var destinationFolder = DriveApp.getFolderById(parentFolderID).createFolder(newFolderName)
  Logger.log("[" + newFolderName + "] was created in " + parentFolderName)

  return destinationFolder
}
