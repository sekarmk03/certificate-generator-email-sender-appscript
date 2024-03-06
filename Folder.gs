function getFolderByName(folderName) {
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const parentFolder = DriveApp.getFileById(ssId).getParents().next();

  const subFolders = parentFolder.getFolders();
  while (subFolders.hasNext()) {
    let folder = subFolders.next();

    if (folder.getName() === folderName) {
      return folder;
    }
  }

  return parentFolder.createFolder(folderName)
  .setDescription('Created by Generate Certificate application to store PDF output files');
}