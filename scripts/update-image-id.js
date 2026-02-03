function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Art Portfolio')
    .addItem('Update Image IDs', 'updateImageIds')
    .addItem('Set Folder ID', 'setFolderId')
    .addToUi();
}

function getFolderId() {
  const folderId = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
  if (!folderId) {
    SpreadsheetApp.getUi().alert(
      'FOLDER_ID not set. Use "Art Portfolio > Set Folder ID" to configure it.'
    );
    return null;
  }
  return folderId;
}

function setFolderId() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Set Folder ID',
    'Enter the Google Drive folder ID containing your images:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const folderId = response.getResponseText().trim();
    if (folderId) {
      PropertiesService.getScriptProperties().setProperty('FOLDER_ID', folderId);
      ui.alert('Folder ID saved successfully.');
    }
  }
}

function findColumns(headers) {
  const normalized = headers.map(h => String(h).toLowerCase().trim());
  const imageCol = normalized.indexOf('image');
  const imageIdCol = normalized.indexOf('image_id');

  if (imageCol === -1 || imageIdCol === -1) {
    const missing = [];
    if (imageCol === -1) missing.push('image');
    if (imageIdCol === -1) missing.push('image_id');
    SpreadsheetApp.getUi().alert(
      'Missing required columns: ' + missing.join(', ')
    );
    return null;
  }

  return { imageCol, imageIdCol };
}

function updateImageIds() {
  const folderId = getFolderId();
  if (!folderId) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const columns = findColumns(data[0]);
  if (!columns) return;

  const { imageCol, imageIdCol } = columns;

  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const fileMap = {};

  // Map file name to file object (not just ID, we need to update permissions)
  while (files.hasNext()) {
    const file = files.next();
    fileMap[file.getName()] = file;
    console.log("Found file: " + file.getName());
  }

  const filenamesListed = new Set();

  for (let i = 1; i < data.length; i++) {
    const filename = data[i][imageCol];
    console.log("Processing row " + (i + 1) + ": " + filename);
    filenamesListed.add(filename);
    const range = sheet.getRange(i + 1, imageIdCol + 1);

    if (filename && fileMap[filename]) {
      const file = fileMap[filename];
      // Insert file ID
      range.setValue(file.getId());
    } else if (filename) {
      range.setValue("");
    }
  }

  // List unlisted files from folder below the data
  let row = data.length + 5;
  Object.keys(fileMap).forEach((filename) => {
    const range = sheet.getRange(row, imageCol + 1);
    if (filename.endsWith(".png") && !filenamesListed.has(filename)) {
      range.setValue(filename);
      row += 1;
    }
  });

  SpreadsheetApp.getUi().alert('Image IDs updated successfully.');
}
