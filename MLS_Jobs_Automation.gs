/* CONFIG */
const TARGET_SHEET_NAME = 'Info';       
const CONFIG_SHEET_NAME = 'Config';     
const HEADER_ROW = 1;
const FOLDER_URL_COLUMN = 7;            
const DELIVERABLES_URL_COLUMN = 8;      
const FOLDER_ID_COLUMN = 9;             
const COLUMNS = {
  BID: 1,       // A
  CLIENT: 2,    // B
  TMK: 3,       // C
  ADDRESS: 4,   // D
  STATUS: 5     // E
};

/* Helpers */
function cleanName(name) {
  if (!name) return 'unnamed';
  return name.toString().replace(/[\/\\\?\%\*\:\|\"<>\.]/g, '-').trim();
}

function getRootFolder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!config) throw new Error(`Config sheet "${CONFIG_SHEET_NAME}" not found`);

  const values = config.getDataRange().getValues();
  for (let r = 0; r < values.length; r++) {
    if (values[r][0] === 'ROOT_FOLDER_ID') {
      const folderId = values[r][1];
      if (!folderId) throw new Error("ROOT_FOLDER_ID is empty in Config sheet");
      return DriveApp.getFolderById(folderId);
    }
  }
  throw new Error("ROOT_FOLDER_ID not found in Config sheet");
}

/* Menu */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('MLS Jobs')
    .addItem('Process all rows', 'processAllRows')
    .addItem('Process current row', 'processCurrentRowMenu')
    .addToUi();
}

function processCurrentRowMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
  const row = sheet.getActiveRange().getRow();
  if (row <= HEADER_ROW) {
    SpreadsheetApp.getUi().alert('Select a data row.');
    return;
  }
  createOrUpdateFolderForRow(sheet, row);
  SpreadsheetApp.getUi().alert('Processed row ' + row);
}

/* onEdit trigger */
function onEdit(e) {
  try {
    if (!e) return;
    const sheet = e.source.getActiveSheet();
    if (!sheet || sheet.getName() !== TARGET_SHEET_NAME) return;

    const row = e.range.getRow();
    if (row <= HEADER_ROW) return;

    const editedStartCol = e.range.getColumn();
    const editedNumCols = e.range.getNumColumns ? e.range.getNumColumns() : 1;
    const editedCols = [];
    for (let c = editedStartCol; c < editedStartCol + editedNumCols; c++) editedCols.push(c);

    const allowedCols = [COLUMNS.BID, COLUMNS.CLIENT, COLUMNS.TMK, COLUMNS.ADDRESS, COLUMNS.STATUS];
    if (!editedCols.some(c => allowedCols.indexOf(c) !== -1)) return;

    const bid = sheet.getRange(row, COLUMNS.BID).getValue().toString().trim();
    const client = sheet.getRange(row, COLUMNS.CLIENT).getValue().toString().trim();
    if (!bid || !client) return;

    createOrUpdateFolderForRow(sheet, row);
  } catch (err) {
    Logger.log('onEdit error: ' + err);
  }
}

/* Main logic */
function createOrUpdateFolderForRow(sheet, row) {
  try {
    function safeGet(r, c) {
      const cell = sheet.getRange(r, c);
      const d = cell.getDisplayValue();
      return (d !== undefined && d !== null && d !== '') ? d : cell.getValue();
    }

    const bid = safeGet(row, COLUMNS.BID);
    const client = safeGet(row, COLUMNS.CLIENT);
    const folderName = cleanName(bid + ' ' + client);

    // Get root MLS folder from Config
    const root = getRootFolder();

    // Ensure 2025 folder exists inside MLS
    const yearFolderName = "2025";
    const yearFolders = root.getFoldersByName(yearFolderName);
    const yearFolder = yearFolders.hasNext() ? yearFolders.next() : root.createFolder(yearFolderName);

    // Reuse folder if ID already stored in row
    let mainFolder;
    let folderId = sheet.getRange(row, FOLDER_ID_COLUMN).getValue();

    if (folderId) {
      try {
        mainFolder = DriveApp.getFolderById(folderId);
      } catch (err) {
        Logger.log("Invalid folderId in row " + row + ", creating new.");
        mainFolder = findOrCreateFolderInParent(yearFolder, folderName);
        sheet.getRange(row, FOLDER_ID_COLUMN).setValue(mainFolder.getId());
      }
    } else {
      mainFolder = findOrCreateFolderInParent(yearFolder, folderName);
      sheet.getRange(row, FOLDER_ID_COLUMN).setValue(mainFolder.getId());
    }

    // Subfolders
    const subfolderNames = ['Field', 'Drafting', 'Deliverables'];
    const subfolders = {};
    subfolderNames.forEach(name => {
      subfolders[name] = findOrCreateFolderInParent(mainFolder, name);
    });

    // Notes file
    const notesContent = buildNotes(sheet, row);
    replaceFile(mainFolder, 'Job Notes.txt', notesContent);

    // Write URLs
    sheet.getRange(row, FOLDER_URL_COLUMN).setValue(mainFolder.getUrl());
    sheet.getRange(row, DELIVERABLES_URL_COLUMN).setValue(subfolders['Deliverables'].getUrl());

  } catch (err) {
    Logger.log('createOrUpdateFolderForRow error: ' + err);
  }
}

/* Drive helpers */
function findOrCreateFolderInParent(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function replaceFile(folder, fileName, content) {
  const files = folder.getFilesByName(fileName);
  while (files.hasNext()) files.next().setTrashed(true);
  folder.createFile(fileName, content);
}

/* Notes formatting */
function buildNotes(sheet, row) {
  function safeGet(r, c) {
    const cell = sheet.getRange(r, c);
    const d = cell.getDisplayValue();
    return (d !== undefined && d !== null && d !== '') ? d : cell.getValue();
  }

  const bid = safeGet(row, COLUMNS.BID);
  const client = safeGet(row, COLUMNS.CLIENT);
  const tmk = safeGet(row, COLUMNS.TMK);
  const address = safeGet(row, COLUMNS.ADDRESS);
  const status = safeGet(row, COLUMNS.STATUS);
  const updated = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  return [
    '=== Job Notes ===',
    '',
    `Bid#:      ${bid || ''}`,
    `Client:    ${client || ''}`,
    `TMK:       ${tmk || ''}`,
    `Address:   ${address || ''}`,
    `Status:    ${status || ''}`,
    '',
    `Row #:     ${row}`,
    `Last Updated: ${updated}`
  ].join('\n');
}

/* Bulk process */
function processAllRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  for (let r = HEADER_ROW + 1; r <= lastRow; r++) {
    const bid = sheet.getRange(r, COLUMNS.BID).getValue();
    const client = sheet.getRange(r, COLUMNS.CLIENT).getValue();
    if (bid && client) createOrUpdateFolderForRow(sheet, r);
  }
}
