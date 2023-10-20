/**
 * Преобразует буквенное имя столбца в его номер.
 *
 * @param {string} columnLetter - Буквенное имя столбца, например, "A" или "AB".
 * @returns {number} - Номер столбца, начиная с 0.
 */
function letterToIndex(columnLetter) {
  const base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  let result = 0;
  let multiplier = 1;

  columnLetter = columnLetter.toUpperCase();

  for (let i = columnLetter.length - 1; i >= 0; i--) {
    const char = columnLetter.charAt(i);
    const charValue = base.indexOf(char) + 1;
    result += charValue * multiplier;
    multiplier *= 26;
  }
  return result - 1;
}


/**
 * Returns a Google Drive folder in the same location
 * in Drive where the spreadsheet is located. First, it checks if the folder
 * already exists and returns that folder. If the folder doesn't already
 * exist, the script creates a new one. The folder's name is set by the
 * "OUTPUT_FOLDER_NAME" variable from the Code.gs file.
 *
 * @param {string} folderName - Name of the Drive folder.
 * @return {object} Google Drive Folder
 *
 * https://developers.google.com/apps-script/samples/automations/generate-pdfs#utilities.gs
 */
function getFolderByName_(folderName) {

  // Gets the Drive Folder of where the current spreadsheet is located.
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const parentFolder = DriveApp.getFileById(ssId).getParents().next();

  // Iterates the subfolders to check if the PDF folder already exists.
  const subFolders = parentFolder.getFolders();
  while (subFolders.hasNext()) {
    let folder = subFolders.next();

    // Returns the existing folder if found.
    if (folder.getName() === folderName) {
      return folder;
    }
  }
  // Creates a new folder if one does not already exist.
  return parentFolder.createFolder(folderName)
    .setDescription(`Created by application to store PDF output files`);
}


/**
 * Получает URL-ссылку на ZIP-архив содержимого папки в Google Диске.
 *
 * @param {string} folderId - Идентификатор целевой папки.
 * @returns {string} - URL-ссылка для скачивания ZIP-архива.
 */
function getFolderLink(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const blobs = [];
  while (files.hasNext()) {
    blobs.push(files.next().getBlob());
  }
  const zipBlob = Utilities.zip(blobs, folder.getName() + ".zip");
  const fileId = DriveApp.createFile(zipBlob).getId();

  return "https://drive.google.com/uc?export=download&id=" + fileId;
}

const A4_HEIGHT = 1123;
const EMPTY_ROW_HEIGHT = 21;

/**
 * Получает суммарную высоту всех строк в указанном листе.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Лист, для которого необходимо получить высоту строк.
 * @returns {number} - Суммарная высота строк в пикселях.
 */
function getHeight(sheet) {
  const lastRow = sheet.getLastRow();
  let rowsHeight = 0;
  for (let i = 1; i <= lastRow; i++) {
    rowsHeight += sheet.getRowHeight(i);
  }

  return rowsHeight
}

/**
 * Рассчитывает количество пустых строк, которые необходимо добавить на указанный лист, чтобы заполнить его до размера страницы A4.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Лист, для которого рассчитывается количество пустых строк.
 * @returns {number} - Количество пустых строк, необходимых для заполнения листа до размера страницы A4.
 */
function calculatePageFeed(sheet) {
  const currentHeight = getHeight(sheet);
  const reminder = currentHeight % A4_HEIGHT;

  return Math.ceil((A4_HEIGHT - reminder) / EMPTY_ROW_HEIGHT);
}

/**
 * Returns merged pdf, blobs are merged in the same order they are proivded.
 * @param {Blob[]} blobs Blob array
 * @param {String} fileName output PDF name
 * @return {Promise} Promise object, blob of merged blobs
 *
 * https://stackoverflow.com/questions/15414077/merge-multiple-pdfs-into-one-pdf
 */
async function mergeAllPDFs(blobs, fileName) {
  eval(UrlFetchApp.fetch("https://unpkg.com/pdf-lib/dist/pdf-lib.min.js").getContentText());
  setTimeout = (func, sleep) => (Utilities.sleep(sleep), func())

  const pdf = await PDFLib.PDFDocument.create();
  for (let i = 0; i < blobs.length; i++) {
    const tempBytes = await new Uint8Array(blobs[i].getBytes());
    const tempPdf = await PDFLib.PDFDocument.load(tempBytes);

    const pageCount = tempPdf.getPageCount()
    const pageIndicesArray = [...Array(pageCount).keys()]
    const pages = await pdf.copyPages(tempPdf, pageIndicesArray)
    pages.forEach(page => pdf.addPage(page))
  }
  const pdfDoc = await pdf.save()
  return Utilities.newBlob(pdfDoc).setName(fileName)
}



