/**
 * 変換モジュール
 * Drive API v2 を使用してファイルを Google Workspace 形式に変換する
 */

/**
 * PDF/画像をDrive API OCRでGoogle Documentに変換
 * @param {string} fileId - 元ファイルのDrive ID
 * @return {string} 変換後のGoogle Document ID
 */
function convertWithOcr(fileId) {
  const sourceFile = DriveApp.getFileById(fileId);
  const blob = sourceFile.getBlob();
  const baseName = sourceFile.getName().replace(/\.[^.]+$/, '');

  const resource = {
    name: baseName + '_OCR',
    mimeType: MimeType.GOOGLE_DOCS,
  };

  // v3: 変換先がGoogle DocsならPDF/画像のOCRが自動で行われる
  const options = {
    ocrLanguage: CFG.ocr.language,
  };

  const convertedFile = Drive.Files.create(resource, blob, options);
  console.log('OCR変換完了: ' + sourceFile.getName() + ' → DocID: ' + convertedFile.id);

  return convertedFile.id;
}

/**
 * PDFをOCRなしでGoogle Documentに変換（埋め込みテキスト優先）
 * デジタル生成PDFの場合は埋め込みテキストがそのまま使える
 * @param {string} fileId - 元ファイルのDrive ID
 * @return {string} 変換後のGoogle Document ID
 */
function convertWithoutOcr(fileId) {
  const sourceFile = DriveApp.getFileById(fileId);
  const blob = sourceFile.getBlob();
  const baseName = sourceFile.getName().replace(/\.[^.]+$/, '');

  const resource = {
    name: baseName + '_TEXT',
    mimeType: MimeType.GOOGLE_DOCS,
  };

  const convertedFile = Drive.Files.create(resource, blob);
  console.log('テキスト変換完了: ' + sourceFile.getName() + ' → DocID: ' + convertedFile.id);

  return convertedFile.id;
}

/**
 * Office系ファイルをGoogle Workspace形式に直接変換（OCRなし）
 * @param {string} fileId - 元ファイルのDrive ID
 * @param {string} targetMimeType - 変換先のMIMEタイプ（MimeType.GOOGLE_DOCS / GOOGLE_SHEETS）
 * @return {string} 変換後のファイルID
 */
function convertOfficeFile(fileId, targetMimeType) {
  const sourceFile = DriveApp.getFileById(fileId);
  const blob = sourceFile.getBlob();
  const baseName = sourceFile.getName().replace(/\.[^.]+$/, '');

  const resource = {
    name: baseName,
    mimeType: targetMimeType,
  };

  const convertedFile = Drive.Files.create(resource, blob);
  console.log('直接変換完了: ' + sourceFile.getName() + ' → ID: ' + convertedFile.id);

  return convertedFile.id;
}

/**
 * Google DocumentからOCRテキストを抽出
 * @param {string} docId - Google Document ID
 * @return {string} 抽出されたテキスト
 */
function extractTextFromDoc(docId) {
  const doc = DocumentApp.openById(docId);
  return doc.getBody().getText();
}

/**
 * Google Spreadsheetから全セル内容をテキストとして抽出
 * Excel/CSV由来のSheetからテキストを取り出すために使用
 * @param {string} sheetId
 * @return {string}
 */
function extractTextFromSheet(sheetId) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheets = ss.getSheets();
  var allText = [];
  for (var i = 0; i < sheets.length; i++) {
    var data = sheets[i].getDataRange().getValues();
    for (var r = 0; r < data.length; r++) {
      allText.push(data[r].join('\t'));
    }
  }
  return allText.join('\n');
}

/**
 * 一時Google Documentを削除
 * @param {string} docId - 削除するDocument ID
 */
function deleteTemporaryDoc(docId) {
  DriveApp.getFileById(docId).setTrashed(true);
  console.log('一時Docを削除: ' + docId);
}
