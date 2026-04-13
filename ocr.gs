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
    title: baseName + '_OCR',
    mimeType: MimeType.GOOGLE_DOCS,
  };

  const options = {
    ocr: true,
    ocrLanguage: CFG.ocr.language,
  };

  const convertedFile = Drive.Files.insert(resource, blob, options);
  console.log('OCR変換完了: ' + sourceFile.getName() + ' → DocID: ' + convertedFile.id);

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
    title: baseName,
    mimeType: targetMimeType,
  };

  const convertedFile = Drive.Files.insert(resource, blob);
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
 * 一時Google Documentを削除
 * @param {string} docId - 削除するDocument ID
 */
function deleteTemporaryDoc(docId) {
  DriveApp.getFileById(docId).setTrashed(true);
  console.log('一時Docを削除: ' + docId);
}
