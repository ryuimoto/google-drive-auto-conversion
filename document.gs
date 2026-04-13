/**
 * ドキュメント出力モジュール
 * OCR変換されたGoogle Documentを整形・移動する
 */

/**
 * OCR変換後のGoogle Documentを出力フォルダへ移動
 * @param {string} ocrDocId - OCR変換で作成されたDocument ID
 * @param {string} originalFileName - 元PDFのファイル名
 * @return {string} ドキュメントのID
 */
function handleDocumentOutput(ocrDocId, originalFileName) {
  var docFile = DriveApp.getFileById(ocrDocId);

  // ファイル名を整形（_OCR接尾辞を除去し、元のファイル名ベースにリネーム）
  var cleanName = originalFileName.replace(/\.pdf$/i, '');
  docFile.setName(cleanName);

  // 出力フォルダへ移動
  if (CFG.folders.output) {
    var outputFolder = DriveApp.getFolderById(CFG.folders.output);
    docFile.moveTo(outputFolder);
  }

  console.log('ドキュメント出力完了: ' + cleanName + ' (ID: ' + ocrDocId + ')');
  return ocrDocId;
}
