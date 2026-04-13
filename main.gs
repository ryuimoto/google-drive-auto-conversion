/**
 * PDF / 画像 / Office → Google スプレッドシート / ドキュメント 変換ツール
 * メインエントリーポイント
 */

/**
 * メイン処理: 統合アップロードフォルダをスキャンして未処理ファイルを自動振り分け変換
 * タイムドリブントリガーから定期的に呼び出される
 */
function scanAndProcessFiles() {
  if (!CFG.folders.upload) {
    console.error('CFG.folders.upload が未設定です。createFolderStructure() を実行してから設定してください');
    return;
  }

  var processedCount = 0;
  var maxFiles = CFG.processing.maxFilesPerExecution;

  var files = getUnprocessedFiles(CFG.folders.upload);
  for (var i = 0; i < files.length && processedCount < maxFiles; i++) {
    var file = files[i];
    var result = safeExecute(function() {
      return processSingleFile(file);
    }, 'process: ' + file.getName(), file.getName());

    if (result) processedCount++;
  }

  if (processedCount > 0) {
    console.log('処理完了: ' + processedCount + ' ファイル');
  }
}

/**
 * 単一ファイルを処理（MIMEタイプから自動振り分け）
 * @param {GoogleAppsScript.Drive.File} file
 * @return {boolean} 処理成功の場合true
 */
function processSingleFile(file) {
  var fileId = file.getId();
  var fileName = file.getName();
  var mimeType = file.getMimeType();
  var route = getConversionRoute(mimeType);

  console.log('処理開始: ' + fileName + ' (mimeType: ' + mimeType + ', route: ' + route + ')');

  if (route === 'skip') {
    console.warn('未対応形式のためスキップ: ' + fileName);
    return false;
  }

  if (route === 'ocr') {
    handleOcrRoute(fileId, fileName);
  } else if (route === 'toSheet') {
    handleDirectSheetRoute(fileId, fileName);
  } else if (route === 'toDoc') {
    handleDirectDocRoute(fileId, fileName);
  }

  markAsProcessed(fileId);
  console.log('処理完了: ' + fileName);
  return true;
}

/**
 * OCR経路: PDF/画像 → OCR → 内容判定 → Sheet or Doc
 * @param {string} fileId
 * @param {string} fileName
 */
function handleOcrRoute(fileId, fileName) {
  var ocrDocId = convertWithOcr(fileId);
  var text = extractTextFromDoc(ocrDocId);
  var contentType = detectContentType(text);

  console.log('内容判定: ' + fileName + ' → ' + contentType);

  if (contentType === 'invoice') {
    var invoice = parseInvoice(text);
    var sheetId = createInvoiceSheet(invoice, fileName);
    deleteTemporaryDoc(ocrDocId);
    logResult(fileName, 'success', sheetId);
    notifySuccess(fileName, sheetId, 'sheet');

  } else if (contentType === 'table') {
    var rows = parseGenericTable(text);
    var tableTitle = fileName.replace(/\.[^.]+$/, '');
    var tableSheetId = createGenericTableSheet(rows, tableTitle);
    deleteTemporaryDoc(ocrDocId);
    logResult(fileName, 'success', tableSheetId);
    notifySuccess(fileName, tableSheetId, 'sheet');

  } else {
    handleDocumentOutput(ocrDocId, fileName);
    logResult(fileName, 'success', ocrDocId);
    notifySuccess(fileName, ocrDocId, 'doc');
  }
}

/**
 * Sheet直接変換経路: Excel/CSV → Google Sheet
 * @param {string} fileId
 * @param {string} fileName
 */
function handleDirectSheetRoute(fileId, fileName) {
  var sheetId = convertOfficeFile(fileId, MimeType.GOOGLE_SHEETS);

  if (CFG.folders.output) {
    var outputFolder = DriveApp.getFolderById(CFG.folders.output);
    DriveApp.getFileById(sheetId).moveTo(outputFolder);
  }

  logResult(fileName, 'success', sheetId);
  notifySuccess(fileName, sheetId, 'sheet');
}

/**
 * Doc直接変換経路: Word/PowerPoint/Text → Google Doc
 * @param {string} fileId
 * @param {string} fileName
 */
function handleDirectDocRoute(fileId, fileName) {
  var docId = convertOfficeFile(fileId, MimeType.GOOGLE_DOCS);
  handleDocumentOutput(docId, fileName);
  logResult(fileName, 'success', docId);
  notifySuccess(fileName, docId, 'doc');
}

/**
 * 手動実行用: 指定したファイルIDを変換
 * Apps Scriptエディタから直接実行する場合に使用
 * @param {string} fileId
 */
function processManual(fileId) {
  if (!fileId) {
    console.error('fileIdを指定してください');
    return;
  }

  var file = DriveApp.getFileById(fileId);
  console.log('手動処理開始: ' + file.getName());

  safeExecute(function() {
    return processSingleFile(file);
  }, 'manual: ' + file.getName(), file.getName());

  console.log('手動処理完了');
}

// ===== トリガー管理 =====

/**
 * 定期スキャン用のタイムドリブントリガーを作成
 * 初回セットアップ時に一度だけ実行する
 */
function setupTrigger() {
  removeTrigger();

  ScriptApp.newTrigger(CFG.trigger.functionName)
    .timeBased()
    .everyMinutes(CFG.trigger.intervalMinutes)
    .create();

  console.log('トリガーを作成しました: ' + CFG.trigger.intervalMinutes + '分間隔で ' + CFG.trigger.functionName + ' を実行');
}

/**
 * 本ツールのトリガーをすべて削除
 */
function removeTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === CFG.trigger.functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
      console.log('トリガーを削除しました: ' + triggers[i].getUniqueId());
    }
  }
}

// ===== セットアップ支援 =====

/**
 * Google Driveにフォルダ構成を自動作成
 * 初回セットアップ時に実行し、作成されたフォルダIDをconfig.gsに設定する
 */
function createFolderStructure() {
  var root = DriveApp.getRootFolder();
  var parentFolder = root.createFolder('PDF自動変換');

  var uploadFolder = parentFolder.createFolder('UPLOAD');
  var processedFolder = parentFolder.createFolder('処理済み');
  var outputFolder = parentFolder.createFolder('変換結果');

  console.log('=== フォルダを作成しました ===');
  console.log('以下のIDをconfig.gsに設定してください:');
  console.log('');
  console.log('upload:    ' + uploadFolder.getId());
  console.log('processed: ' + processedFolder.getId());
  console.log('output:    ' + outputFolder.getId());
  console.log('');
  console.log('親フォルダ「PDF自動変換」: ' + parentFolder.getUrl());
}
