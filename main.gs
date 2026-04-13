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
    handleOcrRoute(fileId, fileName, mimeType);
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
 * OCR経路: PDF/画像 → テキスト抽出 → 内容判定 → Sheet or Doc
 * PDFはまず埋め込みテキスト抽出を試み、失敗時にOCRへフォールバック
 * @param {string} fileId
 * @param {string} fileName
 * @param {string} mimeType
 */
function handleOcrRoute(fileId, fileName, mimeType) {
  var isPdf = (mimeType === 'application/pdf');
  var ocrDocId;
  var text;

  if (isPdf) {
    ocrDocId = convertWithoutOcr(fileId);
    text = extractTextFromDoc(ocrDocId);

    if (!hasUsableText(text)) {
      console.log('埋め込みテキスト不足、OCRにフォールバック: ' + fileName);
      deleteTemporaryDoc(ocrDocId);
      ocrDocId = convertWithOcr(fileId);
      text = extractTextFromDoc(ocrDocId);
    } else {
      console.log('埋め込みテキスト使用: ' + fileName + ' (' + text.length + ' chars)');
    }
  } else {
    ocrDocId = convertWithOcr(fileId);
    text = extractTextFromDoc(ocrDocId);
  }

  var contentType = detectContentType(text);
  console.log('内容判定: ' + fileName + ' → ' + contentType);

  // ビジネス書類グループ（Sheet出力）
  var businessDocTypes = {
    invoice:  '請求書',
    receipt:  '領収書',
    quote:    '見積書',
    order:    '注文書',
    delivery: '納品書',
  };

  if (businessDocTypes[contentType]) {
    var docTypeLabel = businessDocTypes[contentType];
    var invoice = parseInvoice(text);
    var sheetId = createBusinessDocSheet(invoice, fileName, docTypeLabel);
    deleteTemporaryDoc(ocrDocId);
    logResult(fileName, 'success', sheetId);
    notifySuccess(fileName, sheetId, 'sheet');
    return;
  }

  // 汎用テーブル
  if (contentType === 'table') {
    var rows = parseGenericTable(text);
    var tableTitle = fileName.replace(/\.[^.]+$/, '');
    var tableSheetId = createGenericTableSheet(rows, tableTitle);
    deleteTemporaryDoc(ocrDocId);
    logResult(fileName, 'success', tableSheetId);
    notifySuccess(fileName, tableSheetId, 'sheet');
    return;
  }

  // テキスト系書類（議事録/契約書/報告書/その他）→ Doc 出力
  var docPrefixes = {
    contract: '契約書_',
    minutes:  '議事録_',
    report:   '報告書_',
  };
  var prefix = docPrefixes[contentType] || '';
  // 元ファイル名が既に同じ接頭辞で始まっていれば重複を防ぐ
  var renamedFileName = (prefix && fileName.indexOf(prefix) === 0)
    ? fileName
    : prefix + fileName;

  handleDocumentOutput(ocrDocId, renamedFileName);
  logResult(fileName, 'success', ocrDocId);
  notifySuccess(fileName, ocrDocId, 'doc');
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

// ===== 診断 =====

/**
 * 設定とフォルダ内容を確認するデバッグ関数
 */
function diagnose() {
  console.log('===== 診断開始 =====');
  console.log('CFG.folders.upload: "' + CFG.folders.upload + '"');
  console.log('CFG.folders.processed: "' + CFG.folders.processed + '"');
  console.log('CFG.folders.output: "' + CFG.folders.output + '"');

  if (!CFG.folders.upload) {
    console.error('upload フォルダIDが空です');
    return;
  }

  try {
    var folder = DriveApp.getFolderById(CFG.folders.upload);
    console.log('uploadフォルダ名: ' + folder.getName());
    console.log('uploadフォルダURL: ' + folder.getUrl());

    var files = folder.getFiles();
    var count = 0;
    while (files.hasNext()) {
      var f = files.next();
      count++;
      var processed = isProcessed(f);
      var route = getConversionRoute(f.getMimeType());
      console.log('  - ' + f.getName() +
                  ' | mimeType: ' + f.getMimeType() +
                  ' | processed: ' + processed +
                  ' | route: ' + route);
    }
    console.log('合計ファイル数: ' + count);

    var unprocessed = getUnprocessedFiles(CFG.folders.upload);
    console.log('未処理かつ対応形式: ' + unprocessed.length + ' 件');
  } catch (e) {
    console.error('uploadフォルダにアクセスできません: ' + e.message);
  }

  console.log('===== 診断終了 =====');
}

// ===== セットアップ支援 =====

/**
 * Google Driveにフォルダ構成を自動作成
 * 初回セットアップ時に実行し、作成されたフォルダIDをconfig.gsに設定する
 */
function createFolderStructure() {
  var root = DriveApp.getRootFolder();
  var parentFolder = root.createFolder('Googleドライブ自動変換');

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
  console.log('親フォルダ「Googleドライブ自動変換」: ' + parentFolder.getUrl());
}
