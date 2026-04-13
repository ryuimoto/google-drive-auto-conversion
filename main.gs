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
  var fileLink = file.getUrl();
  var mimeType = file.getMimeType();
  var route = getConversionRoute(mimeType);

  console.log('処理開始: ' + fileName + ' (mimeType: ' + mimeType + ', route: ' + route + ')');

  if (route === 'skip') {
    console.warn('未対応形式のためスキップ: ' + fileName);
    return false;
  }

  // すべての経路を統一: テキスト抽出 → エントリ生成 → 台帳追記
  var textResult = extractTextForLedger(fileId, fileName, mimeType, route);
  var entry = parseLedgerEntry(textResult.text, fileName, fileLink);
  appendToLedger(entry);

  // 一時的に作った変換物を削除
  if (textResult.tempDocId) {
    deleteTemporaryDoc(textResult.tempDocId);
  }

  organizeProcessedFile(fileId, entry);
  logResult(fileName, 'success', CFG.ledger.spreadsheetId);
  notifySuccess(fileName, CFG.ledger.spreadsheetId, 'ledger');
  console.log('処理完了: ' + fileName);
  return true;
}

/**
 * ファイルからテキストを抽出（経路統合）
 * PDF/画像 → OCR or 埋め込みテキスト
 * Excel/CSV → Sheet経由でテキスト化
 * Word/PPT/Text → Doc経由でテキスト化
 * @param {string} fileId
 * @param {string} fileName
 * @param {string} mimeType
 * @param {string} route - 'ocr' | 'toSheet' | 'toDoc'
 * @return {{text: string, tempDocId: string}}
 */
function extractTextForLedger(fileId, fileName, mimeType, route) {
  if (route === 'ocr') {
    var isPdf = (mimeType === 'application/pdf');
    var docId, text;

    if (isPdf) {
      docId = convertWithoutOcr(fileId);
      text = extractTextFromDoc(docId);

      if (!hasUsableText(text)) {
        console.log('埋め込みテキスト不足、OCRにフォールバック: ' + fileName);
        deleteTemporaryDoc(docId);
        docId = convertWithOcr(fileId);
        text = extractTextFromDoc(docId);
      } else {
        console.log('埋め込みテキスト使用: ' + fileName + ' (' + text.length + ' chars)');
      }
    } else {
      docId = convertWithOcr(fileId);
      text = extractTextFromDoc(docId);
    }
    return { text: text, tempDocId: docId };
  }

  if (route === 'toDoc') {
    var docId2 = convertOfficeFile(fileId, MimeType.GOOGLE_DOCS);
    var text2 = extractTextFromDoc(docId2);
    console.log('Office→Doc変換: ' + fileName + ' (' + text2.length + ' chars)');
    return { text: text2, tempDocId: docId2 };
  }

  if (route === 'toSheet') {
    var sheetId = convertOfficeFile(fileId, MimeType.GOOGLE_SHEETS);
    var text3 = extractTextFromSheet(sheetId);
    console.log('Office→Sheet変換: ' + fileName + ' (' + text3.length + ' chars)');
    return { text: text3, tempDocId: sheetId };
  }

  throw new Error('未対応の経路: ' + route);
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
 * 取引台帳スプレッドシートを新規作成
 * 初回セットアップ時に1度だけ実行する
 * 実行後、ログに表示されるIDをCFG.ledger.spreadsheetIdに設定すること
 */
function createLedgerSpreadsheet() {
  var ss = SpreadsheetApp.create('取引台帳');
  var sheet = ss.getActiveSheet();
  sheet.setName(CFG.ledger.sheetName);

  var headers = [
    '処理日時', '元ファイル名', 'ファイルリンク', '書類種別',
    '取引先', '発行日', '書類番号', '合計金額', '小計', '消費税',
    '支払期限', '内容', '抽出元テキスト', 'ステータス'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#4a86c8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);

  // 列幅を調整
  sheet.setColumnWidth(1, 150);  // 処理日時
  sheet.setColumnWidth(2, 200);  // 元ファイル名
  sheet.setColumnWidth(3, 250);  // ファイルリンク
  sheet.setColumnWidth(4, 80);   // 書類種別
  sheet.setColumnWidth(5, 200);  // 取引先
  sheet.setColumnWidth(6, 100);  // 発行日
  sheet.setColumnWidth(7, 150);  // 書類番号
  sheet.setColumnWidth(8, 100);  // 合計金額
  sheet.setColumnWidth(9, 100);  // 小計
  sheet.setColumnWidth(10, 100); // 消費税
  sheet.setColumnWidth(11, 100); // 支払期限
  sheet.setColumnWidth(12, 250); // 内容
  sheet.setColumnWidth(13, 400); // 抽出元テキスト
  sheet.setColumnWidth(14, 80);  // ステータス

  // 親フォルダ（Googleドライブ自動変換）の直下に移動
  if (CFG.folders.processed) {
    var processedFolder = DriveApp.getFolderById(CFG.folders.processed);
    var parents = processedFolder.getParents();
    if (parents.hasNext()) {
      var parentFolder = parents.next();
      DriveApp.getFileById(ss.getId()).moveTo(parentFolder);
    }
  }

  console.log('=== 取引台帳を作成しました ===');
  console.log('');
  console.log('以下のIDをconfig.gsのledger.spreadsheetIdに設定してください:');
  console.log('');
  console.log(ss.getId());
  console.log('');
  console.log('URL: ' + ss.getUrl());
}


/**
 * Google Driveにフォルダ構成を自動作成
 * 初回セットアップ時に実行し、作成されたフォルダIDをconfig.gsに設定する
 */
function createFolderStructure() {
  var root = DriveApp.getRootFolder();
  var parentFolder = root.createFolder('Googleドライブ自動変換');

  var uploadFolder = parentFolder.createFolder('UPLOAD');
  var processedFolder = parentFolder.createFolder('処理済み');

  console.log('=== フォルダを作成しました ===');
  console.log('以下のIDをconfig.gsに設定してください:');
  console.log('');
  console.log('upload:    ' + uploadFolder.getId());
  console.log('processed: ' + processedFolder.getId());
  console.log('');
  console.log('親フォルダ「Googleドライブ自動変換」: ' + parentFolder.getUrl());
}
