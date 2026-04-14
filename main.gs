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
    console.error('初期セットアップ未完了です。GASエディタで setup() を実行してください');
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
  console.log('CFG.ledger.spreadsheetId: "' + CFG.ledger.spreadsheetId + '"');

  if (!CFG.folders.upload || !CFG.folders.processed || !CFG.folders.output || !CFG.ledger.spreadsheetId) {
    console.error('初期セットアップが未完了です。GASエディタで setup() を実行してください');
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
 * 初期セットアップを一括実行
 *
 * これ1つを実行すれば以下がすべて完了する:
 *   1. マイドライブに「Googleドライブ自動変換」フォルダ階層を作成
 *   2. 取引台帳スプレッドシートを作成
 *   3. ScriptProperties に各IDを保存（config.gs の編集は不要）
 *   4. 5分間隔のトリガーを登録
 *
 * 何度実行しても重複は作られない（冪等）。
 */
function setup() {
  console.log('===== セットアップ開始 =====');

  // 1. フォルダ構成を確保
  var root = DriveApp.getRootFolder();
  var parentFolder = getOrCreateSubfolder(root, 'Googleドライブ自動変換');
  var uploadFolder = getOrCreateSubfolder(parentFolder, 'UPLOAD');
  var processedFolder = getOrCreateSubfolder(parentFolder, '処理済み');
  var outputFolder = getOrCreateSubfolder(parentFolder, '出力');

  // 2. 取引台帳スプレッドシートを確保
  var ledgerFile = findOrCreateLedger_(parentFolder);

  // 3. ScriptProperties に保存
  var props = PropertiesService.getScriptProperties();
  props.setProperties({
    [PROP_KEYS.upload]: uploadFolder.getId(),
    [PROP_KEYS.processed]: processedFolder.getId(),
    [PROP_KEYS.output]: outputFolder.getId(),
    [PROP_KEYS.ledger]: ledgerFile.getId(),
  });

  // 4. トリガー登録（既存があれば再作成）
  setupTrigger();

  console.log('');
  console.log('===== セットアップ完了 =====');
  console.log('親フォルダ : ' + parentFolder.getUrl());
  console.log('UPLOAD    : ' + uploadFolder.getUrl());
  console.log('処理済み   : ' + processedFolder.getUrl());
  console.log('出力      : ' + outputFolder.getUrl());
  console.log('取引台帳   : ' + ledgerFile.getUrl());
  console.log('');
  console.log('使い方: 上記「UPLOAD」フォルダにPDF/画像/Officeファイルを入れると、');
  console.log('        ' + CFG.trigger.intervalMinutes + '分以内に自動変換されます。');
}

/**
 * 親フォルダ内に取引台帳スプレッドシートを取得 or 新規作成（冪等）
 * @param {GoogleAppsScript.Drive.Folder} parentFolder
 * @return {GoogleAppsScript.Drive.File}
 */
function findOrCreateLedger_(parentFolder) {
  // 既存のIDがあればそれを優先
  var existingId = PropertiesService.getScriptProperties().getProperty(PROP_KEYS.ledger);
  if (existingId) {
    try {
      return DriveApp.getFileById(existingId);
    } catch (e) {
      // ID無効化（削除されたなど） → 続けて新規作成
    }
  }

  // 親フォルダ内を名前検索
  var iter = parentFolder.getFilesByName('取引台帳');
  if (iter.hasNext()) return iter.next();

  // 新規作成
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

  var widths = [150, 200, 250, 80, 200, 100, 150, 100, 100, 100, 100, 250, 400, 80];
  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }

  var file = DriveApp.getFileById(ss.getId());
  file.moveTo(parentFolder);
  return file;
}
