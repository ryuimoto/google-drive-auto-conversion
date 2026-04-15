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

  // 前回実行が長引いて次のトリガーが重なると、同名フォルダを複数作成する競合が起きる。
  // スクリプトロックで並行実行を防止(10秒以内にロックが取れなければスキップ)
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10 * 1000)) {
    console.log('他のトリガー実行中のためスキップ');
    return;
  }

  try {
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
  } finally {
    lock.releaseLock();
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

  // Gemini 統計
  try {
    var stats = getGeminiStats();
    console.log('');
    console.log('[Gemini]');
    console.log('  有効: ' + stats.enabled);
    console.log('  APIキー設定: ' + (stats.hasKey ? '済' : '未'));
    console.log('  本日の呼び出し: ' + stats.count + ' / ' + stats.limit);
    console.log('  本日のトークン消費: ' + stats.totalTokens.toLocaleString() +
                ' (入力=' + stats.promptTokens.toLocaleString() +
                ' / 出力=' + stats.outputTokens.toLocaleString() + ')');
    console.log('  概算コスト: $' + stats.estimatedCostUsd.toFixed(6) +
                ' (無料枠 ' + stats.limit + 'req/日 以内なら実コスト $0)');
    if (stats.count > 0) {
      var avgTokensPerCall = Math.round(stats.totalTokens / stats.count);
      console.log('  1回あたり平均: ' + avgTokensPerCall.toLocaleString() + ' tokens');
    }
    if (stats.enabled && !stats.hasKey) {
      console.log('  → API を有効化するには ScriptProperties の GEMINI_API_KEY にキーを設定してください');
      console.log('  → キー取得: https://aistudio.google.com/apikey (無料)');
    }
  } catch (e) {
    console.warn('Gemini統計取得失敗: ' + e.message);
  }

  console.log('===== 診断終了 =====');
}

/**
 * Gemini API の動作確認
 * 架空の請求書テキストをGeminiに送って正しくパースできるかを確認する
 */
function testGemini() {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    console.error('GEMINI_API_KEY が未設定です。');
    console.log('1. https://aistudio.google.com/apikey でAPIキーを取得(無料)');
    console.log('2. GASエディタ → プロジェクトの設定 → スクリプトプロパティ で GEMINI_API_KEY に設定');
    return;
  }

  var sampleText =
    '請求書\n' +
    'No. TEST-001 / 2026/04/15\n' +
    'テスト株式会社 御中\n' +
    'サンプル商事株式会社\n' +
    '品目 数量 単価 金額\n' +
    '商品A 2 ¥1,000 ¥2,000\n' +
    '商品B 1 ¥3,000 ¥3,000\n' +
    '小計 ¥5,000\n' +
    '消費税 ¥500\n' +
    '合計 ¥5,500';

  console.log('===== Gemini テスト開始 =====');
  console.log('送信テキスト:');
  console.log(sampleText);
  console.log('');

  var result = extractInvoiceWithGemini(sampleText);
  if (!result) {
    console.error('Gemini 呼び出し失敗。ログを確認してください。');
    return;
  }

  console.log('結果:');
  console.log(JSON.stringify(result, null, 2));
  console.log('===== Gemini テスト完了 =====');
}

/**
 * このAPIキーで利用可能な Gemini モデル一覧を表示する
 * 404/429 でモデル呼び出しが失敗した場合、この関数で実際に呼び出せるモデルを確認する
 */
function listGeminiModels() {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    console.error('GEMINI_API_KEY が未設定です');
    return;
  }

  var url = 'https://generativelanguage.googleapis.com/v1beta/models?key=' + encodeURIComponent(apiKey);
  var response;
  try {
    response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  } catch (e) {
    console.error('ネットワークエラー: ' + e.message);
    return;
  }

  var code = response.getResponseCode();
  if (code !== 200) {
    console.error('ListModels HTTP ' + code + ': ' + response.getContentText());
    return;
  }

  var body = JSON.parse(response.getContentText());
  var models = body.models || [];
  var generative = models.filter(function(m) {
    return (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0;
  });

  console.log('===== 利用可能な Gemini モデル(' + generative.length + '件) =====');
  generative.forEach(function(m) {
    var shortName = m.name.replace(/^models\//, '');
    console.log('  ' + shortName + '  (' + (m.displayName || '-') + ')');
  });
  console.log('');
  console.log('現在の設定: ' + CFG.gemini.model);
  console.log('上のリストから generateContent 対応モデルを選び、CFG.gemini.model に設定してください');
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

  // Gemini API の案内(オプション)
  var hasGeminiKey = !!props.getProperty('GEMINI_API_KEY');
  if (!hasGeminiKey) {
    console.log('');
    console.log('[オプション] 抽出精度を最大化するには Gemini API の無料キーを設定してください:');
    console.log('  1. https://aistudio.google.com/apikey でAPIキーを発行(無料)');
    console.log('  2. GASエディタ → プロジェクトの設定 → スクリプトプロパティ');
    console.log('  3. GEMINI_API_KEY にキー文字列を貼り付けて保存');
    console.log('  → 設定後、testGemini() で動作確認できます');
  }
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

  // 金額列（合計8, 小計9, 消費税10）はカンマ区切り表示
  sheet.getRange('H2:J').setNumberFormat('#,##0');

  var file = DriveApp.getFileById(ss.getId());
  file.moveTo(parentFolder);
  return file;
}

/**
 * 処理済みフォルダ配下の同名サブフォルダ(請求書/注文書/契約書 等)を統合する
 * 過去の重複実行や古いコードで作られた複数フォルダを1つにまとめ、
 * 中身を全部最古のフォルダに集約し、空になった重複フォルダはゴミ箱に送る。
 *
 * 冪等: 重複がなければ何もしない。いつでも実行可能。
 */
function consolidateProcessedFolders() {
  if (!CFG.folders.processed) {
    console.error('CFG.folders.processed が未設定です。setup() を実行してください');
    return;
  }

  console.log('===== 処理済みフォルダ統合開始 =====');
  var processedRoot = DriveApp.getFolderById(CFG.folders.processed);

  // 直下の全サブフォルダを名前でグルーピング
  var foldersByName = {};
  var iter = processedRoot.getFolders();
  while (iter.hasNext()) {
    var f = iter.next();
    var name = f.getName();
    if (!foldersByName[name]) foldersByName[name] = [];
    foldersByName[name].push(f);
  }

  var mergedCount = 0;
  for (var name in foldersByName) {
    var group = foldersByName[name];
    if (group.length <= 1) continue;

    console.log('同名フォルダ「' + name + '」を ' + group.length + ' 個発見');

    // 最初のフォルダを「正」とし、残りの中身を移動
    var canonical = group[0];
    for (var i = 1; i < group.length; i++) {
      var dup = group[i];
      mergeFolderContents_(dup, canonical);
      dup.setTrashed(true);
      mergedCount++;
      console.log('  → 統合: ' + dup.getId() + ' を ' + canonical.getId() + ' に集約してゴミ箱へ');
    }
  }

  console.log('===== 統合完了: ' + mergedCount + ' 個の重複フォルダを統合 =====');
}

/**
 * source フォルダの中身(ファイル+サブフォルダ)を target フォルダへ再帰的にマージ
 * 同名サブフォルダがあれば再帰的に統合、同名ファイルがあれば (2), (3)... で一意化
 * @param {GoogleAppsScript.Drive.Folder} source
 * @param {GoogleAppsScript.Drive.Folder} target
 */
function mergeFolderContents_(source, target) {
  // ファイルを移動
  var files = source.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    var uniqueName = ensureUniqueName(target, file.getName());
    if (uniqueName !== file.getName()) file.setName(uniqueName);
    file.moveTo(target);
  }

  // サブフォルダを再帰的にマージ
  var subIter = source.getFolders();
  while (subIter.hasNext()) {
    var sub = subIter.next();
    var targetSub = getOrCreateSubfolder(target, sub.getName());
    mergeFolderContents_(sub, targetSub);
    sub.setTrashed(true);
  }
}
