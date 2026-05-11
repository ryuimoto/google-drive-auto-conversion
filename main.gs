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
 *
 * パース/抽出のいずれかの段階で例外が発生しても、可能な限りの情報(ファイル名・rawText・
 * エラー内容)で 台帳に NG 行を必ず書き込む。これにより「処理は走ったが台帳に記録が無い」
 * という不可視の失敗を防ぐ。
 *
 * @param {GoogleAppsScript.Drive.File} file
 * @return {boolean} 台帳に行が書かれた場合 true(成功・失敗いずれも)
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
  // テキスト抽出 / パースのいずれかで例外が出ても、最低限の NG 行を書くために
  // ローカル try/catch で 2 段階に分ける
  var text = '';
  var tempDocId = '';
  var entry = null;
  var processError = null;
  try {
    var textResult = extractTextForLedger(fileId, fileName, mimeType, route);
    text = textResult.text || '';
    tempDocId = textResult.tempDocId;
    entry = parseLedgerEntry(text, fileName, fileLink);
  } catch (e) {
    processError = e;
    console.error('[parse] ' + fileName + ' / ' + e.message);
    if (e.stack) console.error(e.stack);
  }

  // 失敗時のフォールバック: rawText とエラー内容だけでも記録する
  if (!entry) {
    entry = {
      processedAt: new Date(),
      fileName: fileName,
      fileLink: fileLink,
      docType: 'その他',
      vendor: '',
      issueDate: '',
      docNumber: '',
      total: 0,
      subtotal: 0,
      tax: 0,
      paymentDueDate: '',
      contentSummary: '',
      rawText: text,
      status: 'NG: ' + (processError ? processError.message : 'パース失敗'),
    };
  }

  // 台帳への書き込みは最低限保証する(ここで失敗すると上層 safeExecute に伝播)
  appendToLedger(entry);

  // 一時的に作った変換物を削除(失敗しても処理継続)
  if (tempDocId) {
    try { deleteTemporaryDoc(tempDocId); }
    catch (delErr) { console.warn('一時ファイル削除失敗: ' + delErr.message); }
  }

  // パース/抽出失敗時は処理済みフォルダへ移動して同じファイルを無限再処理しないようにする
  if (processError) {
    try { organizeProcessedFile(fileId, entry); }
    catch (orgErr) { console.warn('processed フォルダ移動失敗: ' + orgErr.message); }
    logResult(fileName, 'error', processError.message);
    if (CFG.notification.enabled && CFG.notification.notifyOnError) {
      try { notifyError(fileName, processError.message + '\n\n' + (processError.stack || '')); }
      catch (notifyErr) { console.warn('通知失敗: ' + notifyErr.message); }
    }
    console.log('処理失敗(台帳には NG 行を記録): ' + fileName);
    return true;
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

/**
 * Gemini API を呼び出さない (正規表現パースのみで完結) で関数を実行する
 * 加えて通知メールも抑止する(GAS のメール送信は 100通/日 のクォータがあり、
 * バッチ再処理では簡単に超えるため)
 * @param {Function} fn
 * @return {*}
 */
function withoutGemini_(fn) {
  var savedGemini = CFG.gemini.enabled;
  var savedNotify = CFG.notification.enabled;
  CFG.gemini.enabled = false;
  CFG.notification.enabled = false;
  try {
    return fn();
  } finally {
    CFG.gemini.enabled = savedGemini;
    CFG.notification.enabled = savedNotify;
  }
}

/**
 * scanAndProcessFiles の Gemini 不使用版
 * UPLOAD フォルダの未処理ファイルを処理するが、Gemini にはフォールバックしない。
 * 通知メールも抑止する(大量処理時のメールクォータ超過対策)
 */
function scanAndProcessFilesNoGemini() {
  console.log('Gemini 抑止モード(通知メールも抑止)で実行します');
  withoutGemini_(scanAndProcessFiles);
}

/**
 * processManual の Gemini 不使用版
 * @param {string} fileId
 */
function processManualNoGemini(fileId) {
  console.log('Gemini 抑止モード(通知メールも抑止)で手動処理します');
  withoutGemini_(function() { processManual(fileId); });
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
  console.log('CFG.ledger.spreadsheetId: "' + CFG.ledger.spreadsheetId + '"');

  if (!CFG.folders.upload || !CFG.folders.processed || !CFG.ledger.spreadsheetId) {
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

  // 取引台帳の最近の OK 率(精度の定量化)
  try {
    var ledger = SpreadsheetApp.openById(CFG.ledger.spreadsheetId);
    var sheet = ledger.getSheetByName(CFG.ledger.sheetName);
    if (sheet) {
      var lastRow = sheet.getLastRow();
      var sampleSize = Math.min(50, lastRow - 1);
      if (sampleSize > 0) {
        // ステータス列(14列目)を末尾から取得
        var statusCol = 14;
        var values = sheet.getRange(lastRow - sampleSize + 1, statusCol, sampleSize, 1).getValues();
        var okCount = 0;
        for (var si = 0; si < values.length; si++) {
          if (String(values[si][0]) === 'OK') okCount++;
        }
        var rate = (okCount / sampleSize * 100).toFixed(1);
        console.log('');
        console.log('[台帳精度] 直近 ' + sampleSize + ' 件の OK 率: ' + rate + '% (' + okCount + '/' + sampleSize + ')');
      }
    }
  } catch (e) {
    console.warn('台帳精度の取得失敗: ' + e.message);
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
 * 抽出精度の回帰テスト
 *
 * 台帳に誤記録された行の rawText 列を TEST_CASES に貼り付け、期待値を書く。
 * Gemini を抑止した状態で parseLedgerEntry を呼び出し、各項目が期待通りに
 * 取れるかを確認する。GASエディタから testParser() を直接実行する。
 *
 * ケース例:
 *   {
 *     name: '請求書#1',
 *     text: '<台帳の rawText 列をそのままコピー>',
 *     expect: {
 *       vendor: '株式会社サンプル',
 *       total: 110000,
 *       subtotal: 100000,
 *       tax: 10000,
 *       issueDate: '2026/04/15',
 *       docNumber: 'INV-2026-001',
 *       docType: '請求書',
 *     },
 *   }
 *
 * expect に書いたキーのみ検証されるため、確認したい項目だけ指定すれば良い。
 */
var TEST_CASES = [
  {
    name: 'Row1: 合計が書類番号の数字を拾う',
    text: '株式会社スカイ 御中 下記のとおりご請求申し上げます。 ご請求金額\n' +
          'お支払期限：2026年05月25日\n' +
          '2026年04月01日\n' +
          '請求書番号: 20260403-2 請 求 書\n' +
          '¥ 12,100 anono design 竹下菜生 登録番号: T8810710649516 anonodesign52@gmail.com 品目\n' +
          '数量\n単価\n金額\n' +
          'CTA画像作成費（2枚セット） 2 5,500 11,000\n' +
          '小計 11,000\n' +
          '消費税 (10%) 1,100\n' +
          '合計 12,100',
    expect: { vendor: '株式会社スカイ', total: 12100, subtotal: 11000, tax: 1100, issueDate: '2026/4/1', docNumber: '20260403-2' },
  },
  {
    name: 'Row3: 御中なき宛名(冒頭+〒)、備考の他社名に騙されない',
    text: '請求書\n' +
          '株式会社スカイ \n' +
          '〒150-0031 東京都渋谷区桜丘町1番2号\n' +
          '渋谷サクラステージセントラルビル9階\n' +
          '下記のとおりご請求申し上げます。 合計金額 ￥4,950- お支払い期限： 2026年4月末日\n' +
          '請求日：2026年4月8日\n' +
          '木山 彩花\n' +
          'TEL:080-6921-0494 〒475-0918 愛知県半田市雁宿町3-203-4 パークサイド雁宿1-702 品名 数量 単価 金額\n' +
          '2026/3/1 ショート動画作成 30 ¥50 ¥1500\n' +
          '税(10%) ¥450 合計 ¥4950 備考 株式会社ラオルカ様より 2026年3月依頼分',
    expect: { vendor: '株式会社スカイ', total: 4950 },
  },
  {
    name: 'Row10: 一般社団法人プレフィックス',
    text: '一般社団法人　お金の学校 様\n' +
          '件名 : 3月度　Bコミさんプロジェクト報酬 下記のとおりご請求申し上げます。 請求書\n' +
          '株式会社FIRE 〒4110022 静岡県三島市川原ケ谷\n' +
          '62-3 2026年04月15日\n' +
          '請求番号: 20260415-002 ご請求金額 ¥ 8,975',
    expect: { vendor: '一般社団法人お金の学校', docNumber: '20260415-002' },
  },
  {
    name: 'Row12: OCR で ¥X, YYY のスペース割れ',
    text: '請　求　書\n' +
          '株式会社スカイ 御中 DMS038 請求No. ご担当： 大塚　悠大 様 請求日\n' +
          '件名： トレフル・エデュケーションラボ　田村幸仁 令和8年3月度業務委託費\n' +
          '令和8年4月1日\n' +
          'ご請求合計金額 ¥222, 672 （税込） お支払期限： 令和8年4月30日\n' +
          'No. 摘要 数量 費用明細 単価 金額\n' +
          '1 令和８年3月『負けないデイトレスクール』DMS、及び DMCサポート業務代として\n' +
          '1 スクールサポート業 務一式 223,141 ¥223,141\n' +
          '小計 ¥223,141\n' +
          '消費税額（１０％対象) ¥22,314\n' +
          'ご請求金額 ¥245, 455',
    expect: { total: 245455, subtotal: 223141, tax: 22314 },
  },
  {
    name: 'Row15: 複数行ラベル(小計\\n消費税 合計金額\\n値3つ)',
    text: '請 求 書\n' +
          '株式会社 スカイ 御中日付：2026年4月5日 請求書番号：INV-0000000030 所在地： 小西 歩\n' +
          '下記のとおりご請求申し上げます。 合計金額 110,000円 (税込) 振込期限：2026年4月25日\n' +
          '詳細 数量 単価 金額 2026年3月分報酬 1 100,000 100,000 備考： この度はご利用いただきありがとうございました。\n' +
          '小計\n' +
          '消費税 合計金額\n' +
          '100,000 10,000 110,000',
    expect: { vendor: '株式会社スカイ', subtotal: 100000, tax: 10000, total: 110000, issueDate: '2026/4/5' },
  },
  {
    name: 'Row21: 多列ラベル「合計 税抜 消費税 総額」と 3 値',
    text: '請 求 書\n' +
          '〒150-0031 発行日： 2026/4/1(水) 東京都渋谷区桜丘町1番2号 お支払期限： 2026/4/30(木) 渋谷サクラステージ セントラ\n' +
          '株式会社スカイ 御中 奥名　貴士 〒166-0003 ご請求金額 ¥220,000 (税込) 東京都杉並区 高円寺南2-41-10\n' +
          '下記のとおりご請求申し上げます。 適格請求書発行事業者　登録番号： T9810103828081 品目 数量(時間) 金額（税別） 金額（税込） 備考\n' +
          '【検定ビジネス案件】 1 カ月分 ¥200,000 ¥220,000 単価：200,000円/月（税別） PM業務 3月分\n' +
          '合計 税抜 消費税 総額(税込) ¥200,000) ¥20,000) ¥220,000 振込先\n' +
          '楽天銀行　ドラム支店',
    expect: { subtotal: 200000, tax: 20000, total: 220000 },
  },
  {
    name: 'Row22: 源泉徴収税額を消費税と取り違えない / Unicodeハイフン',
    text: '株式会社スカイ 御中\n' +
          '件名 : 2026年3月度デザイン制作費ご請求 下記のとおりご請求申し上げます。 請求書\n' +
          'Tell Global Design Works 山本 朋実\n' +
          '〒5100802 四日市市三ツ谷町4-22 2026年04月02日\n' +
          '請求番号: 20260402‑001 ご請求金額 ¥ 50,000 - お支払い期限 : 2026年04月30日\n' +
          'TEL: 080-6907-1614\n' +
          'tomomi.araki.614y@gmail.com 品番・品名 数量 単価 金額\n' +
          'ニュースレターデザインご対応費 1 55,685 55,685 小計 55,685 源泉徴収税額 -5,685 合計 50,000',
    expect: { docNumber: '20260402-001', total: 50000, tax: 0 },
  },
  {
    name: 'Row5: 差引請求額(源泉徴収後)を合計と取り違えない',
    text: '3 月分 請求書\n' +
          '株式会社スカイ　御中\n' +
          '▼請求日： 2026/03/31 下記のとおり、ご請求もうしあげます。 ご請求金額合計 ¥4,092 ▼請求内容： ショート動画作業料一式\n' +
          '・ 郵便番号： 940-2059 ▼振込先情報：\n' +
          '作業料の内訳項目 数量 単価（税抜） 金額\n' +
          'ショート動画作成 82 ¥50 ¥4,100\n' +
          '作業料計（税抜） ¥4,100\n' +
          '消費税額 ¥410\n' +
          '合計（税込） ¥4,510\n' +
          '源泉徴収税額 （適用税率： 10.21% ） ¥418\n' +
          '差引請求額 ¥4,092',
    expect: { total: 4510, tax: 410, subtotal: 4100 },
  },
  // ===== Phase 4 ケース =====
  {
    name: 'Row2: ヘッダ「№ ID 株式会社スカイ 御中」/「合計(税込) 値 値 うち消費税」両端ラベル',
    text: '№ H2026-03-09 株式会社スカイ　御中\n' +
          '〒150-0031 東京都渋谷区桜丘町1番2号\n' +
          '請 求 書\n' +
          'ご請求日 2026年3月31日 発行日 2026年4月1日 お支払期限 2026年4月30日\n' +
          'ご請求金額\n' +
          '¥129,300 品目 単価(税込) 金額(税込) 日付 数量\n' +
          '合計(税込) 129,300 11,755 うち消費税\n' +
          '129,300 11,755\n' +
          '(10%対象 うち消費税 )',
    expect: { vendor: '株式会社スカイ', total: 129300, tax: 11755, docNumber: 'H2026-03-09' },
  },
  {
    name: 'Row14: 行中段の 株式会社X 御中 (40文字前置プレフィックスを取り込まない)',
    text: 'ご請求書\n' +
          '東京都渋谷区桜丘町1-2 渋谷サクラステージ セントラルビル9階 　発行日　 2026年4月1日 株式会社スカイ　御中 登録番号\n' +
          '株式会社アンドキュー 〒252-0311 神奈川県相模原市南区東林間1-24-7 担当・田尻賢　090-6047-7788\n' +
          '合計金額 ¥144,000 　（税込） 品名 数量 金額\n' +
          '広報・ブランディング業務 100,000\n' +
          '消費税 10% 10,000',
    expect: { vendor: '株式会社スカイ', total: 144000 },
  },
  {
    name: 'Row13: 御中位置崩れた OCR (株式会社SKY ... 御中 株式会社SIX) — 冒頭 corp を recipient とする',
    text: '株式会社SKY 下記の通りご請求申し上げます。 発行日：2026年 4月 10日\n' +
          '御請求書\n' +
          '御中 株式会社SIX 〒 852-8036 長崎県長崎市青山町10-18 適格事業者登録番号：T9290001096059 ご請求金額\n' +
          '\\936,643\n' +
          'NO 商品名 数量 単価 金額\n' +
          '1 Bコミさんプロジェクト報酬 1 \\786,643 \\786,643\n' +
          '小計 \\936,643\n' +
          '合計 \\936,643',
    expect: { vendor: '株式会社SKY', total: 936643 },
  },
  {
    name: 'Row7: 税込合計金額(値なしヘッダ) → 「合計 60,000」を採用',
    text: '御請求書\n' +
          '株式会社スカイ 御中 塩崎　弘美 〒166-0004 令和8年4月15日\n' +
          '件名 業務委託契約3月分\n' +
          'TEL 080-5124-4669 毎度ありがとうございます。 下記の通り、ご請求申し上げます。\n' +
          '¥60,000 消費税（10%）\n' +
          '税込合計金額 内税\n' +
          '品名 数量 単価 金額 備考\n' +
          '業務委託契約3月分 1 50,000 50,000\n' +
          'リスト作成・その他業務 1 10,000 10,000\n' +
          '合計 60,000',
    expect: { vendor: '株式会社スカイ', total: 60000 },
  },
  {
    name: 'Row3-2: 書類番号なし請求書で明細の No.41-45 を拾わない',
    text: '請求書\n' +
          '株式会社スカイ \n' +
          '〒150-0031 東京都渋谷区桜丘町1番2号\n' +
          '下記の通りご請求申し上げます。 合計金額 ￥4,950 お支払い期限： 2026年4月末日\n' +
          '請求日：2026年4月8日\n' +
          '木山 彩花\n' +
          '品名 数量 単価 金額\n' +
          '2026/3/1 ショート動画作成「ライオン父さん」No.41-45 30 ¥50 ¥1500\n' +
          '2026/3/3 ショート動画作成「黒字社長」No.81-90 60 ¥50 ¥3000\n' +
          '税(10%) ¥450 合計 ¥4950',
    expect: { vendor: '株式会社スカイ', total: 4950, docNumber: '' },
  },
  // ===== Phase 5 ケース =====
  {
    name: 'Row2-P5: ご請求日と発行日が併記 → 発行日を採用',
    text: '№\nH2026-03-09\n株式会社スカイ 御中\n〒150-0031\n東京都渋谷区桜丘町1番2号渋谷サクラステージセントラルビル9階 代表取締役\n山本智也 様\n請 求 書\n\nご請求日 2026年3月31日 発行日 2026年4月1日 お支払期限 2026年4月30日\nご請求金額\n¥129,300\n品目 単価(税込) 金額(税込) 日付 数量 バズぬき君\n5,000 10,000\n合計(税込) 129,300 11,755\nうち消費税\n129,300 11,755',
    expect: { vendor: '株式会社スカイ', issueDate: '2026/4/1', total: 129300, tax: 11755, docNumber: 'H2026-03-09' },
  },
  {
    name: 'Row5-P5(本番): 合計（税込）¥4,510 を total に採用',
    text: '3 月分 請求書\n株式会社スカイ 御中\n▼請求日： 2026/03/31\n下記のとおり、ご請求もうしあげます。\nご請求金額合計 ¥4,092\n▼請求内容： ショート動画作業料一式 ▼請求者情報：\n・ 郵便番号： 940-2059\n▼振込先情報： ・ ご住所： 新潟県長岡市大荒戸町２６２ -１ ・ 金融機関名： 楽天銀行\n・ 支店名： コード支店 ・ 電話番号： 080-2107-1620 ・ 口座種類： 普通 ・ メールアドレス： aoirunta@gmail.com ・ 口座番号\n登録番号（13桁）： 有： T\n作業料の内訳項目 数量 単価（税抜） 金額\nショート動画作成 82 ¥50 ¥4,100 ¥0\n¥0\n作業料計（税抜） ¥4,100 消費税額 ¥410\n合計（税込） ¥4,510 源泉徴収税額 （適用税率： 10.21% ） ¥418 差引請求額 ¥4,092',
    expect: { vendor: '株式会社スカイ', total: 4510 },
  },
  {
    name: 'Row8-P5(悠楽): 多重「合計」+「請求書No. 202604-027」',
    text: '御請求書\n発行日： 2026/4/15\n請求書No. 202604-027\n株式会社スカイ 御中\n合計 ¥550,000) 税込み\n登録番号： T8011101091283\n株式会社悠楽\n東京都新宿区西新宿7-7-25\nワコーレ新宿第二ビル３階 下記の通りご請求申し上げます。 Tel. : 03-\n但し、コンサルティング代として。\nNo. 項目 数量 単位 単価 金額 備考 1 経営コンサルティング（2026年5月分） 1 件 500,000 500,000\n小計 500,000\n上記明細をご確認後、下記口座に\n4月25日までに振込みをお願い致します。\n消費税 50,000 合計 550,000 内訳 10%対象額 500,000',
    expect: { vendor: '株式会社スカイ', total: 550000, subtotal: 500000, tax: 50000, docNumber: '202604-027' },
  },
  {
    name: 'Row10-P5(本番): 一般社団法人 + corp/〒 が別行で OCR 出力',
    text: '一般社団法人 お金の学校 様\n件名 : 3月度 Bコミさんプロジェクト報酬 下記のとおりご請求申し上げます。\n請求書\n株式会社FIRE\n〒4110022\n静岡県三島市川原ケ谷\n62-3\n2026年04月15日\n請求番号: 20260415-002\nご請求金額 ¥ 8,975 - お支払い期限 : 2026年04月30日\nTEL: 09035014980\n登録番号: T1011201022913\n品番・品名 数量 単価 金額\n3月度 Bコミさんプロジェクト報酬 1 8,975 8,975\n小計 8,975\n消費税 (10% 内税) (815)\n合計 8,975',
    expect: { vendor: '一般社団法人お金の学校', docNumber: '20260415-002', total: 8975 },
  },
  {
    name: 'Row12-P5(本番DMS): ¥245, 455 のスペース割れ',
    text: '請 求 書\n株式会社スカイ 御中 DMS038 請求No.\nご担当： 大塚 悠大 様 請求日\n件名： トレフル・エデュケーションラボ 田村幸仁 令和8年3月度業務委託費\n令和8年4月1日\nトレフル\n下記の通り、ご請求申し上げます。\n〒116-0003\n田村幸仁\n東京都荒川区南千住２丁目１９－１0 松本ビル２０１\nTEL： FAX：\n070-8324-0785\nE-Mail： poppy-tamura@pop21.odn.ne.jp 担当： 田村 幸仁\nご請求合計金額 ¥222, 672 （税込） お支払期限：\n令和8年4月30日\n\nNo.\n摘要\n数量\n費用明細\n単価\n金額\n1\n令和８年3月『負けないデイトレスクール』DMS、及び DMCサポート業務代として\n1\nスクールサポート業 務一式\n223,141\n¥223,141\n2\n\n小計\n\n¥223,141\n\n（消費税10％対象： 223,141円)\n消費税額（１０％対象)\n\n¥22,314\n\n（消費税８％対象：        0円)\nご請求金額\n\n¥245, 455\n\n源泉所得税・復興特別所得税\n\n¥22,783',
    expect: { total: 245455, subtotal: 223141, tax: 22314 },
  },
  {
    name: 'Row13-P5(Six): 消費税ラベル直下が「税込価格」(値なし)で tax を 0 に',
    text: '株式会社SKY\n下記の通りご請求申し上げます。\n発行日：2026年 4月 10日\n御請求書\n御中 株式会社SIX\n〒 852-8036\n長崎県長崎市青山町10−18\n適格事業者登録番号：T9290001096059\n\nご請求金額\n\\936,643\n\nNO\n商品名\n数量\n単価\n金額\n1\nBコミさんプロジェクト報酬(2026年4月清算分)\n1\n\\786,643\n\\786,643\n2\nマネーアカデミーサポート代行費用(2026年3月分)\n1\n\\100,000\n\\100,000\n\n小 計\n\\936,643\n\n消費税\n税込価格\n お振り込み先\n 三菱UFJ銀行 長崎支店\n\n合 計\n\\936,643',
    expect: { vendor: '株式会社SIX', total: 936643, tax: 0 },
  },
  {
    name: 'Row14-P5(本番): 株式会社スカイ 御中 が長い行内 + 直後に corp+〒 が別行',
    text: 'ご請求書\n東京都渋谷区桜丘町1-2 渋谷サクラステージ セントラルビル9階 発行日 2026年4月1日 株式会社スカイ 御中 登録番号\n株式会社アンドキュー\n〒252-0311 神奈川県相模原市南区東林間1-24-7\n担当・田尻賢 ０９０−６０４７−７７８８\n毎々格別のお取立を賜り厚く御礼申し上げます。\n下記の通りご請求申し上げます。\n広報・ブランディング業務 業務委託費として お支払い期日 2026年4月15日\n合計金額 ¥144,000 （税込）\n品 名 数量 金 額\n1',
    expect: { vendor: '株式会社スカイ', total: 144000 },
  },
  {
    name: 'Row15-P5(本番): 「2 0 2 6年4月5日」のように数字をスペースで分割',
    text: '請求書\n株式会社 スカイ 御中日付：2 0 2 6年4月5日 請求書番号：I N V - 0 0 0 0 0 0 0 0 3 0\n所在地：\n小西 歩\n下記の通りご請求申し上げます。 合計金額 110,000円 (税込)\n振込期限：2026年4月25日\n振込先：楽天銀行(0036)エレキ支店(230) 普通 1774810\n名義：小西 歩\n詳細 数量 単価 金額 2 0 2 6年3月分報酬 1 1 0 0 , 0 0 0 1 0 0 , 0 0 0\n小計\n消費税 合計金額\n100,000 10,000 110,000',
    expect: { vendor: '株式会社スカイ', issueDate: '2026/4/5', docNumber: 'INV-0000000030', total: 110000 },
  },
  {
    name: 'Row11-P5(福原): 「請求番号: 1」(1桁) は docNumber に採用しない',
    text: '請 求 書\n2026年04月01日 請求番号: 1\n株式会社スカイ 様 福原 圭子 件名 : 2026年2・3月分\n下記のとおりご請求申し上げます。\nご請求金額 ¥ 7,000 -\n品番・品名 数量 単価 金額\nビジネス開脳アカデミーバックオフィス業務 2月分 1 2,000 2,000\nビジネス開脳アカデミーバックオフィス業務 3月分 1 5,000 5,000\n小計 7,000\n消費税(10% 内税) (636)\n合計 7,000',
    expect: { vendor: '株式会社スカイ', total: 7000, docNumber: '' },
  },
  // ===== Phase 6 ケース =====
  {
    name: 'Row8-P6(悠楽 本番): 同行に「消費税 X 合計 Y 内訳 Z」/ 値入れ替わり防止',
    text: '御請求書\n発行日： 2026/4/15\n請求書No. 202604-027\n株式会社スカイ 御中\n合計 ¥550,000) 税込み\n登録番号： T8011101091283\n株式会社悠楽\n東京都新宿区西新宿7-7-25\nワコーレ新宿第二ビル３階 下記の通りご請求申し上げます。 Tel. : 03-\n但し、コンサルティング代として。\nNo. 項目 数量 単位 単価 金額 備考 1 経営コンサルティング（2026年5月分） 1 件 500,000 500,000\n小計 500,000\n上記明細をご確認後、下記口座に\n4月25日までに振込みをお願い致します。\n恐れ入りますが、振り込み手数料は貴社にてご負担下さい。\n消費税 50,000 合計 550,000 内訳 10%対象額 500,000\n消費税\n50,000\n代金振込先 軽減8%対象額 0 銀行名 楽天銀行 0\n消費税\n支店名 第二営業支店（252）\n口座種別 普通\n口座番号 7985997',
    expect: { vendor: '株式会社スカイ', total: 550000, subtotal: 500000, tax: 50000, docNumber: '202604-027' },
  },
  {
    name: 'Row3-P6(kii 本番): 小計ラベル無し / 「税(10%) ¥450 合計 ¥4950」で subtotal=4500 を逆算',
    text: '請求書\n株式会社スカイ\n〒150-0031 東京都渋谷区桜丘町1番2号\n渋谷サクラステージセントラルビル9階\n下記の通りご請求申し上げます。 合計金額 ￥4,950-\nお支払い期限： 2026年4月末日\n請求日：2026年4月8日\n木山 彩花\nTEL:080-6921-0494\n〒475-0918\n愛知県半田市雁宿町3-203-4 パークサイド雁宿1-702\n品名 数量 単価 金額\n2026/3/1 ショート動画作成「ライオン父さん」No.41-45 30 ¥50 ¥1500 2026/3/3 ショート動画作成「黒字社長」No.81-90 60\n税(10%) ¥450\n合計 ¥4950',
    expect: { total: 4950, subtotal: 4500, tax: 450 },
  },
];

function testParser() {
  if (!TEST_CASES || TEST_CASES.length === 0) {
    console.log('TEST_CASES が空です。main.gs の TEST_CASES に rawText と期待値を追加してください');
    return;
  }

  // Gemini を一時的に抑止(純粋な正規表現精度を測る)
  var savedEnabled = CFG.gemini.enabled;
  CFG.gemini.enabled = false;

  var pass = 0, fail = 0;
  try {
    TEST_CASES.forEach(function(c) {
      var result = parseLedgerEntry(c.text, c.fileName || 'test.pdf', 'http://test');
      Object.keys(c.expect).forEach(function(k) {
        var actual = result[k];
        var expected = c.expect[k];
        var ok = String(actual) === String(expected);
        console.log((ok ? '✓' : '✗') + ' [' + c.name + '] ' + k +
                    ': "' + actual + '" (期待: "' + expected + '")');
        if (ok) pass++; else fail++;
      });
    });
  } finally {
    CFG.gemini.enabled = savedEnabled;
  }

  console.log('===== 結果: ' + pass + ' pass / ' + fail + ' fail =====');
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

  // 2. 取引台帳スプレッドシートを確保
  var ledgerFile = findOrCreateLedger_(parentFolder);

  // 3. ScriptProperties に保存
  var props = PropertiesService.getScriptProperties();
  props.setProperties({
    [PROP_KEYS.upload]: uploadFolder.getId(),
    [PROP_KEYS.processed]: processedFolder.getId(),
    [PROP_KEYS.ledger]: ledgerFile.getId(),
  });

  // 4. トリガー登録（既存があれば再作成）
  setupTrigger();

  console.log('');
  console.log('===== セットアップ完了 =====');
  console.log('親フォルダ : ' + parentFolder.getUrl());
  console.log('UPLOAD    : ' + uploadFolder.getUrl());
  console.log('処理済み   : ' + processedFolder.getUrl());
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
