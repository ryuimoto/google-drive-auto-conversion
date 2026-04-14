/**
 * ユーティリティ関数
 */

/**
 * 指定フォルダから未処理のサポート対象ファイル一覧を取得
 * @param {string} folderId - 対象フォルダID
 * @return {GoogleAppsScript.Drive.File[]} 未処理ファイルの配列
 */
function getUnprocessedFiles(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const result = [];

  while (files.hasNext()) {
    const file = files.next();
    if (isProcessed(file)) continue;
    if (!isSupportedMimeType(file.getMimeType())) continue;
    result.push(file);
  }

  return result;
}

/**
 * MIMEタイプがサポート対象か判定
 * @param {string} mimeType
 * @return {boolean}
 */
function isSupportedMimeType(mimeType) {
  return getConversionRoute(mimeType) !== 'skip';
}

/**
 * MIMEタイプから変換経路を判定
 * @param {string} mimeType
 * @return {string} 'ocr' | 'toSheet' | 'toDoc' | 'skip'
 */
function getConversionRoute(mimeType) {
  if (CFG.supportedMimeTypes.ocr.indexOf(mimeType) !== -1) return 'ocr';
  if (CFG.supportedMimeTypes.toSheet.indexOf(mimeType) !== -1) return 'toSheet';
  if (CFG.supportedMimeTypes.toDoc.indexOf(mimeType) !== -1) return 'toDoc';
  return 'skip';
}

/**
 * ファイルが処理済みかどうか判定
 * @param {GoogleAppsScript.Drive.File} file
 * @return {boolean}
 */
function isProcessed(file) {
  const description = file.getDescription() || '';
  return description.indexOf('[PROCESSED]') !== -1;
}

/**
 * ファイルを処理済みフォルダに整頓
 * - [PROCESSED] マークを付与
 * - 書類種別 × 年月のサブフォルダに振り分け
 * - {日付}_{取引先}_{書類番号} 形式にリネーム
 * @param {string} fileId
 * @param {Object} entry - parseLedgerEntry()の戻り値
 */
function organizeProcessedFile(fileId, entry) {
  var file = DriveApp.getFileById(fileId);

  var currentDesc = file.getDescription() || '';
  file.setDescription('[PROCESSED] ' + currentDesc);

  if (!CFG.folders.processed) return;

  // 振り分け先の年月を決定（発行日 → なければ処理日時）
  var dateForFolder = parseDateForFolder(entry.issueDate) || entry.processedAt;
  var yearMonth = Utilities.formatDate(dateForFolder, 'Asia/Tokyo', 'yyyy-MM');

  // 振り分け先フォルダ: 処理済み/書類種別/年月/
  var processedRoot = DriveApp.getFolderById(CFG.folders.processed);
  var typeFolder = getOrCreateSubfolder(processedRoot, entry.docType || 'その他');
  var monthFolder = getOrCreateSubfolder(typeFolder, yearMonth);

  // ファイル名のリネーム
  var newName = buildOrganizedFileName(file.getName(), entry, dateForFolder);
  if (newName !== file.getName()) {
    var uniqueName = ensureUniqueName(monthFolder, newName);
    file.setName(uniqueName);
  }

  file.moveTo(monthFolder);
  console.log('整頓: ' + (entry.docType || 'その他') + '/' + yearMonth + '/' + file.getName());
}

/**
 * サブフォルダを取得（なければ作成）
 * @param {GoogleAppsScript.Drive.Folder} parent
 * @param {string} name
 * @return {GoogleAppsScript.Drive.Folder}
 */
function getOrCreateSubfolder(parent, name) {
  var iter = parent.getFoldersByName(name);
  if (iter.hasNext()) return iter.next();
  return parent.createFolder(name);
}

/**
 * 整頓後のファイル名を組み立て
 * 形式: {yyyy-MM-dd}_{取引先}_{書類番号}.{元拡張子}
 * 抽出失敗項目は省略
 * @param {string} originalName
 * @param {Object} entry
 * @param {Date} dateObj
 * @return {string}
 */
function buildOrganizedFileName(originalName, entry, dateObj) {
  var ext = '';
  var extMatch = originalName.match(/\.([^.]+)$/);
  if (extMatch) ext = '.' + extMatch[1];

  var parts = [];
  if (dateObj) {
    parts.push(Utilities.formatDate(dateObj, 'Asia/Tokyo', 'yyyy-MM-dd'));
  }
  if (entry.vendor) parts.push(sanitizeForFileName(entry.vendor));
  if (entry.docNumber) parts.push(sanitizeForFileName(entry.docNumber));

  if (parts.length === 0) return originalName;
  return parts.join('_') + ext;
}

/**
 * ファイル名に使えない文字を置換
 * @param {string} str
 * @return {string}
 */
function sanitizeForFileName(str) {
  return String(str).replace(/[\/\\:*?"<>|]/g, '_').trim();
}

/**
 * 同名衝突時に (2), (3) ... を付与
 * @param {GoogleAppsScript.Drive.Folder} folder
 * @param {string} name
 * @return {string}
 */
function ensureUniqueName(folder, name) {
  if (!folder.getFilesByName(name).hasNext()) return name;

  var ext = '';
  var base = name;
  var extMatch = name.match(/^(.+?)(\.[^.]+)$/);
  if (extMatch) {
    base = extMatch[1];
    ext = extMatch[2];
  }

  for (var i = 2; i < 100; i++) {
    var candidate = base + '(' + i + ')' + ext;
    if (!folder.getFilesByName(candidate).hasNext()) return candidate;
  }
  return name;
}

/**
 * 文字列日付（'2026/04/13' 等）を Date オブジェクトに変換
 * @param {string} dateStr
 * @return {Date|null}
 */
function parseDateForFolder(dateStr) {
  if (!dateStr) return null;
  var match = String(dateStr).match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (!match) return null;
  return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
}

/**
 * 抽出テキストが使える状態か判定
 * デジタルPDFの埋め込みテキスト抽出が成功したかを判定するために使用
 * @param {string} text
 * @return {boolean}
 */
function hasUsableText(text) {
  if (!text) return false;
  var visible = text.replace(/\s/g, '');
  return visible.length >= 50;
}

/**
 * 全角英数字・記号を半角に変換
 * @param {string} text - 変換対象テキスト
 * @return {string} 半角変換後のテキスト
 */
function normalizeFullWidth(text) {
  return text.replace(/[\uff01-\uff5e]/g, function(ch) {
    return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0);
  });
}

/**
 * OCR特有のノイズを正規化する共通レイヤー
 * - 全角→半角
 * - 全角スラッシュ・縦棒を改行扱いに（"小計 ¥X ／ 消費税 ¥Y" を行分解するため）
 * - 連続空白を1個に縮約
 * 装飾スペース("I n v o i c e")は正規表現側で `\s*` を挟んで吸収する方針とし、
 * 本関数では破壊しない（"is a test" のような単語連結を誤って詰めるのを避けるため）
 * @param {string} text
 * @return {string}
 */
function normalizeOcrText(text) {
  if (!text) return '';
  var t = normalizeFullWidth(text);
  t = t.replace(/[／｜]/g, '\n');
  t = t.replace(/[ \u3000]{2,}/g, ' ');
  return t;
}

/**
 * 指定ラベルを「文字間に任意の空白が入っていても」マッチできる正規表現を構築
 * 例: "Invoice" → /I\s*n\s*v\s*o\s*i\s*c\s*e/
 *     "請求書"   → /請\s*求\s*書/
 * これにより OCRが "I n v o i c e" / "請 求 書" のように装飾スペースを挟んでも検出できる
 * @param {string} label
 * @param {string} [flags] - 正規表現フラグ (例: 'i')
 * @return {RegExp}
 */
/**
 * 装飾スペース(1文字ずつ空白で区切られた区間)を検出して圧縮
 * 例: "I S S U E  N o .  N X - 2 0 2 6 - 0 8 2 2" → "ISSUE No. NX-2026-0822"
 *     "株 式 会 社 ネ ク ス ト イ ノ ベ ー シ ョ ン" → "株式会社ネクストイノベーション"
 *
 * 「1文字+空白」が4回以上連続する区間のみ圧縮するため、
 * "is a test"(3連続) や "株 式 会 社" 単独(3連続) など通常の文章は壊さない
 *
 * @param {string} text
 * @return {string}
 */
function decompactLetterSpacing(text) {
  if (!text) return '';
  return text.replace(/((?:[^\s\n]\s){4,}[^\s\n])/g, function(seg) {
    return seg.replace(/[ \u3000]/g, '');
  });
}

function buildSpacedLabelRegex(label, flags) {
  // ラベル内の空白は落とす("Amount Due" も "AmountDue" もマッチさせるため)
  var chars = label.split('').filter(function(ch) {
    return !/\s/.test(ch);
  }).map(function(ch) {
    return ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  });
  return new RegExp(chars.join('\\s*'), flags || '');
}

/**
 * 数値文字列をパース（カンマ除去、全角→半角変換）
 * @param {string} str - 数値文字列
 * @return {number} パース結果（NaNの場合は0）
 */
function parseNumber(str) {
  if (!str) return 0;
  const normalized = normalizeFullWidth(str).replace(/[,、]/g, '').replace(/[¥￥$]/g, '').trim();
  const num = parseFloat(normalized);
  return isNaN(num) ? 0 : num;
}

/**
 * 処理結果をログ出力
 * @param {string} fileName - ファイル名
 * @param {string} status - 処理ステータス（success/error）
 * @param {string} outputId - 出力ファイルID
 * @param {string} [error] - エラーメッセージ
 */
function logResult(fileName, status, outputId, error) {
  const timestamp = Utilities.formatDate(new Date(), CFG.ocr.language === 'ja' ? 'Asia/Tokyo' : 'UTC', 'yyyy-MM-dd HH:mm:ss');
  const message = '[' + timestamp + '] ' + status.toUpperCase() + ': ' + fileName;

  if (status === 'success') {
    console.log(message + ' → 出力ID: ' + outputId);
  } else {
    console.error(message + ' → エラー: ' + (error || '不明'));
  }
}

/**
 * 安全に関数を実行（エラーをキャッチしてログ＆通知）
 * @param {Function} fn - 実行する関数
 * @param {string} context - エラーログ用のコンテキスト名
 * @param {string} [fileName] - 通知用ファイル名（指定時はエラーメール送信）
 * @return {*} 関数の戻り値（エラー時はnull）
 */
function safeExecute(fn, context, fileName) {
  try {
    return fn();
  } catch (e) {
    console.error('[' + context + '] エラー: ' + e.message);
    console.error(e.stack);

    if (fileName && CFG.notification.enabled && CFG.notification.notifyOnError) {
      try {
        notifyError(fileName, e.message + '\n\n' + (e.stack || ''));
      } catch (notifyErr) {
        console.error('通知送信失敗: ' + notifyErr.message);
      }
    }

    return null;
  }
}
