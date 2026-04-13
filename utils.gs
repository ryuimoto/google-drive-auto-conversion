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
 * ファイルを処理済みとしてマーク＆処理済みフォルダへ移動
 * @param {string} fileId - 対象ファイルID
 */
function markAsProcessed(fileId) {
  const file = DriveApp.getFileById(fileId);

  const currentDesc = file.getDescription() || '';
  file.setDescription('[PROCESSED] ' + currentDesc);

  if (CFG.folders.processed) {
    const processedFolder = DriveApp.getFolderById(CFG.folders.processed);
    file.moveTo(processedFolder);
  }
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
