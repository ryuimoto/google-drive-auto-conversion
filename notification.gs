/**
 * 通知モジュール
 * 変換完了・失敗時にユーザーへGmailで通知する
 */

/**
 * 通知メールの宛先を決定
 * CFG.notification.recipientEmail が空ならアクティブユーザーのメールアドレスを返す
 * @return {string}
 */
function getNotificationRecipient() {
  if (CFG.notification.recipientEmail) {
    return CFG.notification.recipientEmail;
  }
  return Session.getActiveUser().getEmail();
}

/**
 * 変換完了メールを送信
 * @param {string} fileName - 元ファイル名
 * @param {string} outputFileId - 出力ファイルID
 * @param {string} outputType - 'ledger' | 'sheet' | 'doc'
 */
function notifySuccess(fileName, outputFileId, outputType) {
  if (!CFG.notification.enabled || !CFG.notification.notifyOnSuccess) return;

  const recipient = getNotificationRecipient();
  if (!recipient) {
    console.warn('通知宛先が取得できません。通知をスキップします');
    return;
  }

  var typeLabel, url, headline;
  if (outputType === 'ledger') {
    typeLabel = '取引台帳';
    url = 'https://docs.google.com/spreadsheets/d/' + outputFileId + '/edit';
    headline = 'ファイルを取引台帳に追記しました。';
  } else if (outputType === 'sheet') {
    typeLabel = 'スプレッドシート';
    url = 'https://drive.google.com/file/d/' + outputFileId + '/view';
    headline = 'ファイルの変換が完了しました。';
  } else {
    typeLabel = 'ドキュメント';
    url = 'https://drive.google.com/file/d/' + outputFileId + '/view';
    headline = 'ファイルの変換が完了しました。';
  }

  const subject = CFG.notification.subjectPrefix + '変換完了: ' + fileName;

  const htmlBody =
    '<p><b>' + headline + '</b></p>' +
    '<table style="border-collapse:collapse">' +
    '<tr><td style="padding:4px 12px 4px 0"><b>元ファイル</b></td><td>' + escapeHtml(fileName) + '</td></tr>' +
    '<tr><td style="padding:4px 12px 4px 0"><b>出力先</b></td><td>' + typeLabel + '</td></tr>' +
    '<tr><td style="padding:4px 12px 4px 0"><b>リンク</b></td><td><a href="' + url + '">' + url + '</a></td></tr>' +
    '</table>';

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: htmlBody,
  });

  console.log('完了通知を送信: ' + recipient + ' (' + fileName + ')');
}

/**
 * 変換失敗メールを送信
 * @param {string} fileName - 元ファイル名
 * @param {string} errorMessage - エラーメッセージ（スタックトレース含めてもよい）
 */
function notifyError(fileName, errorMessage) {
  if (!CFG.notification.enabled || !CFG.notification.notifyOnError) return;

  const recipient = getNotificationRecipient();
  if (!recipient) {
    console.warn('通知宛先が取得できません。通知をスキップします');
    return;
  }

  const subject = CFG.notification.subjectPrefix + '変換失敗: ' + fileName;
  const htmlBody =
    '<p><b>ファイルの変換に失敗しました。</b></p>' +
    '<p><b>元ファイル:</b> ' + escapeHtml(fileName) + '</p>' +
    '<p><b>エラー内容:</b></p>' +
    '<pre style="background:#f5f5f5;padding:8px;border:1px solid #ddd;white-space:pre-wrap">' +
    escapeHtml(errorMessage) +
    '</pre>';

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: htmlBody,
  });

  console.log('失敗通知を送信: ' + recipient + ' (' + fileName + ')');
}

/**
 * HTMLエスケープ
 * @param {string} str
 * @return {string}
 */
function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
