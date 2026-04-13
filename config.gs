/**
 * PDF → Google スプレッドシート / ドキュメント 変換ツール
 * 設定ファイル
 */

const CFG = {
  // フォルダID設定（Google DriveのフォルダIDを設定してください）
  folders: {
    upload: '',          // 統合アップロードフォルダID
    processed: '',       // 処理済み移動先フォルダID
    output: '',          // 変換結果の出力先フォルダID
  },

  // OCR設定
  ocr: {
    language: 'ja',
  },

  // トリガー設定
  trigger: {
    intervalMinutes: 5,
    functionName: 'scanAndProcessFiles',
  },

  // 処理設定
  processing: {
    maxFilesPerExecution: 5,  // 1回の実行で処理する最大ファイル数（6分制限対策）
  },

  // 通知設定
  notification: {
    enabled: true,
    recipientEmail: '',       // 空の場合は Session.getActiveUser().getEmail() を使用
    notifyOnSuccess: true,
    notifyOnError: true,
    subjectPrefix: '[PDF自動変換] ',
  },

  // 対応MIMEタイプと振り分け先
  supportedMimeTypes: {
    // OCR経路: 画像/PDF → OCRでテキスト化 → 内容判定で Sheet/Doc に振り分け
    ocr: [
      'application/pdf',
      'image/jpeg',
      'image/png',
      'image/gif',
      'image/bmp',
      'image/tiff',
    ],
    // Sheet直接変換経路: 表形式ファイル
    toSheet: [
      'application/vnd.ms-excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'text/csv',
    ],
    // Doc直接変換経路: 文書/プレゼン/テキスト
    toDoc: [
      'application/msword',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'application/vnd.ms-powerpoint',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'text/plain',
      'text/rtf',
      'application/rtf',
    ],
  },
};
