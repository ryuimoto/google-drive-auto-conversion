/**
 * PDF → Google スプレッドシート / ドキュメント 変換ツール
 * 設定ファイル
 *
 * フォルダID・台帳IDは setup() を実行するとScriptPropertiesに自動保存され、
 * このファイルを編集する必要はありません。
 */

// ScriptProperties に保存するキー（setup() が書き込み、CFG が読み込む）
const PROP_KEYS = {
  upload: 'FOLDER_UPLOAD_ID',
  processed: 'FOLDER_PROCESSED_ID',
  output: 'FOLDER_OUTPUT_ID',
  ledger: 'LEDGER_SPREADSHEET_ID',
};

function _getProp(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || '';
}

const CFG = {
  // フォルダIDは setup() 実行時に ScriptProperties に保存される
  folders: {
    get upload() { return _getProp(PROP_KEYS.upload); },
    get processed() { return _getProp(PROP_KEYS.processed); },
    get output() { return _getProp(PROP_KEYS.output); },
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

  // 取引台帳設定（spreadsheetId は setup() 実行時に ScriptProperties に保存される）
  ledger: {
    get spreadsheetId() { return _getProp(PROP_KEYS.ledger); },
    sheetName: '取引台帳',     // 台帳のタブ名
  },

  // 通知設定
  notification: {
    enabled: true,
    recipientEmail: '',       // 空の場合は Session.getActiveUser().getEmail() を使用
    notifyOnSuccess: true,
    notifyOnError: true,
    subjectPrefix: '[Googleドライブ自動変換] ',
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
