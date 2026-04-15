/**
 * Gemini API 連携モジュール
 *
 * 正規表現パーサーで抽出精度が不十分な場合のフォールバック。
 * - 無料枠: gemini-2.0-flash は 1日1500リクエストまで無料
 * - API キーは ScriptProperties の GEMINI_API_KEY に手動設定
 * - 未設定なら静かにスキップ(正規表現のみで動作)
 */

// 抽出したいフィールドの JSON Schema (Gemini の Structured Output で型保証)
const GEMINI_INVOICE_SCHEMA = {
  type: 'object',
  properties: {
    docType: {
      type: 'string',
      enum: ['請求書', '注文書', '見積書', '領収書', '納品書', '契約書', '議事録', '報告書', 'その他'],
    },
    vendorName:        { type: 'string' },
    recipientName:     { type: 'string' },
    invoiceNumber:     { type: 'string' },
    issueDate:         { type: 'string', description: 'yyyy/MM/dd 形式' },
    paymentDueDate:    { type: 'string', description: 'yyyy/MM/dd 形式' },
    subtotal:          { type: 'number' },
    taxAmount:         { type: 'number' },
    total:             { type: 'number' },
    registrationNumber: { type: 'string', description: '適格請求書の登録番号 T+13桁' },
    items: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          name:      { type: 'string' },
          quantity:  { type: 'number' },
          unit:      { type: 'string' },
          unitPrice: { type: 'number' },
          amount:    { type: 'number' },
        },
      },
    },
  },
};

/**
 * 抽出済みOCRテキストから請求書データを Gemini で構造化抽出
 * エラー・未設定・枠超過時は null を返す(呼び出し側は正規表現結果にフォールバック)
 * @param {string} text - OCR抽出テキスト
 * @return {Object|null}
 */
function extractInvoiceWithGemini(text) {
  if (!CFG.gemini || !CFG.gemini.enabled) return null;

  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return null;

  if (!canCallGeminiToday_()) {
    console.warn('[Gemini] 本日の無料枠上限(' + CFG.gemini.maxCallsPerDay + ')に達したためスキップ');
    return null;
  }

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
            CFG.gemini.model + ':generateContent?key=' + encodeURIComponent(apiKey);

  var prompt =
    '次のOCRテキストは請求書/注文書/契約書/領収書等のビジネス書類から抽出されたものです。\n' +
    '文書内容を分析し、指定されたJSONスキーマに従って構造化データを返してください。\n\n' +
    '注意事項:\n' +
    '- vendorName は発行元(請求元・注文元)。recipientName は宛先(御中/様/Bill To等)。混同しないこと\n' +
    '- 金額は円単位の数値(カンマ・¥記号なし)。取得できない場合は 0\n' +
    '- 日付は yyyy/MM/dd 形式。取得できない場合は空文字\n' +
    '- items は明細の配列。品名・数量・単位・単価・金額を抽出\n' +
    '- 装飾スペース("I N V O I C E"等)は解除して解釈すること\n' +
    '- 郵便番号や電話番号を金額と混同しないこと\n\n' +
    'OCRテキスト:\n```\n' + text + '\n```';

  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: GEMINI_INVOICE_SCHEMA,
      temperature: 0,
    },
  };

  // 503/500/504 は一時的な過負荷なので指数バックオフでリトライ
  var maxAttempts = 3;
  var response;
  var code;
  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      response = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      });
    } catch (e) {
      console.error('[Gemini] ネットワークエラー: ' + e.message);
      return null;
    }

    code = response.getResponseCode();
    if (code === 200) break;

    var isTransient = (code === 503 || code === 500 || code === 504);
    if (isTransient && attempt < maxAttempts) {
      var waitMs = 1000 * Math.pow(2, attempt - 1); // 1秒 → 2秒 → 4秒
      console.warn('[Gemini] HTTP ' + code + ' (一時的) 試行 ' + attempt + '/' + maxAttempts +
                   ' → ' + waitMs + 'ms 待機後リトライ');
      Utilities.sleep(waitMs);
      continue;
    }
    break;
  }

  if (code !== 200) {
    console.error('[Gemini] HTTP ' + code + ': ' + response.getContentText().substring(0, 500));
    if (code === 404) {
      console.error('  → モデル名が間違っているか、このAPIキーで利用できません');
      console.error('  → listGeminiModels() を実行して利用可能なモデル一覧を確認してください');
    } else if (code === 429) {
      console.error('  → 無料枠を超過したか、このプロジェクトで無料枠が 0 です');
      console.error('  → https://aistudio.google.com/apikey で「Create API key in new project」を選び、');
      console.error('    新規プロジェクトでキーを再発行して GEMINI_API_KEY を上書きしてください');
    } else if (code === 503 || code === 500 || code === 504) {
      console.error('  → モデル側の一時的な過負荷です。数分後にもう一度試してください');
      console.error('  → 頻発する場合は CFG.gemini.model を別モデル(例: gemini-2.5-flash-lite)に変更');
    }
    return null;
  }

  var body;
  try {
    body = JSON.parse(response.getContentText());
  } catch (e) {
    console.error('[Gemini] レスポンスJSONパース失敗: ' + e.message);
    return null;
  }

  if (!body.candidates || !body.candidates[0] || !body.candidates[0].content) {
    console.error('[Gemini] 予期しないレスポンス構造');
    return null;
  }

  var parts = body.candidates[0].content.parts || [];
  var jsonText = parts[0] && parts[0].text ? parts[0].text : '';
  if (!jsonText) return null;

  var parsed;
  try {
    parsed = JSON.parse(jsonText);
  } catch (e) {
    console.error('[Gemini] 抽出JSONのパース失敗: ' + e.message + ' / text=' + jsonText.substring(0, 200));
    return null;
  }

  var usage = body.usageMetadata || {};
  var promptTokens = usage.promptTokenCount || 0;
  var outputTokens = usage.candidatesTokenCount || 0;
  var totalTokens = usage.totalTokenCount || (promptTokens + outputTokens);

  incrementGeminiCallCount_(totalTokens, promptTokens, outputTokens);
  console.log('[Gemini] 抽出成功: vendor=' + (parsed.vendorName || '-') +
              ' total=' + (parsed.total || 0) +
              ' items=' + ((parsed.items && parsed.items.length) || 0) +
              ' tokens=' + totalTokens +
              ' (prompt=' + promptTokens + ' / output=' + outputTokens + ')');
  return parsed;
}

/**
 * 当日の Gemini 呼び出し回数が無料枠内か確認
 * @return {boolean}
 */
function canCallGeminiToday_() {
  var props = PropertiesService.getScriptProperties();
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var storedDay = props.getProperty('GEMINI_COUNT_DAY');

  if (storedDay !== today) {
    props.setProperties({
      'GEMINI_COUNT_DAY': today,
      'GEMINI_COUNT': '0',
      'GEMINI_TOKENS_TOTAL': '0',
      'GEMINI_TOKENS_PROMPT': '0',
      'GEMINI_TOKENS_OUTPUT': '0',
    });
    return true;
  }

  var count = parseInt(props.getProperty('GEMINI_COUNT') || '0', 10);
  return count < CFG.gemini.maxCallsPerDay;
}

/**
 * Gemini 呼び出し回数とトークン消費量をインクリメント
 * @param {number} totalTokens
 * @param {number} promptTokens
 * @param {number} outputTokens
 */
function incrementGeminiCallCount_(totalTokens, promptTokens, outputTokens) {
  var props = PropertiesService.getScriptProperties();
  var count = parseInt(props.getProperty('GEMINI_COUNT') || '0', 10);
  var total = parseInt(props.getProperty('GEMINI_TOKENS_TOTAL') || '0', 10);
  var prompt = parseInt(props.getProperty('GEMINI_TOKENS_PROMPT') || '0', 10);
  var output = parseInt(props.getProperty('GEMINI_TOKENS_OUTPUT') || '0', 10);
  props.setProperties({
    'GEMINI_COUNT': String(count + 1),
    'GEMINI_TOKENS_TOTAL': String(total + (totalTokens || 0)),
    'GEMINI_TOKENS_PROMPT': String(prompt + (promptTokens || 0)),
    'GEMINI_TOKENS_OUTPUT': String(output + (outputTokens || 0)),
  });
}

/**
 * 当日の Gemini コール統計を取得(diagnose 用)
 * @return {{count: number, limit: number, hasKey: boolean}}
 */
function getGeminiStats() {
  var props = PropertiesService.getScriptProperties();
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var storedDay = props.getProperty('GEMINI_COUNT_DAY');
  var isToday = (storedDay === today);
  var count = isToday ? parseInt(props.getProperty('GEMINI_COUNT') || '0', 10) : 0;
  var totalTokens = isToday ? parseInt(props.getProperty('GEMINI_TOKENS_TOTAL') || '0', 10) : 0;
  var promptTokens = isToday ? parseInt(props.getProperty('GEMINI_TOKENS_PROMPT') || '0', 10) : 0;
  var outputTokens = isToday ? parseInt(props.getProperty('GEMINI_TOKENS_OUTPUT') || '0', 10) : 0;

  // gemini-2.0-flash 料金 (2026年4月時点): input $0.10 / 1M, output $0.40 / 1M
  // 無料枠内なら実際のコストは $0
  var estimatedCostUsd = (promptTokens * 0.10 + outputTokens * 0.40) / 1000000;

  return {
    count: count,
    limit: CFG.gemini.maxCallsPerDay,
    hasKey: !!props.getProperty('GEMINI_API_KEY'),
    enabled: CFG.gemini.enabled,
    totalTokens: totalTokens,
    promptTokens: promptTokens,
    outputTokens: outputTokens,
    estimatedCostUsd: estimatedCostUsd,
  };
}
