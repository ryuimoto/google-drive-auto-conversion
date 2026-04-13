/**
 * 請求書パーサー
 * OCRテキストから請求書データを構造化して抽出する
 */

/**
 * OCRテキストから請求書データをパース
 * @param {string} text - OCR抽出テキスト
 * @return {Object} パースされた請求書データ
 */
function parseInvoice(text) {
  // 全角→半角変換済みテキストを用意
  const normalized = normalizeFullWidth(text);

  return {
    invoiceNumber: extractInvoiceNumber(normalized, text),
    issueDate: extractDate(normalized, text),
    vendorName: extractVendorName(text),
    recipientName: extractRecipientName(text),
    items: extractItems(normalized, text),
    subtotal: extractAmount(normalized, '小計'),
    taxAmount: extractAmount(normalized, '消費税'),
    total: extractAmount(normalized, '合計'),
    rawText: text,
  };
}

/**
 * 請求番号を抽出
 * @param {string} normalized - 半角変換済みテキスト
 * @param {string} original - 元テキスト
 * @return {string}
 */
function extractInvoiceNumber(normalized, original) {
  var patterns = [
    /請求番号[：:\s]*([A-Za-z0-9\-_]+)/,
    /請求書番号[：:\s]*([A-Za-z0-9\-_]+)/,
    /No\.\s*([A-Za-z0-9\-_]+)/,
    /Invoice\s*#?\s*([A-Za-z0-9\-_]+)/i,
    /伝票番号[：:\s]*([A-Za-z0-9\-_]+)/,
  ];

  for (var i = 0; i < patterns.length; i++) {
    var match = normalized.match(patterns[i]) || original.match(patterns[i]);
    if (match) return match[1].trim();
  }

  return '';
}

/**
 * 日付を抽出
 * @param {string} normalized - 半角変換済みテキスト
 * @param {string} original - 元テキスト
 * @return {string}
 */
function extractDate(normalized, original) {
  var patterns = [
    // 和暦: 令和6年1月15日
    /令和\s*(\d+)\s*年\s*(\d+)\s*月\s*(\d+)\s*日/,
    // 西暦: 2024年1月15日
    /(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/,
    // スラッシュ・ハイフン区切り: 2024/01/15, 2024-01-15
    /(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/,
  ];

  // 和暦パターン
  var reiwaMatch = normalized.match(patterns[0]) || original.match(patterns[0]);
  if (reiwaMatch) {
    var year = 2018 + parseInt(reiwaMatch[1]);
    return year + '/' + reiwaMatch[2] + '/' + reiwaMatch[3];
  }

  // 西暦パターン
  var dateMatch = normalized.match(patterns[1]) || original.match(patterns[1]);
  if (dateMatch) {
    return dateMatch[1] + '/' + dateMatch[2] + '/' + dateMatch[3];
  }

  // スラッシュ・ハイフン区切り
  var slashMatch = normalized.match(patterns[2]) || original.match(patterns[2]);
  if (slashMatch) {
    return slashMatch[1] + '/' + slashMatch[2] + '/' + slashMatch[3];
  }

  // 請求日・発行日の前後で日付を探す
  var contextPatterns = [
    /(?:請求日|発行日|日付)[：:\s]*(.+)/,
  ];
  for (var i = 0; i < contextPatterns.length; i++) {
    var contextMatch = normalized.match(contextPatterns[i]);
    if (contextMatch) {
      // 再帰的にマッチした文字列から日付を抽出
      var innerDate = contextMatch[1].trim();
      for (var j = 0; j < patterns.length; j++) {
        var innerMatch = innerDate.match(patterns[j]);
        if (innerMatch) {
          if (j === 0) {
            return (2018 + parseInt(innerMatch[1])) + '/' + innerMatch[2] + '/' + innerMatch[3];
          }
          return innerMatch[1] + '/' + innerMatch[2] + '/' + innerMatch[3];
        }
      }
    }
  }

  return '';
}

/**
 * 発行元（会社名）を抽出
 * @param {string} text - 元テキスト
 * @return {string}
 */
function extractVendorName(text) {
  var patterns = [
    /^(.+(?:株式会社|有限会社|合同会社|（株）|\(株\)))/m,
    /^((?:株式会社|有限会社|合同会社).+)$/m,
  ];

  for (var i = 0; i < patterns.length; i++) {
    var match = text.match(patterns[i]);
    if (match) return match[1].trim();
  }

  return '';
}

/**
 * 宛先を抽出
 * @param {string} text - 元テキスト
 * @return {string}
 */
function extractRecipientName(text) {
  var patterns = [
    /(.+?)\s*(?:御中|様|殿)/,
  ];

  for (var i = 0; i < patterns.length; i++) {
    var match = text.match(patterns[i]);
    if (match) return match[1].trim();
  }

  return '';
}

/**
 * 明細行を抽出
 * @param {string} normalized - 半角変換済みテキスト
 * @param {string} original - 元テキスト
 * @return {Object[]} 明細行の配列 [{name, quantity, unitPrice, amount}]
 */
function extractItems(normalized, original) {
  var items = [];
  var lines = normalized.split('\n');

  // 明細ヘッダー行を検出
  var headerIndex = -1;
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    if ((line.indexOf('品名') !== -1 || line.indexOf('品目') !== -1 || line.indexOf('内容') !== -1 || line.indexOf('摘要') !== -1) &&
        (line.indexOf('金額') !== -1 || line.indexOf('数量') !== -1)) {
      headerIndex = i;
      break;
    }
  }

  if (headerIndex === -1) return items;

  // ヘッダー以降の行をパース
  // 数量・単価・金額パターン: テキスト + 数値 + 数値 + 数値
  var itemPattern = /(.+?)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)/;
  // 品名 + 金額のみのパターン
  var simplePattern = /(.+?)\s+([\d,]+)$/;

  for (var j = headerIndex + 1; j < lines.length; j++) {
    var itemLine = lines[j].trim();
    if (!itemLine) continue;

    // 小計・合計行に到達したら終了
    if (/^(小計|合計|消費税|税込|税抜)/.test(itemLine)) break;

    var itemMatch = itemLine.match(itemPattern);
    if (itemMatch) {
      items.push({
        name: itemMatch[1].trim(),
        quantity: parseNumber(itemMatch[2]),
        unitPrice: parseNumber(itemMatch[3]),
        amount: parseNumber(itemMatch[4]),
      });
      continue;
    }

    // 品名 + 金額のみ
    var simpleMatch = itemLine.match(simplePattern);
    if (simpleMatch) {
      var amount = parseNumber(simpleMatch[2]);
      if (amount > 0) {
        items.push({
          name: simpleMatch[1].trim(),
          quantity: 1,
          unitPrice: amount,
          amount: amount,
        });
      }
    }
  }

  return items;
}

/**
 * 指定ラベルの金額を抽出
 * @param {string} text - テキスト
 * @param {string} label - ラベル（小計、消費税、合計）
 * @return {number}
 */
function extractAmount(text, label) {
  var patterns = [
    new RegExp(label + '[：:　\\s]*[¥￥]?\\s*([\\d,]+)'),
    new RegExp(label + '.*?([\\d,]+)\\s*円'),
    new RegExp(label + '金額[：:　\\s]*[¥￥]?\\s*([\\d,]+)'),
  ];

  for (var i = 0; i < patterns.length; i++) {
    var match = text.match(patterns[i]);
    if (match) return parseNumber(match[1]);
  }

  return 0;
}

/**
 * 汎用テーブルテキストを2次元配列に分割
 * タブまたは2つ以上の連続空白を区切り文字として行をパース
 * @param {string} text - OCR抽出テキスト
 * @return {string[][]} 行ごとのセル配列
 */
function parseGenericTable(text) {
  const lines = text.split('\n');
  const rows = [];

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].replace(/\s+$/, '');
    if (!line.trim()) continue;

    // タブ優先、なければ2連続以上の空白で分割
    var cells;
    if (line.indexOf('\t') !== -1) {
      cells = line.split('\t');
    } else {
      cells = line.split(/\s{2,}/);
    }

    cells = cells.map(function(c) { return c.trim(); }).filter(function(c) { return c.length > 0; });
    if (cells.length > 0) rows.push(cells);
  }

  // 列数を最大値に揃える
  var maxCols = 0;
  for (var j = 0; j < rows.length; j++) {
    if (rows[j].length > maxCols) maxCols = rows[j].length;
  }
  for (var k = 0; k < rows.length; k++) {
    while (rows[k].length < maxCols) rows[k].push('');
  }

  return rows;
}

/**
 * テキストの内容タイプを自動判定
 * @param {string} text - OCR抽出テキスト
 * @return {string} 'invoice' | 'table' | 'text'
 */
function detectContentType(text) {
  var invoiceKeywords = ['請求書', '請求番号', '御請求', 'Invoice', '請求金額'];
  for (var i = 0; i < invoiceKeywords.length; i++) {
    if (text.indexOf(invoiceKeywords[i]) !== -1) return 'invoice';
  }

  // テーブルっぽい構造を検出（タブ区切りや連続スペースの行が複数）
  var lines = text.split('\n');
  var tableLineCount = 0;
  for (var j = 0; j < lines.length; j++) {
    if (/\t/.test(lines[j]) || /\S\s{2,}\S/.test(lines[j])) {
      tableLineCount++;
    }
  }
  if (tableLineCount >= 3) return 'table';

  return 'text';
}
