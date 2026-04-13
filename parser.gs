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

  // ヘッダー行が見つからない場合は別レイアウト用のパーサにフォールバック
  if (headerIndex === -1) {
    var cellPerLine = extractItemsCellPerLine(normalized);
    if (cellPerLine.length > 0) return cellPerLine;
    return extractItemsColumnMajor(normalized);
  }

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

  // 行レイアウトで取れなかった場合は他のレイアウトでフォールバック
  if (items.length === 0) {
    var cellPerLine = extractItemsCellPerLine(normalized);
    if (cellPerLine.length > 0) return cellPerLine;
    return extractItemsColumnMajor(normalized);
  }

  return items;
}

/**
 * 1セル1行レイアウトから明細を抽出
 * デジタルPDFで「品目→数量→単価→金額」が縦に並び、続いて各行の値も縦に並ぶケース
 * @param {string} normalized - 半角変換済みテキスト
 * @return {Object[]}
 */
function extractItemsCellPerLine(normalized) {
  // 全角スペース・タブ・改行などを含めて空白を除去するヘルパー
  function strip(s) {
    return (s || '').replace(/[\s\u3000\u00a0]+/g, '');
  }

  var lines = normalized.split('\n')
    .map(function(l) { return l.replace(/\s+$/, '').replace(/^\s+/, ''); })
    .filter(function(l) { return l.length > 0; });

  // 品目ヘッダーを含む行を探す（短い行のみ＝ヘッダー単独行を想定）
  var nameHeaderIdx = -1;
  for (var i = 0; i < lines.length; i++) {
    var clean = strip(lines[i]);
    if (clean.length > 8) continue;  // 長い行はヘッダーではない
    if (clean === '品名' || clean === '品目' || clean === '品番' ||
        clean === '品番・品名' || clean === '内容' || clean === '摘要') {
      nameHeaderIdx = i;
      break;
    }
  }
  if (nameHeaderIdx === -1) {
    console.log('[cellPerLine] 品名ヘッダーが見つかりません');
    return [];
  }

  // 品目ヘッダー以降の近傍5行以内に 数量/単価/金額 ヘッダーを探す
  var qtyIdx = -1, unitIdx = -1, amountIdx = -1;
  var searchEnd = Math.min(lines.length, nameHeaderIdx + 6);
  for (var j = nameHeaderIdx + 1; j < searchEnd; j++) {
    var c = strip(lines[j]);
    if (c.length > 6) continue;
    if (c === '数量' && qtyIdx === -1) qtyIdx = j;
    else if (c === '単価' && unitIdx === -1) unitIdx = j;
    else if (c === '金額' && amountIdx === -1) amountIdx = j;
  }

  console.log('[cellPerLine] ヘッダー位置: 品名=' + nameHeaderIdx +
              ' 数量=' + qtyIdx + ' 単価=' + unitIdx + ' 金額=' + amountIdx);

  if (qtyIdx === -1 || unitIdx === -1 || amountIdx === -1) {
    console.log('[cellPerLine] 列ヘッダーが揃わないため終了');
    return [];
  }

  // 値の開始位置 = 一番後ろのヘッダーの次の行
  var startIdx = Math.max(nameHeaderIdx, qtyIdx, unitIdx, amountIdx) + 1;

  // 4行ずつ1明細として読む
  var items = [];
  var idx = startIdx;
  while (idx + 3 < lines.length) {
    var name = lines[idx];
    var qty = lines[idx + 1];
    var unit = lines[idx + 2];
    var amount = lines[idx + 3];

    // 合計部に到達したら終了
    if (/^(小計|合計|消費税|税込|税抜|総額|請求|お振込|お支払)/.test(strip(name))) break;

    // ハイフンだけの空行はスキップ
    if (name === '-' || name === '−' || name === '–' || name === 'ー') {
      idx += 4;
      continue;
    }

    var amountNum = parseNumber(amount);
    var qtyNum = parseNumber(qty);

    if (name && (amountNum > 0 || qtyNum > 0)) {
      items.push({
        name: name,
        quantity: qtyNum,
        unitPrice: parseNumber(unit),
        amount: amountNum,
      });
    }

    idx += 4;
  }

  console.log('[cellPerLine] 抽出件数: ' + items.length);
  return items;
}

/**
 * 列レイアウトのOCR結果から明細を抽出
 * （OCRが表を列方向に読み取った場合に対応）
 * @param {string} normalized - 半角変換済みテキスト
 * @return {Object[]}
 */
function extractItemsColumnMajor(normalized) {
  var lines = normalized.split('\n')
    .map(function(l) { return l.trim(); })
    .filter(function(l) { return l.length > 0; });

  // 1. 品名ヘッダーを探す（金額・数量を含まない 品名/品目/品番 行）
  var nameHeaderIdx = -1;
  for (var i = 0; i < lines.length; i++) {
    if (/品名|品目|品番/.test(lines[i]) && !/数量|単価|金額/.test(lines[i])) {
      nameHeaderIdx = i;
      break;
    }
  }
  if (nameHeaderIdx === -1) return [];

  // 2. 金額ヘッダーを探す（列ヘッダーの末尾を想定）
  var amountHeaderIdx = -1;
  for (var j = nameHeaderIdx + 1; j < lines.length; j++) {
    if (/^金額$/.test(lines[j])) {
      amountHeaderIdx = j;
      break;
    }
  }
  if (amountHeaderIdx === -1) return [];

  // 3. 品名ヘッダーと金額ヘッダーの間にある行から品名を収集
  //    （他の列ヘッダー・短いノイズ行はスキップ）
  var names = [];
  for (var k = nameHeaderIdx + 1; k < amountHeaderIdx; k++) {
    var line = lines[k];
    if (line.length <= 2) continue;                                  // ノイズ
    if (/^(数量|単価|金額|備考|摘要|内容)$/.test(line)) continue;       // 他列ヘッダー
    names.push(line);
  }
  if (names.length === 0) return [];

  // 4. 金額ヘッダー以降から値を収集（小計・合計に達したら終了）
  var values = [];
  for (var m = amountHeaderIdx + 1; m < lines.length; m++) {
    var l = lines[m];
    if (/^(小計|小叶|合計|合叶|消費税|税込|税抜|総額|請求)/.test(l)) break;
    if (l.length <= 1) continue;
    values.push(l);
  }

  // 5. 値を3つずつ（数量・単価・金額）グループ化して品名と対応付け
  var items = [];
  for (var n = 0; n < names.length; n++) {
    var idx = n * 3;
    if (idx + 2 >= values.length) break;
    items.push({
      name: names[n],
      quantity: parseNumber(values[idx]),
      unitPrice: parseNumber(values[idx + 1]),
      amount: parseNumber(values[idx + 2]),
    });
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
  // ビジネス書類のキーワード判定（誤分類が起きにくい順）
  if (/見積書|御見積|見積金額|見積番号|Quotation/.test(text)) return 'quote';
  if (/注文書|発注書|注文番号|発注番号|Purchase\s*Order/.test(text)) return 'order';
  if (/納品書|納品日/.test(text)) return 'delivery';
  if (/領収書|領収証|但し書き/.test(text)) return 'receipt';
  if (/請求書|御請求|請求金額|請求番号|Invoice/.test(text)) return 'invoice';

  // テキスト系書類
  if (/議事録|出席者|議題|開催日時/.test(text)) return 'minutes';
  if (/契約書|契約期間|第\s*[一二三四五六七八九十0-9]+\s*条/.test(text)) return 'contract';
  if (/報告書|業務報告|月次報告/.test(text)) return 'report';

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
