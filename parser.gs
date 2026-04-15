/**
 * 請求書パーサー
 * OCRテキストから請求書データを構造化して抽出する
 */

/**
 * OCRテキストから請求書データをパース
 * @param {string} text - OCR抽出テキスト
 * @return {Object} パースされた請求書データ
 */
/**
 * ファイルから抽出したテキストを取引台帳の1エントリに変換
 * @param {string} text - 抽出済みテキスト
 * @param {string} fileName - 元ファイル名
 * @param {string} fileLink - 元ファイルへのリンク
 * @return {Object} 台帳エントリ
 */
function parseLedgerEntry(text, fileName, fileLink) {
  var invoice = parseInvoice(text);
  var contentType = detectContentType(text);

  var docTypeLabels = {
    invoice: '請求書',
    receipt: '領収書',
    quote: '見積書',
    order: '注文書',
    delivery: '納品書',
    contract: '契約書',
    minutes: '議事録',
    report: '報告書',
    table: '表データ',
    text: 'その他',
  };

  // 正規表現パース結果が不十分なら Gemini API で救済
  if (shouldCallGemini_(invoice, contentType)) {
    console.log('[parseLedgerEntry] 正規表現の結果が不十分なためGeminiを呼び出します: ' + fileName);
    var geminiResult = extractInvoiceWithGemini(text);
    if (geminiResult) {
      invoice = mergeGeminiResult_(invoice, geminiResult);
      var mappedType = mapGeminiDocType_(geminiResult.docType);
      if (mappedType) contentType = mappedType;
    }
  }

  // 内容サマリー: 明細の品名を最大3つカンマ区切り
  // 品名が全て欠落している場合は件数のみ表示
  var contentSummary = '';
  if (invoice.items && invoice.items.length > 0) {
    var namedItems = invoice.items.filter(function(it) { return it.name && it.name.length > 0; });
    if (namedItems.length === 0) {
      contentSummary = '明細' + invoice.items.length + '件';
    } else {
      contentSummary = namedItems.slice(0, 3)
        .map(function(it) { return it.name; })
        .join(', ');
      if (namedItems.length > 3) contentSummary += ' 他';
    }
  }

  // ステータス判定: 取引先と合計金額の両方が取れていればOK
  var hasKeyInfo = !!(invoice.vendorName && invoice.total);
  var status = hasKeyInfo ? 'OK' : '部分抽出';

  return {
    processedAt: new Date(),
    fileName: fileName,
    fileLink: fileLink,
    docType: docTypeLabels[contentType] || 'その他',
    vendor: invoice.vendorName,
    issueDate: invoice.issueDate,
    docNumber: invoice.invoiceNumber,
    total: invoice.total,
    subtotal: invoice.subtotal,
    tax: invoice.taxAmount,
    paymentDueDate: invoice.paymentDueDate || '',
    contentSummary: contentSummary,
    rawText: text,
    status: status,
  };
}

function parseInvoice(text) {
  // 3戦略を試行し、品質スコアが最高の結果を採用
  //   戦略A: 正規化なし(既存互換)
  //   戦略B: normalizeOcrText適用(全角区切り→改行、スペース縮約)
  //   戦略C: 装飾スペース(1文字単位空白区切り)を圧縮したバージョン
  // 全改行空白化の "flat" 戦略は extractAmount が口座番号等を誤検出するため不採用
  var normalized = normalizeOcrText(text);
  var decompacted = decompactLetterSpacing(normalized);
  var variants = [
    { name: 'raw',         text: text },
    { name: 'normalized',  text: normalized },
    { name: 'decompacted', text: decompacted },
    { name: 'dedoubled',   text: dedoubleRepeatedChars(decompacted) },
  ];

  var best = null;
  var bestScore = -1;
  for (var v = 0; v < variants.length; v++) {
    var result = parseInvoiceOnce(variants[v].text, text);
    var score = scoreParseResult(result);
    console.log('[parseInvoice] 戦略=' + variants[v].name + ' score=' + score);
    if (score > bestScore) {
      bestScore = score;
      best = result;
    }
  }

  return best;
}

/**
 * 1パス分のパーサー実行
 * @param {string} normalized - 解析対象(正規化済み or 元テキスト)
 * @param {string} original - 抽出元テキスト(rawText用)
 */
function parseInvoiceOnce(normalized, original) {
  var multiRate = extractMultiTaxRateTotals(normalized);

  var result = {
    invoiceNumber: extractInvoiceNumber(normalized, original),
    issueDate: extractDate(normalized, original),
    vendorName: extractVendorName(normalized),
    recipientName: extractRecipientName(normalized),
    items: extractItems(normalized, original),
    subtotal: multiRate.subtotal || extractAmount(normalized, '小計'),
    taxAmount: multiRate.taxAmount || extractAmount(normalized, '消費税'),
    total: extractAmount(normalized, '合計'),
    paymentDueDate: extractPaymentDueDate(normalized),
    registrationNumber: extractRegistrationNumber(normalized),
    rawText: original,
  };

  // 年なし日付のフォールバック: MM/DD のみでも文書中の年で補完
  var fallbackYear = extractFirstYear(normalized) || new Date().getFullYear();
  if (!result.issueDate) {
    result.issueDate = findShortDateNearLabel(
      normalized, ['発行日', 'Issue', 'ISSUED', 'Date', 'DATE'], fallbackYear
    );
  }
  if (!result.paymentDueDate) {
    result.paymentDueDate = findShortDateNearLabel(
      normalized, ['支払期日', '支払期限', 'お支払期日', 'お支払期限', 'Payment Due', 'Due Date', 'DUE', 'Due'], fallbackYear
    );
  }

  // 明細金額の合計を補助として計算
  var itemsSum = 0;
  if (result.items && result.items.length > 0) {
    for (var ii = 0; ii < result.items.length; ii++) {
      itemsSum += result.items[ii].amount || 0;
    }
  }

  // 小計が取れていない場合は明細合計で補完
  if (!result.subtotal && itemsSum > 0) {
    result.subtotal = itemsSum;
  }

  // 合計-小計で消費税を逆算
  if (!result.taxAmount && result.total && result.subtotal && result.total > result.subtotal) {
    result.taxAmount = result.total - result.subtotal;
  }

  // 小計+消費税から合計を逆算
  if (!result.total && result.subtotal && result.taxAmount) {
    result.total = result.subtotal + result.taxAmount;
  }

  // 合計-消費税から小計を逆算
  if (!result.subtotal && result.total && result.taxAmount) {
    result.subtotal = result.total - result.taxAmount;
  }

  return result;
}

/**
 * 正規表現パース結果が不十分でGeminiを呼ぶべきかを判定
 * @param {Object} invoice
 * @param {string} contentType
 * @return {boolean}
 */
function shouldCallGemini_(invoice, contentType) {
  if (!CFG.gemini || !CFG.gemini.enabled) return false;

  // 分類失敗(その他/表/テキスト)は対象外(ビジネス書類じゃない可能性)
  if (contentType === 'table' || contentType === 'text') return false;

  // 重要フィールドが欠落していれば呼ぶ
  if (!invoice.vendorName) return true;
  if (!invoice.total) return true;
  if (!invoice.items || invoice.items.length === 0) {
    if (!invoice.subtotal) return true;
  }

  // 品質スコアが閾値未満なら呼ぶ
  var score = scoreParseResult(invoice);
  var threshold = CFG.gemini.scoreThreshold || 8;
  if (score < threshold) return true;

  // 明細の品名が全て空(名前欠落テーブルなど)も呼ぶ
  if (invoice.items && invoice.items.length > 0) {
    var hasName = false;
    for (var i = 0; i < invoice.items.length; i++) {
      if (invoice.items[i].name && invoice.items[i].name.length > 0) {
        hasName = true;
        break;
      }
    }
    if (!hasName) return true;
  }

  return false;
}

/**
 * Gemini の結果を正規表現結果にマージ(Gemini優先、不足分は既存値)
 * @param {Object} base - 正規表現パース結果
 * @param {Object} gemini - Gemini抽出結果
 * @return {Object}
 */
function mergeGeminiResult_(base, gemini) {
  var merged = {
    invoiceNumber:      gemini.invoiceNumber      || base.invoiceNumber      || '',
    issueDate:          gemini.issueDate          || base.issueDate          || '',
    vendorName:         gemini.vendorName         || base.vendorName         || '',
    recipientName:      gemini.recipientName      || base.recipientName      || '',
    items:              (gemini.items && gemini.items.length > 0) ? gemini.items : (base.items || []),
    subtotal:           gemini.subtotal           || base.subtotal           || 0,
    taxAmount:          gemini.taxAmount          || base.taxAmount          || 0,
    total:              gemini.total              || base.total              || 0,
    paymentDueDate:     gemini.paymentDueDate     || base.paymentDueDate     || '',
    registrationNumber: gemini.registrationNumber || base.registrationNumber || '',
    rawText:            base.rawText,
  };
  return merged;
}

/**
 * Gemini の docType 文字列を内部識別子にマップ
 * @param {string} docType
 * @return {string|null}
 */
function mapGeminiDocType_(docType) {
  if (!docType) return null;
  var m = {
    '請求書': 'invoice',
    '注文書': 'order',
    '見積書': 'quote',
    '領収書': 'receipt',
    '納品書': 'delivery',
    '契約書': 'contract',
    '議事録': 'minutes',
    '報告書': 'report',
  };
  return m[docType] || null;
}

/**
 * 文書中で最初に出現する4桁年(1990〜2099)を返す
 * @param {string} text
 * @return {string}
 */
function extractFirstYear(text) {
  var m = text.match(/\b(19|20)(\d{2})\b/);
  if (m) return m[1] + m[2];
  return '';
}

/**
 * 年なし日付(MM/DD など)をラベル近傍から探し、指定年で補完して返す
 * @param {string} text
 * @param {string[]} labels
 * @param {string} fallbackYear
 * @return {string} "YYYY/MM/DD" または空
 */
function findShortDateNearLabel(text, labels, fallbackYear) {
  var lines = text.split('\n');
  for (var v = 0; v < labels.length; v++) {
    var labRe = buildSpacedLabelRegex(labels[v], 'i');
    for (var i = 0; i < lines.length; i++) {
      var labMatch = lines[i].match(labRe);
      if (!labMatch) continue;

      // 英語ラベルの単語境界チェック
      var charBefore = labMatch.index > 0 ? lines[i].charAt(labMatch.index - 1) : '';
      var charAfter = lines[i].charAt(labMatch.index + labMatch[0].length);
      if (/[A-Za-z]/.test(charBefore) || /[A-Za-z]/.test(charAfter)) continue;

      // 同一行の末尾と次の3行を連結して探索
      var target = lines[i].substring(labMatch.index + labMatch[0].length);
      for (var k = i + 1; k < Math.min(i + 4, lines.length); k++) {
        if (!lines[k].trim()) continue;
        target += ' ' + lines[k];
        break;
      }

      // まず完全日付を試す
      var full = extractDate(target, target);
      if (full) return full;

      // MM/DD 形式を探す(括弧や他の数字との混同を避けるため厳しめ)
      var shortMatch = target.match(/(?:^|\s)(\d{1,2})[\/\-.](\d{1,2})(?!\d)/);
      if (shortMatch) {
        var mm = parseInt(shortMatch[1], 10);
        var dd = parseInt(shortMatch[2], 10);
        if (mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) {
          return fallbackYear + '/' + (shortMatch[1].length === 1 ? '0' + mm : mm) + '/' + (shortMatch[2].length === 1 ? '0' + dd : dd);
        }
      }
    }
  }
  return '';
}

/**
 * パース結果の品質スコアを計算(複数戦略の中から最良を選ぶため)
 * @param {Object} r
 * @return {number}
 */
function scoreParseResult(r) {
  var score = 0;
  if (r.vendorName) score += 2;
  if (r.issueDate) score += 1;
  if (r.total) score += 2;
  if (r.subtotal) score += 1;
  if (r.taxAmount) score += 1;
  if (r.items && r.items.length > 0) score += 2;
  if (r.paymentDueDate) score += 1;
  if (r.invoiceNumber) score += 1;
  return score;
}

/**
 * 請求番号を抽出
 * @param {string} normalized - 半角変換済みテキスト
 * @param {string} original - 元テキスト
 * @return {string}
 */
function extractInvoiceNumber(normalized, original) {
  // 許容文字: 英数字 + - _ と、ハイフン/スラッシュ周辺の空白も許容
  // 例: "NX -2026-0011" / "2026 / 0103" / "PO-2026-0234"
  var core = '[A-Za-z0-9](?:[A-Za-z0-9_]|\\s*-\\s*[A-Za-z0-9_])*(?:\\s*\\/\\s*[A-Za-z0-9\\-_]+)?';
  var patterns = [
    new RegExp('請求書番号[:：\\s]*(' + core + ')'),
    new RegExp('請求番号[:：\\s]*(' + core + ')'),
    new RegExp('注文書番号[:：\\s]*(' + core + ')'),
    new RegExp('注文番号[:：\\s]*(' + core + ')'),
    new RegExp('発注番号[:：\\s]*(' + core + ')'),
    new RegExp('伝票番号[:：\\s]*(' + core + ')'),
    new RegExp('契約番号[:：\\s]*(' + core + ')'),
    // No. はドット必須(NOTE等の誤マッチ回避)。単語境界は付けない(OCRが
    // "I N V O I C ENo." のように直前の英字と結合するケースに対応するため)
    new RegExp('No\\.\\s*(' + core + ')', 'i'),
    // No (ドットなし) は単語境界ありで安全側に
    new RegExp('\\bNo[:：\\s]+(' + core + ')(?![A-Za-z])', 'i'),
    // Invoice No (ドット省略、decompact後に "INVOICENO" になるケース)
    new RegExp('\\bInvoice\\s*No\\.?\\s*(' + core + ')', 'i'),
    // Invoice No の文字間に空白が入っているケース "INV O ICE N O" (OCR装飾)
    new RegExp('I\\s*n\\s*v\\s*o\\s*i\\s*c\\s*e\\s*N\\s*o\\.?[\\s　]*(' + core + ')', 'i'),
    new RegExp('\\bInvoice\\s*#?\\s*(' + core + ')', 'i'),
    new RegExp('\\bPO\\s*Number[:：\\s]*(' + core + ')', 'i'),
  ];

  for (var i = 0; i < patterns.length; i++) {
    var sourceText = normalized;
    var match = sourceText.match(patterns[i]);
    if (!match) {
      sourceText = original;
      match = sourceText.match(patterns[i]);
    }
    if (!match) {
      // ラベルと番号が別行にあるケースの行ベース検索
      var labelWords = ['請求書番号','請求番号','注文書番号','注文番号','発注番号','伝票番号','契約番号'];
      var sourceLines = original.split('\n');
      for (var lw = 0; lw < labelWords.length; lw++) {
        var labelRe = buildSpacedLabelRegex(labelWords[lw]);
        for (var ln = 0; ln < sourceLines.length; ln++) {
          if (!labelRe.test(sourceLines[ln])) continue;
          // 次の2行から英数字の番号を探す
          for (var nxt = ln + 1; nxt <= Math.min(ln + 2, sourceLines.length - 1); nxt++) {
            var numMatch = sourceLines[nxt].match(/^[\s　]*([A-Za-z0-9][A-Za-z0-9\-_]{2,30})(?:[\s　]|$)/);
            if (numMatch && /\d/.test(numMatch[1])) return numMatch[1];
          }
        }
      }
      continue;
    }

    // キャプチャが数字を1つも含まなければ本物の書類番号でない可能性が高い
    // (例: "Invoice PASS" の "PASS" を拾ってしまうケース)
    if (!/\d/.test(match[1])) continue;

    var captured = match[1].replace(/\s+/g, '').trim();

    // 末尾が "/YYYY" または "/YYYY-MM-DD" / "/YYYY/MM/DD" / "/YYYY.MM.DD" の場合は
    // スラッシュ以降が日付混入とみなして剥がす
    var dateTailMatch = captured.match(/^(.+?)\/(\d{4}[-\/.]\d{1,2}[-\/.]\d{1,2})$/);
    if (dateTailMatch) {
      captured = dateTailMatch[1];
    } else {
      // 単独の "/YYYY" は、後続の原文に 年/日付記号 が続く場合のみ剥がす
      var yearTailMatch = captured.match(/^(.+?)\/(\d{4})$/);
      if (yearTailMatch) {
        var afterEndIdx = match.index + match[0].length;
        var following = sourceText.substring(afterEndIdx, afterEndIdx + 10);
        if (/^[.\/-]\d{1,2}[.\/-]\d{1,2}/.test(following) || /^\s*年/.test(following)) {
          captured = yearTailMatch[1];
        }
      }
    }
    return captured;
  }

  // フォールバック: 文字間空白が 3 つ以下の "N O . PW- 2 0 2 6 - 0 7 1 6" 等は
  // 通常の decompactLetterSpacing(閾値 4+)では圧縮されないため、閾値を 2+ に
  // 下げた局所 decompact を施してから "N O ." パターンを試す。
  // 強い decompact は単語境界を潰してしまうため、ここでは厳密な書類番号フォーマット
  // (英字1-5 + 数字2+ + オプションの (-数字/英字) 繰り返し)で切り出す。
  var localAggressive = normalized.replace(/((?:[^\s\n]\s){2,}[^\s\n])/g, function(seg) {
    return seg.replace(/[ \u3000]/g, '');
  });
  if (localAggressive !== normalized) {
    var strictCore = '[A-Z]{1,5}[-_]?\\d{2,}(?:[-_]\\d+)*';
    var aggressiveRe = new RegExp('N\\s*O\\.\\s*(' + strictCore + ')', 'i');
    var aggMatch = localAggressive.match(aggressiveRe);
    if (aggMatch) {
      return aggMatch[1];
    }
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
    // スラッシュ・ハイフン・ドット区切り: 2024/01/15, 2024-01-15, 2024.01.15
    /(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/,
  ];

  // 和暦(漢数字): 令和八年四月十六日
  var kanjiReiwa = (normalized.match(/令和\s*([一二三四五六七八九十]+)\s*年\s*([一二三四五六七八九十]+)\s*月\s*([一二三四五六七八九十]+)\s*日/) ||
                    original.match(/令和\s*([一二三四五六七八九十]+)\s*年\s*([一二三四五六七八九十]+)\s*月\s*([一二三四五六七八九十]+)\s*日/));
  if (kanjiReiwa) {
    var ky = parseKanjiNumeral(kanjiReiwa[1]);
    var km = parseKanjiNumeral(kanjiReiwa[2]);
    var kd = parseKanjiNumeral(kanjiReiwa[3]);
    if (!isNaN(ky) && !isNaN(km) && !isNaN(kd)) {
      return (2018 + ky) + '/' + km + '/' + kd;
    }
  }

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

  // スラッシュ・ハイフン・ドット区切り
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
 *
 * 宛先と混同しないよう、以下のアルゴリズムでスコアリング:
 *   1. 宛先社名を候補から除外
 *   2. 住所(〒)・電話・登録番号 近傍をプラススコア
 *   3. 宛先位置より後ろをプラススコア
 *   4. 最高スコアを返す
 *
 * @param {string} text - 元テキスト
 * @return {string}
 */
function extractVendorName(text) {
  var recipient = extractRecipientName(text);

  // 社名候補を全列挙
  var candidates = [];

  // パターン1: "〇〇株式会社" / "〇〇有限会社" など(法人格が後ろ)
  // 区切り文字に ・ / 記号類 を含めて途中で途切れるようにする
  // 法人格は内部に空白が1個入っていても許容(OCRが "株 式会社" と誤分解するケース)
  var sep = '\\s、。\\n・／/|｜';
  var kabu = '株\\s?式\\s?会\\s?社';
  var yugen = '有\\s?限\\s?会\\s?社';
  var godo = '合\\s?同\\s?会\\s?社';
  var corpPat = '(?:' + kabu + '|' + yugen + '|' + godo + '|（株）|\\(株\\)|K\\.K\\.|G\\.K\\.|Inc\\.|LLC|Corporation|Corp\\.)';

  var candRe = new RegExp('(?:^|[' + sep + '])([^' + sep + ']{1,40}?' + corpPat + ')', 'g');
  var m;
  while ((m = candRe.exec(text)) !== null) {
    candidates.push({ name: m[1].replace(/\s+/g, '').trim(), index: m.index });
  }

  // パターン2: "株式会社〇〇" など(法人格が先頭)
  // 末尾オプション: スペース+カタカナ1〜5文字(OCRが "ネクストイノベーショ ン" と
  // 末尾を分割した場合に再結合する。御中/様/殿は漢字なのでここでは吸収されない)
  var preRe = new RegExp('(' + corpPat + '[^' + sep + ']{1,40}(?:\\s[\\u30A0-\\u30FF]{1,5})?)', 'g');
  while ((m = preRe.exec(text)) !== null) {
    candidates.push({ name: m[1].replace(/\s+/g, '').trim(), index: m.index });
  }

  // パターン3: 法人格を持たない屋号(ダミー製作所、サンプル工房、◯◯商店 など)
  //   〒 郵便番号の直前にある非空行を社名候補とみなす(請求書では発行元の頭に配置されやすい)
  var zipRe = /〒\s*\d{3}[-‐]?\d{4}/g;
  while ((m = zipRe.exec(text)) !== null) {
    // 〒の直前にある非空行を探す
    var prefix = text.substring(Math.max(0, m.index - 200), m.index);
    var prefixLines = prefix.split('\n').map(function(l) { return l.trim(); }).filter(function(l) { return l.length > 0; });
    if (prefixLines.length > 0) {
      var candidate = prefixLines[prefixLines.length - 1];
      // 見出し語や日付などを除外
      if (candidate.length >= 2 && candidate.length <= 40 &&
          !/^(FROM|発行元|請求元|BILL|Issuer|住所|電話|TEL|FAX)/i.test(candidate) &&
          !/^\d/.test(candidate) &&
          candidate !== recipient) {
        candidates.push({ name: candidate, index: m.index - candidate.length });
      }
    }
  }

  if (candidates.length === 0) return '';

  // 宛先位置(スコア加点の基準)
  var recipientIdx = recipient ? text.indexOf(recipient) : -1;

  // スコアリング
  var best = null;
  var bestScore = -Infinity;
  for (var i = 0; i < candidates.length; i++) {
    var cand = candidates[i];
    if (recipient && cand.name === recipient) continue;  // 宛先と同一は除外
    if (recipient && cand.name.indexOf(recipient) !== -1) continue;

    var score = 0;
    // 近傍(前後100文字)を検査
    var start = Math.max(0, cand.index - 50);
    var end = Math.min(text.length, cand.index + cand.name.length + 150);
    var context = text.substring(start, end);

    if (/〒\s*\d{3}[-‐]?\d{4}/.test(context)) score += 5;
    if (/TEL|電話|Tel|Phone|FAX|Fax/i.test(context)) score += 3;
    if (/登録番号|T\d{13}/.test(context)) score += 10;
    if (/発行元|請求元|Issuer|From|BILL\s*FROM/i.test(context)) score += 5;
    if (recipientIdx !== -1 && cand.index > recipientIdx) score += 2;

    if (score > bestScore) {
      bestScore = score;
      best = cand;
    }
  }

  if (best) return best.name;

  // スコアが付かなくても、宛先以外の最初の候補を返す
  for (var j = 0; j < candidates.length; j++) {
    if (!recipient || candidates[j].name !== recipient) {
      return candidates[j].name;
    }
  }
  return '';
}

/**
 * 宛先を抽出
 * @param {string} text - 元テキスト
 * @return {string}
 */
function extractRecipientName(text) {
  // パターン1: 敬称直前の会社名/個人名 (御中|様|さま|殿)
  var honorMatch = text.match(/([^\s、。\n]+?)\s*(?:御中|様|さま|殿)/);
  if (honorMatch) return honorMatch[1].trim();

  // パターン2: 宛先ラベルの次の行(英日混在レイアウト)
  var labels = ['宛先', '注文先', '請求先', '発注先', 'BILL TO', 'BILLTO', 'BUYER', 'TO'];
  var lines = text.split('\n');
  for (var li = 0; li < labels.length; li++) {
    var labRe = buildSpacedLabelRegex(labels[li], 'i');
    for (var i = 0; i < lines.length; i++) {
      if (!labRe.test(lines[i])) continue;
      // ラベル以降の同一行に名前があるか
      var afterMatch = lines[i].match(labRe);
      var rest = lines[i].substring(afterMatch.index + afterMatch[0].length).trim();
      if (rest.length >= 2 && rest.length <= 40 && !/^[:：]/.test(rest)) {
        return rest.replace(/^[:：\s]+/, '').trim();
      }
      // 同一行に名前がなければ次の非空行を返す
      for (var j = i + 1; j < Math.min(i + 3, lines.length); j++) {
        var candidate = lines[j].trim();
        if (candidate.length >= 2 && candidate.length <= 40) {
          return candidate;
        }
      }
    }
  }

  return '';
}

/**
 * 支払期限を抽出
 * @param {string} text - 正規化済みテキスト
 * @return {string} 日付文字列 (yyyy/MM/dd) または空
 */
function extractPaymentDueDate(text) {
  // 具体的なラベルを先に、汎用の "Due" を最後に
  var labels = [
    '支払期日', '支払期限', 'お支払期日', 'お支払期限',
    '支払日', 'お支払日', '振込期限', '振込期日',
    '納期',
    'Payment Due', 'Due Date', 'Pay By', 'Due',
  ];

  var lines = text.split('\n');
  for (var v = 0; v < labels.length; v++) {
    var lab = labels[v];
    var labRe = buildSpacedLabelRegex(lab, 'i');
    for (var i = 0; i < lines.length; i++) {
      var labMatch = lines[i].match(labRe);
      if (!labMatch) continue;

      // 英語ラベルが他の単語の一部になっていないか確認(Due が Reduced の一部等を回避)
      var charBefore = labMatch.index > 0 ? lines[i].charAt(labMatch.index - 1) : '';
      var charAfter = lines[i].charAt(labMatch.index + labMatch[0].length);
      if (/[A-Za-z]/.test(charBefore) || /[A-Za-z]/.test(charAfter)) continue;

      // ラベル位置以降の同一行と次の行を結合して日付抽出
      var target = lines[i].substring(labMatch.index + labMatch[0].length);
      if (i + 1 < lines.length) target += ' ' + lines[i + 1];

      var dateStr = extractDate(target, target);
      if (dateStr) return dateStr;
    }
  }

  return '';
}

/**
 * 適格請求書の複数税率(8%/10%)内訳から小計・消費税の合計を算出
 * 軽減税率が含まれない通常請求書では {0, 0} を返し、呼び出し元は通常の extractAmount を使う
 * @param {string} text
 * @return {{subtotal: number, taxAmount: number}}
 */
function extractMultiTaxRateTotals(text) {
  // 「軽減税率」キーワードが無ければ適格請求書ではないと判定
  if (text.indexOf('軽減税率') === -1) return { subtotal: 0, taxAmount: 0 };

  var subtotal = 0;
  var tax = 0;
  var foundRates = 0;

  var lines = text.split('\n');
  for (var i = 0; i < lines.length; i++) {
    // 税率対象の行だけ検出(軽減税率という注釈単独は除外)
    if (!/(8\s*%|10\s*%)\s*対象/.test(lines[i])) continue;

    // 同一行で2つの数値が取れなければ、次の2行を結合して再取得
    var nums = collectAmountCandidates(lines[i]);
    if (nums.length < 2) {
      var combined = lines[i];
      if (i + 1 < lines.length) combined += ' ' + lines[i + 1];
      if (i + 2 < lines.length) combined += ' ' + lines[i + 2];
      nums = collectAmountCandidates(combined);
    }
    if (nums.length < 2) continue;

    // 先頭2つが [対象金額, 消費税額]
    subtotal += nums[0].value;
    tax += nums[1].value;
    foundRates++;
  }

  if (foundRates === 0) return { subtotal: 0, taxAmount: 0 };
  return { subtotal: subtotal, taxAmount: tax };
}

/**
 * 適格請求書(インボイス制度)の登録番号を抽出
 * パターン: "登録番号 T1234567890123" または単独の "T" + 13桁
 * @param {string} text
 * @return {string}
 */
function extractRegistrationNumber(text) {
  var m = text.match(/登録番号[：: 　]*T?\s*(\d{13})/);
  if (m) return 'T' + m[1];
  m = text.match(/\bT(\d{13})\b/);
  if (m) return 'T' + m[1];
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

  // 明細ヘッダー行を検出（英語/平仮名レイアウトも対応）
  var nameHeaderWords = ['品名', '品目', '内容', '摘要', 'Description', 'Item', 'なに'];
  var dataHeaderWords = ['金額', '数量', '単価', 'Amount', 'Qty', 'Quantity', 'Price', 'きんがく', 'たんか', 'かず'];
  var headerIndex = -1;
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    if (containsAny(line, nameHeaderWords) && containsAny(line, dataHeaderWords)) {
      headerIndex = i;
      break;
    }
  }

  // ヘッダー行が見つからない場合は別レイアウト用のパーサにフォールバック
  if (headerIndex === -1) {
    var cellPerLine = extractItemsCellPerLine(normalized);
    if (cellPerLine.length > 0) return cellPerLine;
    var colMajor = extractItemsColumnMajor(normalized);
    if (colMajor.length > 0) return colMajor;
    var inlineItems = extractItemsInline(normalized);
    if (inlineItems.length > 0) return inlineItems;
    return extractItemsCrossPattern(normalized);
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
    var cellPerLine2 = extractItemsCellPerLine(normalized);
    if (cellPerLine2.length > 0) return cellPerLine2;
    var colMajor2 = extractItemsColumnMajor(normalized);
    if (colMajor2.length > 0) return colMajor2;
    var inline2 = extractItemsInline(normalized);
    if (inline2.length > 0) return inline2;
    return extractItemsCrossPattern(normalized);
  }

  return items;
}

/**
 * 指定文字列にいずれかの単語が含まれるか判定
 * @param {string} str
 * @param {string[]} words
 * @return {boolean}
 */
function containsAny(str, words) {
  for (var i = 0; i < words.length; i++) {
    if (str.indexOf(words[i]) !== -1) return true;
  }
  return false;
}

/**
 * 単一行に全明細が詰まっているレイアウトから抽出
 *
 * 例: "Description Qty Unit Price Amount コンサル費用 10 時間 ¥15,000 ¥150,000 撮影・編集費 1 式 ¥180,000 ¥180,000 ..."
 * 戦略:
 *   1. ヘッダー語(品名/Description/なに等)と金額語(金額/Amount/きんがく等)の後から開始
 *   2. 合計/Subtotal/Tax で終端
 *   3. "名前 数量 (単位) ¥単価 ¥金額" の反復パターンを正規表現で抽出
 *
 * @param {string} normalized
 * @return {Object[]}
 */
function extractItemsInline(normalized) {
  // 改行を空白に変換して1本の文字列にする
  // 装飾スペース(D E S C R I P T I O N等)を圧縮してヘッダー検出の確度を上げる
  var flat = decompactLetterSpacing(normalized.replace(/\n/g, ' '));

  // ヘッダー終端位置を特定
  var headerEndRe = /(Description|Item|品名|品目|内容|摘要|なに)[^¥\n]{0,80}?(Amount|金額|きんがく)/i;
  var headerMatch = flat.match(headerEndRe);
  var startIdx = headerMatch ? (headerMatch.index + headerMatch[0].length) : 0;

  // 終端境界を探す
  var endRe = /(Subtotal|小計|税抜合計|Total|合計|Tax|消費税|Grand\s*Total|Amount\s*Due)/i;
  var tail = flat.substring(startIdx);
  var endMatch = tail.match(endRe);
  var endIdx = endMatch ? (startIdx + endMatch.index) : flat.length;

  var section = flat.substring(startIdx, endIdx);

  // "品名 数量 (単位?) ¥単価 ¥金額" を反復抽出
  // 品名の長さは60文字以内に制限(プレリュード全部を吸収するのを防ぐ)
  var re = /([^¥\n]{1,60}?)\s+(\d+(?:\.\d+)?)\s*([^\s¥\d\n]{0,6})?\s*[¥￥]([\d,]+)\s*[¥￥]([\d,]+)/g;
  var items = [];
  var m;
  while ((m = re.exec(section)) !== null) {
    var name = m[1].trim();
    // 先頭にヘッダー残骸があれば除去
    name = name.replace(/^(Description|Qty|Quantity|Unit\s*Price|Amount|品名|品目|数量|単価|金額|なに|かず|たんい|たんか|きんがく)\s*/i, '').trim();
    // 純粋な数字のみの名前(No列の値)は省略(品名欠落レイアウトの誤検出を避ける)
    if (/^\d{1,3}$/.test(name)) name = '';
    var amountNum = parseNumber(m[5]);
    if (amountNum <= 0) continue;
    items.push({
      name: name,
      quantity: parseNumber(m[2]),
      unitPrice: parseNumber(m[4]),
      amount: amountNum,
    });
  }

  return items;
}

/**
 * "品名 数量単位 × ¥単価" 形式の明細パーサー
 *
 * 例:
 *   SaaS月額利用料 1月 × ¥248,000
 *   初期セットアップ費用
 *   1式 × ¥180,000     ← 直前の行が品名
 *
 * 金額は数量×単価で計算する(原文に明細金額の記載がないケース向け)
 * @param {string} normalized
 * @return {Object[]}
 */
function extractItemsCrossPattern(normalized) {
  var items = [];
  var lines = normalized.split('\n');
  var pendingName = '';

  // 単行完結: "名前 数量単位 × ¥単価"
  var fullRe = /^(.+?)\s+(\d+(?:\.\d+)?)\s*([^\s\d×xX¥]{0,6})?\s*[×xX]\s*[¥￥]([\d,]+)/;
  // 数量+単価のみ: "数量単位 × ¥単価"(直前行が名前)
  var valueOnlyRe = /^\s*(\d+(?:\.\d+)?)\s*([^\s\d×xX¥]{0,6})?\s*[×xX]\s*[¥￥]([\d,]+)/;

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) { continue; }

    // 打ち切りラベル(明細の下に現れる集計行)
    // "請求先" "支払条件" など前部の見出しは除外し、具体的な集計ラベルのみ指定
    if (/^(小計|合計|消費税|消費税額|税額|総額|総計|お振込|振込先|ご請求額|ご請求金額|お支払金額|Subtotal|Sub\s*Total|Grand\s*Total|Total\b|Tax\b)/i.test(line)) {
      break;
    }

    var full = line.match(fullRe);
    if (full) {
      var qty = parseNumber(full[2]);
      var price = parseNumber(full[4]);
      items.push({
        name: full[1].trim(),
        quantity: qty,
        unitPrice: price,
        amount: qty * price,
      });
      pendingName = '';
      continue;
    }

    var valueOnly = line.match(valueOnlyRe);
    if (valueOnly && pendingName) {
      var qty2 = parseNumber(valueOnly[1]);
      var price2 = parseNumber(valueOnly[3]);
      items.push({
        name: pendingName,
        quantity: qty2,
        unitPrice: price2,
        amount: qty2 * price2,
      });
      pendingName = '';
      continue;
    }

    // 品名候補として保留(長すぎ/数字始まり/金額行は除外)
    if (line.length >= 2 && line.length <= 60 &&
        !/^[¥￥\d\-]/.test(line) &&
        !/^(No\.?|Issue|Date|登録)/i.test(line)) {
      pendingName = line;
    }
  }

  return items;
}

/**
 * 1セル1行レイアウトから明細を抽出
 * 品名/数量/単位/単価/金額 が縦に並ぶケース(4〜6列の可変)
 * 英語(Description/Qty/Unit/Price/Amount)・平仮名(なに/かず/たんい/たんか/きんがく)にも対応
 * @param {string} normalized - 半角変換済みテキスト
 * @return {Object[]}
 */
function extractItemsCellPerLine(normalized) {
  function strip(s) {
    return (s || '').replace(/[\s\u3000\u00a0]+/g, '');
  }

  var lines = normalized.split('\n')
    .map(function(l) { return l.replace(/\s+$/, '').replace(/^\s+/, ''); })
    .filter(function(l) { return l.length > 0; });

  // ラベル語は小文字化して格納(大文字小文字を無視)
  var nameHeaderWords = {'品名':1,'品目':1,'品番':1,'品番・品名':1,'内容':1,'摘要':1,'なに':1,'description':1,'item':1};
  var qtyWords       = {'数量':1,'かず':1,'qty':1,'quantity':1};
  var unitLabelWords = {'単位':1,'たんい':1,'unit':1};
  var priceWords     = {'単価':1,'たんか':1,'unitprice':1,'price':1};
  var amountWords    = {'金額':1,'きんがく':1,'amount':1};
  var taxRateWords   = {'税率':1,'taxrate':1};
  var noWords        = {'no':1,'no.':1,'#':1,'番号':1,'no,':1};

  function keyOf(s) { return strip(s).toLowerCase(); }

  // 名前ヘッダーを探す(単独行、10文字以下)
  var nameHeaderIdx = -1;
  for (var i = 0; i < lines.length; i++) {
    var clean = strip(lines[i]);
    if (clean.length > 12) continue;
    if (nameHeaderWords[clean.toLowerCase()]) {
      nameHeaderIdx = i;
      break;
    }
  }
  if (nameHeaderIdx === -1) return [];

  // 名前ヘッダーの直前に "No" / "番号" 列があるか確認
  var hasNoColumn = nameHeaderIdx > 0 && noWords[keyOf(lines[nameHeaderIdx - 1])];

  // 名前ヘッダー以降の列ヘッダーを順に探索
  var taxRateIdx = -1, qtyIdx = -1, unitLabelIdx = -1, priceIdx = -1, amountIdx = -1;
  var searchEnd = Math.min(lines.length, nameHeaderIdx + 10);
  for (var j = nameHeaderIdx + 1; j < searchEnd; j++) {
    var c = keyOf(lines[j]);
    if (c.length > 12) continue;
    if (taxRateWords[c] && taxRateIdx === -1) taxRateIdx = j;
    else if (qtyWords[c] && qtyIdx === -1) qtyIdx = j;
    else if (unitLabelWords[c] && unitLabelIdx === -1) unitLabelIdx = j;
    else if (priceWords[c] && priceIdx === -1) priceIdx = j;
    else if (amountWords[c] && amountIdx === -1) amountIdx = j;
  }

  if (qtyIdx === -1 || amountIdx === -1) return [];

  // 列順: (No?) → 品名 → (税率?) → 数量 → (単位?) → (単価?) → 金額
  var columnOrder = [];
  if (hasNoColumn) columnOrder.push('no');
  columnOrder.push('name');
  if (taxRateIdx !== -1) columnOrder.push('taxRate');
  columnOrder.push('qty');
  if (unitLabelIdx !== -1) columnOrder.push('unit');
  if (priceIdx !== -1) columnOrder.push('price');
  columnOrder.push('amount');
  var cols = columnOrder.length;

  var startIdx = Math.max(nameHeaderIdx, taxRateIdx, qtyIdx, unitLabelIdx, priceIdx, amountIdx) + 1;

  // データ行の name セルが実は数量(数字)になっていないか確認
  // (品名が表の上にまとめ書きされ、データ行には No と数量以降しか無いレイアウト)
  if (hasNoColumn && columnOrder.indexOf('name') !== -1 && startIdx + 1 < lines.length) {
    var noCellIdx = columnOrder.indexOf('no');
    var nameCellIdx = columnOrder.indexOf('name');
    var noVal = parseNumber(lines[startIdx + noCellIdx]);
    var nameVal = parseNumber(lines[startIdx + nameCellIdx]);
    // 両方数値なら name 列は存在しない → 取り除く
    if (noVal > 0 && nameVal > 0 && nameVal < 10000) {
      columnOrder.splice(nameCellIdx, 1);
      cols = columnOrder.length;
    }
  }

  var items = [];
  var idx = startIdx;
  while (idx + cols - 1 < lines.length) {
    var record = {};
    for (var k = 0; k < cols; k++) {
      record[columnOrder[k]] = lines[idx + k];
    }

    // 合計部や空行に到達したら終了
    var nameStripped = strip(record.name);
    if (/^(小計|合計|消費税|税込|税抜|総額|総計|請求|お振込|お支払|区分|Subtotal|Total|Tax|GrandTotal)/i.test(nameStripped)) break;

    // ハイフンだけの空行はスキップ
    if (nameStripped === '-' || nameStripped === '−' || nameStripped === '–' || nameStripped === 'ー') {
      idx += cols;
      continue;
    }

    var amountNum = parseNumber(record.amount);
    if (record.name && amountNum > 0) {
      items.push({
        name: record.name,
        quantity: parseNumber(record.qty || 0),
        unitPrice: parseNumber(record.price || record.amount || 0),
        amount: amountNum,
      });
    }
    idx += cols;
  }

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
 *
 * 特徴:
 * - 日本語/英語/装飾スペースすべてに対応(buildSpacedLabelRegex で吸収)
 * - "10%" など税率を金額と誤認しないよう % 後続の数値は除外
 * - 具体ラベル優先(Subtotal/Grand Total を Total より先に判定し、ヒット行は除外)
 * - ラベル行に数値がなければ次の行も探索
 * - 同一行に複数値があれば最大値を採用
 *
 * @param {string} text - 正規化済みテキスト(normalizeOcrText適用済み想定)
 * @param {string} label - ラベル種別 '合計' | '小計' | '消費税'
 * @return {number}
 */
function extractAmount(text, label) {
  var labelVariants = getAmountLabelVariants(label);

  // より具体的なラベルでヒットした行番号を記録し、後続の汎用ラベルの対象外にする
  // (例: "Subtotal" マッチ行を "Total" 検索時に無視することで Subtotal 誤ヒットを回避)
  var excludedLines = {};

  for (var v = 0; v < labelVariants.length; v++) {
    var lab = labelVariants[v];
    var labRe = buildSpacedLabelRegex(lab, 'i');
    var lines = text.split('\n');

    for (var i = 0; i < lines.length; i++) {
      if (excludedLines[i]) continue;
      var labMatch = lines[i].match(labRe);
      if (!labMatch) continue;

      // 前後が英数字なら部分文字列マッチ(例: "Total" が "Subtotal" の一部)としてスキップ
      var charBefore = labMatch.index > 0 ? lines[i].charAt(labMatch.index - 1) : '';
      var charAfter = lines[i].charAt(labMatch.index + labMatch[0].length);
      if (/[A-Za-z0-9]/.test(charBefore) || /[A-Za-z0-9]/.test(charAfter)) continue;

      // ラベル位置以降の同一行
      var afterLabel = lines[i].substring(labMatch.index + labMatch[0].length);

      // "Rate" / "率" が続く場合は税率表記であり金額ではない → スキップ
      // (例: "10% TAX RATE" の Tax を金額ラベルと誤認しない)
      if (/^[\s:：]*(Rate|rate|RATE|率)/.test(afterLabel)) continue;

      // "税" 単独ラベルは「税込」「税別」「税率」の前半にもマッチするので除外
      if (lab === '税' && /^[込別率]/.test(afterLabel)) continue;

      excludedLines[i] = true;
      var nums = collectAmountCandidates(afterLabel);

      // 同一行で ¥ 付きの値が取れなかった場合、後続5行まで空行スキップして探索
      // 住所の郵便番号等を誤取得しないよう、¥ 付きを最優先で返す
      // 他の金額ラベル(小計/消費税/合計 等)が出現したらそこで打ち切り
      if (!hasYenValue(nums)) {
        var fallbackNums = nums;
        for (var k = i + 1; k < Math.min(i + 6, lines.length); k++) {
          var nextLine = lines[k];
          if (!nextLine.trim()) continue;
          if (isAnotherAmountLabel(nextLine)) break;
          var nextNums = collectAmountCandidates(nextLine);
          if (nextNums.length === 0) continue;
          if (hasYenValue(nextNums)) { nums = nextNums; break; }
          if (!hasYenValue(fallbackNums)) fallbackNums = nextNums;
        }
        if (!hasYenValue(nums) && fallbackNums.length > 0) nums = fallbackNums;
      }

      if (nums.length > 0) {
        return pickBestAmount(nums);
      }
    }
  }

  return 0;
}

/**
 * 論理ラベルから表記バリエーション配列を取得(具体的な順に並ぶ)
 * @param {string} label - '合計' | '小計' | '消費税'
 * @return {string[]}
 */
function getAmountLabelVariants(label) {
  if (label === '消費税') {
    return [
      '消費税額等', '消費税額', '消費税等', '内消費税', 'うち消費税',
      'Sales Tax', 'VAT', 'GST',
      '税額', '消費税', 'Tax',
      '税',  // 最後の手段(短い単独ラベル、他の税系ラベルが先にマッチするので安全)
    ];
  }
  if (label === '小計') {
    return [
      '小計金額', '税抜合計', '税抜金額',
      'Sub-Total', 'Sub Total', 'Subtotal',
      '小計',
    ];
  }
  if (label === '合計') {
    // 具体的な表記を先に、汎用な 'Total' / '合計' を最後に
    return [
      'ご請求金額', '御請求金額', 'ご請求額', '請求金額',
      '合計金額', '税込合計', 'お支払金額', '御注文金額', '注文金額合計',
      '総額', '総計',
      'Grand Total', 'Total Due', 'Balance Due', 'Amount Due',
      '合計', 'Total',
    ];
  }
  return [label];
}

/**
 * 行から金額候補の数値を列挙(% 後続は除外)
 * @param {string} line
 * @return {number[]}
 */
function collectAmountCandidates(line) {
  var amounts = [];
  // ¥記号が直前にあったかを各候補に記録
  var re = /([¥￥])?\s*([\d,]+)(\s*[%％])?/g;
  var m;
  while ((m = re.exec(line)) !== null) {
    if (m[3]) continue;  // % 付き → 税率
    var n = parseNumber(m[2]);
    if (n > 0) amounts.push({ value: n, hasYen: !!m[1] });
  }
  return amounts;
}

/**
 * 金額候補配列から「最もそれっぽい」値を1つ返す
 * - ¥ 記号付きがあれば優先
 * - なければ先頭を採用
 * @param {Array<{value:number,hasYen:boolean}>} candidates
 * @return {number}
 */
function pickBestAmount(candidates) {
  if (!candidates || candidates.length === 0) return 0;
  for (var i = 0; i < candidates.length; i++) {
    if (candidates[i].hasYen) return candidates[i].value;
  }
  return candidates[0].value;
}

/**
 * 金額候補配列に¥記号付きの値が含まれるか
 * @param {Array<{value:number,hasYen:boolean}>} candidates
 * @return {boolean}
 */
function hasYenValue(candidates) {
  if (!candidates) return false;
  for (var i = 0; i < candidates.length; i++) {
    if (candidates[i].hasYen) return true;
  }
  return false;
}

/**
 * 行が別の金額ラベル行かを判定(extractAmount の値探索打ち切り用)
 * @param {string} line
 * @return {boolean}
 */
function isAnotherAmountLabel(line) {
  return /(小計|消費税|税額|合計|総額|総計|ご請求|御請求|請求金額|Subtotal|Sub\s*Total|Grand\s*Total|Amount\s*Due|Total\b|Tax\b|VAT|GST)/i.test(line);
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
  // 装飾で各文字が重複している場合("ごごせせいいききゅゅううししょょ" 等)に備えて
  // dedouble 版でもキーワードマッチを試みる
  var dedoubled = dedoubleRepeatedChars(text);
  var matches = function(pattern) {
    return pattern.test(text) || (dedoubled !== text && pattern.test(dedoubled));
  };

  // 契約書固有のタイトル表記は他の書類名を本文中に含みがちなので最優先で判定
  // (例: "業務委託契約書" の中の "請求書受領後" が 'invoice' と誤判定されるのを防ぐ)
  var strongContractTitles = [
    '業務委託契約書', '売買契約書', '雇用契約書', '賃貸借契約書',
    '秘密保持契約書', '業務提携契約書', '委任契約書', '請負契約書',
    '簡易業務委託契約書', 'Service Agreement', 'Non-Disclosure Agreement', 'NDA',
  ];
  for (var si = 0; si < strongContractTitles.length; si++) {
    if (matches(buildSpacedLabelRegex(strongContractTitles[si], 'i'))) {
      return 'contract';
    }
  }

  // 装飾スペースを吸収するため、キーワード群それぞれを buildSpacedLabelRegex で試す
  var categories = [
    ['quote',    ['見積書', '御見積書', '御見積', '見積金額', '見積番号', 'お見積', 'Quotation', 'Estimate']],
    ['order',    ['注文書', '御注文', '発注書', '注文番号', '発注番号', '工事注文書', 'Purchase Order', 'PurchaseOrder', 'Order']],
    ['delivery', ['納品書', '納品日', 'Delivery Note', 'DeliveryNote']],
    ['receipt',  ['領収書', '領収証', '但し書き', 'Receipt']],
    ['invoice',  ['適格請求書', 'ご請求書', '御請求書', '請求書', '御請求', '請求金額', '請求番号', 'せいきゅうしょ', 'Invoice']],
    ['minutes',  ['議事録', '出席者', '議題', '開催日時', 'Minutes']],
    ['contract', ['契約書', '契約期間', 'Agreement', 'Contract']],
    ['report',   ['報告書', '業務報告', '月次報告', 'Report']],
  ];

  for (var c = 0; c < categories.length; c++) {
    var type = categories[c][0];
    var keywords = categories[c][1];
    for (var k = 0; k < keywords.length; k++) {
      if (matches(buildSpacedLabelRegex(keywords[k], 'i'))) {
        return type;
      }
    }
  }

  // 「第◯条」は契約書の強いシグナル
  if (/第\s*[一二三四五六七八九十0-9]+\s*条/.test(text)) return 'contract';

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
