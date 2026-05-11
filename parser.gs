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
      // Gemini マージ後にも整合性が残っていれば warning を出す(運用観察用、補正はしない)
      if (hasConsistencyIssue_(invoice)) {
        console.warn('[parseLedgerEntry] Geminiマージ後も整合性エラーが残存: ' + fileName);
      }
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

  // 取引先(ユーザーから見た相手側)を決定
  //   'issuer'  : 自社=発行元 → 取引先=宛先(御中/様の側 = recipientName)
  //   'receiver': 自社=受領側 → 取引先=発行元(vendorName)
  var userRole = (CFG.ledger && CFG.ledger.userRole) || 'issuer';
  var counterparty = (userRole === 'issuer')
    ? (invoice.recipientName || invoice.vendorName || '')
    : (invoice.vendorName || invoice.recipientName || '');

  // ステータス判定: 取引先と合計金額の両方が取れていればOK
  var hasKeyInfo = !!(counterparty && invoice.total);
  var status = hasKeyInfo ? 'OK' : '部分抽出';

  return {
    processedAt: new Date(),
    fileName: fileName,
    fileLink: fileLink,
    docType: docTypeLabels[contentType] || 'その他',
    vendor: counterparty,
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

  // 多列ラベル "合計 税抜 消費税 総額" \n "¥200,000 ¥20,000 ¥220,000" のような
  // 表形式レイアウトを最優先で取得(通常の extractAmount より上の優先度)
  var triplet = extractAmountTriplet_(normalized);

  // 信頼度を別途取得(multiRate 経路の場合は信頼度を 100 とみなす = ラベル明示で取れた前提)
  var subtotalC = multiRate.subtotal
    ? { value: multiRate.subtotal, confidence: 100 }
    : (triplet && triplet.subtotal ? { value: triplet.subtotal, confidence: 90 }
                                   : extractAmountWithConfidence(normalized, '小計'));
  var taxC = multiRate.taxAmount
    ? { value: multiRate.taxAmount, confidence: 100 }
    : (triplet && triplet.taxAmount ? { value: triplet.taxAmount, confidence: 90 }
                                    : extractAmountWithConfidence(normalized, '消費税'));
  var totalC = (triplet && triplet.total)
    ? { value: triplet.total, confidence: 90 }
    : extractAmountWithConfidence(normalized, '合計');

  var result = {
    invoiceNumber: extractInvoiceNumber(normalized, original),
    issueDate: extractDate(normalized, original),
    vendorName: extractVendorName(normalized),
    recipientName: extractRecipientName(normalized),
    items: extractItems(normalized, original),
    subtotal: subtotalC.value,
    taxAmount: taxC.value,
    total: totalC.value,
    paymentDueDate: extractPaymentDueDate(normalized),
    registrationNumber: extractRegistrationNumber(normalized),
    rawText: original,
    _confidence: {
      subtotal: subtotalC.confidence,
      taxAmount: taxC.confidence,
      total: totalC.confidence,
    },
  };

  // 年なし日付のフォールバック: MM/DD のみでも文書中の年で補完
  var fallbackYear = extractFirstYear(normalized) || new Date().getFullYear();
  if (!result.issueDate) {
    result.issueDate = findShortDateNearLabel(
      normalized,
      ['発行日', '請求日', '日付', '作成日', '年月日', 'Issue', 'ISSUED', 'Date', 'DATE'],
      fallbackYear
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

  // 源泉徴収補正:
  //   原文に「源泉徴収」がある場合、'ご請求金額合計'/'差引請求額' などのラベルから
  //   源泉控除後の金額を total として拾った可能性が高い。
  //   subtotal + taxAmount = gross total が成立するなら、それを total に採用する。
  //   (例: subtotal=4100, tax=410, total=4092 → grossTotal=4510 に置換)
  if (/源泉徴収/.test(normalized) && result.subtotal > 0 && result.taxAmount > 0) {
    var grossTotal = result.subtotal + result.taxAmount;
    if (result.total === 0 || (result.total < grossTotal && Math.abs(result.total - grossTotal) > 2)) {
      result.total = grossTotal;
    }
  }

  // 内税レイアウト補正:
  //   subtotal == total かつ tax > 0 の状態は「内税表記(税込合計と内消費税のみ表示)」
  //   から subtotal を別ラベル "金額" などで誤って total と同値で取った可能性が高い。
  //   原文に「内税」「税込」シグナルがあれば subtotal を税抜に補正する。
  if (result.total > 0 && result.subtotal === result.total && result.taxAmount > 0 &&
      /(内税|税込|内消費税|うち消費税|（税込）|\(税込\)|税込価格|税込合計|価格\s*\(税込\)|価格\s*（税込）)/.test(normalized) &&
      result.subtotal > result.taxAmount) {
    result.subtotal = result.total - result.taxAmount;
  }

  return result;
}

/**
 * 抽出済みデータの内部整合性をチェックする
 * 正規表現が間違った値を拾った場合の保険(フィールドの有無だけでは検出できない誤抽出を補う)
 * @param {Object} invoice
 * @return {boolean} 矛盾を検出したら true
 */
function hasConsistencyIssue_(invoice) {
  // 1. 金額式: 外税(subtotal + tax = total)または内税(subtotal = total、tax は内訳)のどちらかが成立すればOK
  if (invoice.total > 0 && invoice.subtotal > 0 && invoice.taxAmount > 0) {
    var exclusiveDiff = Math.abs((invoice.subtotal + invoice.taxAmount) - invoice.total);
    var inclusiveDiff = Math.abs(invoice.subtotal - invoice.total);
    // どちらも許容範囲外なら矛盾
    if (exclusiveDiff > 2 && inclusiveDiff > 2) {
      console.warn('[consistency] 金額式不一致: subtotal(' + invoice.subtotal +
                   ') tax(' + invoice.taxAmount + ') total(' + invoice.total +
                   ') 外税差=' + exclusiveDiff + ' 内税差=' + inclusiveDiff);
      return true;
    }
    // 税が小計を超える、または合計を超えるのは常識的に不可能(消費税率は通常 8-10%)
    // OCR で桁が落ちて total=245, tax=455 のような結果になった場合に検出する
    if (invoice.taxAmount > invoice.subtotal || invoice.taxAmount > invoice.total) {
      console.warn('[consistency] 消費税が小計/合計を超える: subtotal(' + invoice.subtotal +
                   ') tax(' + invoice.taxAmount + ') total(' + invoice.total + ')');
      return true;
    }
  }
  // 合計が小計より小さい(税抜が税込より大きい)のも矛盾
  if (invoice.total > 0 && invoice.subtotal > 0 && invoice.total < invoice.subtotal &&
      invoice.subtotal - invoice.total > 2) {
    console.warn('[consistency] 合計 < 小計: subtotal(' + invoice.subtotal +
                 ') total(' + invoice.total + ')');
    return true;
  }

  // 2. 明細合計: 外税(sum ≒ subtotal 税抜)/内税(sum ≒ subtotal 税込=total)の両方を許容
  //   ターゲット候補として [subtotal, total, subtotal-tax, total-tax] の最も近い値と比較
  if (invoice.items && invoice.items.length >= 2 && invoice.subtotal > 0) {
    var allHaveAmount = invoice.items.every(function(it) { return it.amount > 0; });
    if (allHaveAmount) {
      var itemsSum = invoice.items.reduce(function(a, it) { return a + it.amount; }, 0);
      var targets = [invoice.subtotal];
      if (invoice.total > 0) targets.push(invoice.total);
      if (invoice.taxAmount > 0) {
        targets.push(invoice.subtotal - invoice.taxAmount);
        if (invoice.total > 0) targets.push(invoice.total - invoice.taxAmount);
      }
      var minDiff = Infinity;
      for (var ti = 0; ti < targets.length; ti++) {
        var d = Math.abs(itemsSum - targets[ti]);
        if (d < minDiff) minDiff = d;
      }
      var ratio = minDiff / invoice.subtotal;
      if (ratio > 0.02 && minDiff > 100) {
        console.warn('[consistency] 明細合計不一致: itemsSum(' + itemsSum +
                     ') が subtotal/total/税抜後 のいずれとも不一致 minDiff=' + minDiff);
        return true;
      }
    }
  }

  // 3. 合計金額の妥当範囲(100円未満の請求書はほぼ存在しない)
  if (invoice.total > 0 && invoice.total < 100) {
    console.warn('[consistency] 合計金額が常識的範囲外: total=' + invoice.total);
    return true;
  }

  return false;
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

  // 取引先(ユーザーから見た相手側)を userRole に応じて決定
  var userRole = (CFG.ledger && CFG.ledger.userRole) || 'issuer';
  var counterparty = (userRole === 'issuer') ? invoice.recipientName : invoice.vendorName;

  // 重要フィールドが欠落していれば呼ぶ
  if (!counterparty) return true;
  if (!invoice.total) return true;
  if (!invoice.items || invoice.items.length === 0) {
    if (!invoice.subtotal) return true;
  }

  // 整合性チェック(計算の自己矛盾があれば誤抽出の可能性が高いので LLM 救済対象)
  if (hasConsistencyIssue_(invoice)) return true;

  // 信頼度チェック: 値は取れているが ¥なし・遠方の行で取得 = 誤検出の可能性
  // invoice._confidence は parseInvoiceOnce が付与する各金額の信頼度スコア
  if (invoice._confidence) {
    var low = 50;
    if (invoice._confidence.total && invoice._confidence.total < low) {
      console.log('[shouldCallGemini] total の信頼度が低い: ' + invoice._confidence.total);
      return true;
    }
    if (invoice.subtotal > 0 && invoice._confidence.subtotal && invoice._confidence.subtotal < low) {
      console.log('[shouldCallGemini] subtotal の信頼度が低い: ' + invoice._confidence.subtotal);
      return true;
    }
    if (invoice.taxAmount > 0 && invoice._confidence.taxAmount && invoice._confidence.taxAmount < low) {
      console.log('[shouldCallGemini] taxAmount の信頼度が低い: ' + invoice._confidence.taxAmount);
      return true;
    }
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
  // 取引先(ユーザーから見た相手側)を userRole に応じて決定
  var userRole = (CFG.ledger && CFG.ledger.userRole) || 'issuer';
  var counterparty = (userRole === 'issuer') ? r.recipientName : r.vendorName;

  var score = 0;
  if (counterparty) score += 2;
  if (r.issueDate) score += 1;
  if (r.total) score += 2;
  if (r.subtotal) score += 1;
  if (r.taxAmount) score += 1;
  if (r.items && r.items.length > 0) score += 2;
  if (r.paymentDueDate) score += 1;
  if (r.invoiceNumber) score += 1;

  // 整合性ペナルティ: 金額式が外税/内税のいずれにも当てはまらない結果は減点
  // (桁ズレや誤拾を含む変種が「ベスト戦略」になることを防ぐ)
  if (hasConsistencyIssue_(r)) score -= 3;

  // 信頼度ペナルティ: 金額の取得元が ¥なし・遠方の行ばかり
  if (r._confidence) {
    var samples = [];
    if (r.total > 0 && r._confidence.total) samples.push(r._confidence.total);
    if (r.subtotal > 0 && r._confidence.subtotal) samples.push(r._confidence.subtotal);
    if (r.taxAmount > 0 && r._confidence.taxAmount) samples.push(r._confidence.taxAmount);
    if (samples.length > 0) {
      var avg = samples.reduce(function(a, b) { return a + b; }, 0) / samples.length;
      if (avg < 50) score -= 1;
    }
  }

  return score;
}

/**
 * 請求番号を抽出
 * @param {string} normalized - 半角変換済みテキスト
 * @param {string} original - 元テキスト
 * @return {string}
 */
function extractInvoiceNumber(normalized, original) {
  // OCR が ASCII ハイフン "-" を Unicode の様々なハイフン類字に誤認することがあるため、
  // 書類番号探索の前に局所的に "-" に正規化する。
  // 対象: U+2010 ‐ HYPHEN, U+2011 ‑ NON-BREAKING HYPHEN, U+2012 ‒ FIGURE DASH,
  //       U+2013 – EN DASH, U+2014 — EM DASH, U+2015 ― HORIZONTAL BAR,
  //       U+2212 − MINUS SIGN, U+30FC ー KATAKANA-HIRAGANA PROLONGED SOUND MARK,
  //       U+FF0D − FULLWIDTH HYPHEN-MINUS
  var hyphenRe = /[‐‑‒–—―−ー－]/g;
  normalized = normalized.replace(hyphenRe, '-');
  original = original.replace(hyphenRe, '-');

  // 許容文字: 英数字 + - _ と、ハイフン/スラッシュ周辺の空白も許容
  // 例: "NX -2026-0011" / "2026 / 0103" / "PO-2026-0234"
  var core = '[A-Za-z0-9](?:[A-Za-z0-9_]|\\s*-\\s*[A-Za-z0-9_])*(?:\\s*\\/\\s*[A-Za-z0-9\\-_]+)?';
  // 優先順位: 具体的なラベル(請求書/注文書/見積書/契約書) → 汎用フォールバック(No./Invoice #)
  // No. パターンを請求書系より下に置くことで「注文書番号: PO-001 / No. INV-001」のような
  // 同居レイアウトで請求書番号として PO-001 を取ってしまう事故を防ぐ
  var patterns = [
    // 請求書系
    new RegExp('請求書番号[:：\\s]*(' + core + ')'),
    new RegExp('請求番号[:：\\s]*(' + core + ')'),
    // Invoice No (英語請求書、decompact後に "INVOICENO" になるケース含む)
    new RegExp('\\bInvoice\\s*No\\.?\\s*(' + core + ')', 'i'),
    new RegExp('I\\s*n\\s*v\\s*o\\s*i\\s*c\\s*e\\s*N\\s*o\\.?[\\s　]*(' + core + ')', 'i'),
    new RegExp('\\bInvoice\\s*#?\\s*(' + core + ')', 'i'),
    // 注文書系
    new RegExp('注文書番号[:：\\s]*(' + core + ')'),
    new RegExp('注文番号[:：\\s]*(' + core + ')'),
    new RegExp('発注番号[:：\\s]*(' + core + ')'),
    new RegExp('\\bPO\\s*Number[:：\\s]*(' + core + ')', 'i'),
    new RegExp('\\bPurchase\\s*Order[:：\\s]*(' + core + ')', 'i'),
    // 見積書系
    new RegExp('見積書番号[:：\\s]*(' + core + ')'),
    new RegExp('見積番号[:：\\s]*(' + core + ')'),
    new RegExp('御見積番号[:：\\s]*(' + core + ')'),
    new RegExp('\\bQuote\\s*No\\.?[:：\\s]*(' + core + ')', 'i'),
    new RegExp('\\bQuotation\\s*No\\.?[:：\\s]*(' + core + ')', 'i'),
    // その他の書類番号
    new RegExp('伝票番号[:：\\s]*(' + core + ')'),
    new RegExp('契約番号[:：\\s]*(' + core + ')'),
    // 汎用フォールバック(具体的なラベルでは取れなかった場合の最後の手段)
    // No. はドット必須(NOTE等の誤マッチ回避)。単語境界は付けない(OCRが
    // "I N V O I C ENo." のように直前の英字と結合するケースに対応するため)
    new RegExp('No\\.\\s*(' + core + ')', 'i'),
    // No (ドットなし) は単語境界ありで安全側に
    new RegExp('\\bNo[:：\\s]+(' + core + ')(?![A-Za-z])', 'i'),
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
      var labelWords = ['請求書番号','請求番号','注文書番号','注文番号','発注番号',
                        '見積書番号','見積番号','御見積番号',
                        '伝票番号','契約番号'];
      var sourceLines = original.split('\n');
      for (var lw = 0; lw < labelWords.length; lw++) {
        var labelRe = buildSpacedLabelRegex(labelWords[lw]);
        for (var ln = 0; ln < sourceLines.length; ln++) {
          if (!labelRe.test(sourceLines[ln])) continue;
          // 次の4行から英数字の番号を探す(空行を挟むレイアウト対応)
          for (var nxt = ln + 1; nxt <= Math.min(ln + 4, sourceLines.length - 1); nxt++) {
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
  // テキスト全体を順に見て、最初に見つかった妥当な日付を採用する
  var kanjiRe = /令和\s*([一二三四五六七八九十]+)\s*年\s*([一二三四五六七八九十]+)\s*月\s*([一二三四五六七八九十]+)\s*日/g;
  var ymd = scanForValidDate_(normalized, original, kanjiRe, function(m) {
    var ky = parseKanjiNumeral(m[1]);
    var km = parseKanjiNumeral(m[2]);
    var kd = parseKanjiNumeral(m[3]);
    if (isNaN(ky) || isNaN(km) || isNaN(kd)) return null;
    return { y: 2018 + ky, m: km, d: kd };
  });
  if (ymd) return formatYmd_(ymd);

  // 和暦(算用): 令和6年1月15日
  ymd = scanForValidDate_(normalized, original, new RegExp(patterns[0].source, 'g'), function(m) {
    return { y: 2018 + parseInt(m[1], 10), m: parseInt(m[2], 10), d: parseInt(m[3], 10) };
  });
  if (ymd) return formatYmd_(ymd);

  // 西暦: 2024年1月15日
  ymd = scanForValidDate_(normalized, original, new RegExp(patterns[1].source, 'g'), function(m) {
    return { y: parseInt(m[1], 10), m: parseInt(m[2], 10), d: parseInt(m[3], 10) };
  });
  if (ymd) return formatYmd_(ymd);

  // スラッシュ・ハイフン・ドット区切り
  ymd = scanForValidDate_(normalized, original, new RegExp(patterns[2].source, 'g'), function(m) {
    return { y: parseInt(m[1], 10), m: parseInt(m[2], 10), d: parseInt(m[3], 10) };
  });
  if (ymd) return formatYmd_(ymd);

  // 請求日・発行日の前後で日付を探す
  var contextPatterns = [
    /(?:請求日|発行日|日付|作成日|年月日)[：:\s]*(.+)/,
  ];
  for (var i = 0; i < contextPatterns.length; i++) {
    var contextMatch = normalized.match(contextPatterns[i]);
    if (contextMatch) {
      var inner = contextMatch[1].trim();
      var innerYmd = extractDate(inner, inner);
      if (innerYmd) return innerYmd;
    }
  }

  return '';
}

/**
 * 日付の妥当性チェック(年: 1900-2099, 月: 1-12, 日: 1-31)
 * うるう年の厳密判定はしない(過剰検証になる)
 */
function isValidYmd_(y, m, d) {
  if (!y || !m || !d) return false;
  if (y < 1900 || y > 2099) return false;
  if (m < 1 || m > 12) return false;
  if (d < 1 || d > 31) return false;
  return true;
}

/**
 * 正規表現を normalized → original の順に走査し、isValidYmd_ を満たす最初の日付を返す
 * @param {RegExp} re - g フラグ必須
 * @param {Function} extract - マッチオブジェクトから {y, m, d} を返す関数
 * @return {{y:number,m:number,d:number}|null}
 */
function scanForValidDate_(normalized, original, re, extract) {
  // 「お支払期限：YYYY年MM月DD日」が「YYYY年MM月DD日」(発行日) より前にあると
  // extractDate が支払期限を発行日として返してしまう。マッチ位置直前 25文字以内に
  // 期限/期日ラベルがあれば、その候補をスキップする(extractPaymentDueDate は別関数なので
  // 期日抽出には影響しない)
  var DUE_LABEL_RE = /(期限|期日|振込日|お支払|支払日|納期|Due|Pay\s*By)/i;
  var sources = [normalized];
  if (original !== normalized) sources.push(original);
  for (var s = 0; s < sources.length; s++) {
    var src = sources[s];
    re.lastIndex = 0;
    var m;
    while ((m = re.exec(src)) !== null) {
      var preCtx = src.substring(Math.max(0, m.index - 25), m.index);
      // 直前の改行までに限定(前行の「期限」が次行の発行日に影響しないように)
      var lastNewline = preCtx.lastIndexOf('\n');
      if (lastNewline >= 0) preCtx = preCtx.substring(lastNewline + 1);
      if (DUE_LABEL_RE.test(preCtx)) continue;
      var ymd = extract(m);
      if (ymd && isValidYmd_(ymd.y, ymd.m, ymd.d)) return ymd;
    }
  }
  return null;
}

function formatYmd_(ymd) {
  return ymd.y + '/' + ymd.m + '/' + ymd.d;
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
  var ippan = '一\\s?般\\s?社\\s?団\\s?法\\s?人';
  var koueki = '公\\s?益\\s?社\\s?団\\s?法\\s?人';
  var ippanZai = '一\\s?般\\s?財\\s?団\\s?法\\s?人';
  var kouekiZai = '公\\s?益\\s?財\\s?団\\s?法\\s?人';
  var npo = '特\\s?定\\s?非\\s?営\\s?利\\s?活\\s?動\\s?法\\s?人|NPO\\s?法\\s?人';
  var gakkou = '学\\s?校\\s?法\\s?人';
  var iryou = '医\\s?療\\s?法\\s?人';
  var corpPat = '(?:' + kabu + '|' + yugen + '|' + godo + '|' +
                ippan + '|' + koueki + '|' + ippanZai + '|' + kouekiZai + '|' +
                npo + '|' + gakkou + '|' + iryou + '|' +
                '（株）|\\(株\\)|K\\.K\\.|G\\.K\\.|Inc\\.|LLC|Corporation|Corp\\.)';

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
  // 書類タイトル(請求書/見積書/...) は名前ではないので候補から除外
  var INVALID_HONOR_TARGETS = /^(御?ご?(請求|見積|注文|契約|納品|領収)書)$/;

  var corpAlt = '株\\s?式\\s?会\\s?社|有\\s?限\\s?会\\s?社|合\\s?同\\s?会\\s?社|' +
                '一\\s?般\\s?社\\s?団\\s?法\\s?人|公\\s?益\\s?社\\s?団\\s?法\\s?人|' +
                '一\\s?般\\s?財\\s?団\\s?法\\s?人|公\\s?益\\s?財\\s?団\\s?法\\s?人|' +
                '特\\s?定\\s?非\\s?営\\s?利\\s?活\\s?動\\s?法\\s?人|NPO\\s?法\\s?人|' +
                '学\\s?校\\s?法\\s?人|医\\s?療\\s?法\\s?人|宗\\s?教\\s?法\\s?人|社\\s?会\\s?福\\s?祉\\s?法\\s?人|' +
                '（株）|\\(株\\)';

  // パターン0: 冒頭(最初の8行)に「会社名行 + 直後に 〒XXX-XXXX で始まる行」がある
  // 請求書レイアウト用 — 御中/様 がなくても請求書の宛名は冒頭にあるという慣例を活用。
  // 「備考 株式会社ラオルカ様より」のような後段の他社名に騙されないよう最優先で試す。
  // 例: "請求書\n株式会社スカイ\n〒150-0031..."
  var headLines = text.split('\n').slice(0, 12);
  var corpLineRe = new RegExp('(?:' + corpAlt + ')');
  for (var hl = 0; hl < headLines.length - 1; hl++) {
    var line = headLines[hl].trim();
    if (!line || !corpLineRe.test(line)) continue;
    // 次の非空行が 〒 で始まるか
    for (var nx = hl + 1; nx < headLines.length; nx++) {
      var nextLine = headLines[nx].trim();
      if (!nextLine) continue;
      if (/^〒\s*\d{3}/.test(nextLine)) {
        // 行内から法人格を含む社名のみを切り出す(「株式会社スカイ」「ABC株式会社」両対応)
        // 行全体が短ければそのまま社名とみなす
        var hv = line.replace(/\s+/g, '').trim();
        if (hv && hv.length <= 40 && !INVALID_HONOR_TARGETS.test(hv)) return hv;
      }
      break;
    }
  }

  // パターン1a: 法人格を含む長めのマッチを先に試す
  // 「合同会社XYZ 様」「株式会社サンプル 御中」が「会社」だけにならないようにする
  // 法人格は OCR で内部にスペースが入っても許容、trailing 部分はスペース1個まで許容
  // (「株式会社 スカイ 御中」→「株式会社スカイ」)
  var corpHonorRe = new RegExp('([^、。\\n]{0,40}?(?:' + corpAlt + ')[^、。\\n]{0,40})\\s*(?:御中|様|さま|殿)', 'g');
  var cm;
  while ((cm = corpHonorRe.exec(text)) !== null) {
    var raw = cm[1];
    // 「備考」「Notes」セクションの会社名は宛先ではないので除外
    if (/(備考|備　考|Notes?|Remarks?|参考|注意)/i.test(raw)) continue;
    var candName = raw.replace(/\s+/g, '').trim();
    if (candName && !INVALID_HONOR_TARGETS.test(candName)) return candName;
  }

  // パターン1b: 敬称直前の会社名/個人名(法人格を持たない屋号や個人名向け)
  // 書類タイトル「御請求書」「ご請求書」は除外する(タイトル直後に「御中」が来る OCR レイアウト対策)
  var honorMatchRe = /([^\s、。\n]+?)\s*(?:御中|様|さま|殿)/g;
  var hm;
  while ((hm = honorMatchRe.exec(text)) !== null) {
    var name = hm[1].trim();
    if (!name || INVALID_HONOR_TARGETS.test(name.replace(/\s/g, ''))) continue;
    return name;
  }

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
 *
 * 英語キーワード(ASCII)は単語境界で判定し、日本語キーワードは部分一致で判定する。
 * これにより "Quantity" を ['Qty'] で誤検出する/しない、のような ASCII 部分一致の
 * 偽陽性を防ぎつつ、日本語の「数量」「金額」などは従来どおり拾える。
 *
 * @param {string} str
 * @param {string[]} words
 * @return {boolean}
 */
function containsAny(str, words) {
  for (var i = 0; i < words.length; i++) {
    var w = words[i];
    if (/^[\x20-\x7e]+$/.test(w)) {
      // ASCII のみ → 単語境界マッチ
      var escaped = w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      if (new RegExp('\\b' + escaped + '\\b', 'i').test(str)) return true;
    } else {
      if (str.indexOf(w) !== -1) return true;
    }
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
  return extractAmountWithConfidence(text, label).value;
}

/**
 * 金額を抽出し、信頼度スコアと共に返す
 *
 * 信頼度 (0-100) の意味:
 *   100 = 同一行 + ¥付き(最も信頼できる)
 *    60 = 同一行(¥なし)
 *    50 = 後続行 + ¥付き
 *    30 = 後続行(¥なし、5行先読みで取得)
 *     0 = 取得失敗
 *
 * @return {{value: number, confidence: number}}
 */
function extractAmountWithConfidence(text, label) {
  // ノイズ(郵便番号・電話番号・口座番号・登録番号・日付)を金額抽出時に除外
  var masked = maskNonAmountPatterns(text);

  var labelVariants = getAmountLabelVariants(label);
  var excludedLines = {};

  for (var v = 0; v < labelVariants.length; v++) {
    var lab = labelVariants[v];
    var labRe = buildSpacedLabelRegex(lab, 'i');
    var lines = masked.split('\n');

    for (var i = 0; i < lines.length; i++) {
      if (excludedLines[i]) continue;
      var labMatch = lines[i].match(labRe);
      if (!labMatch) continue;

      var charBefore = labMatch.index > 0 ? lines[i].charAt(labMatch.index - 1) : '';
      var charAfter = lines[i].charAt(labMatch.index + labMatch[0].length);
      if (/[A-Za-z0-9]/.test(charBefore) || /[A-Za-z0-9]/.test(charAfter)) continue;

      // ラベル直前 10文字以内に「源泉/徴収/Withhold/差引」があれば、別系統(源泉徴収後の振込額や
      // 差引請求額)を拾うのを避ける。例: "源泉徴収税額" 内の "税額"、"差引請求額" 内の "請求額"
      var preCtx = lines[i].substring(Math.max(0, labMatch.index - 10), labMatch.index);
      if (/(源泉|徴収|Withhold|差引)/i.test(preCtx)) continue;

      var afterLabel = lines[i].substring(labMatch.index + labMatch[0].length);

      if (/^[\s:：]*(Rate|rate|RATE|率)/.test(afterLabel)) continue;
      if (lab === '税' && /^[込別率]/.test(afterLabel)) continue;

      excludedLines[i] = true;
      var nums = collectAmountCandidates(afterLabel);
      var source = 'sameLine';

      if (!hasYenValue(nums)) {
        var fallbackNums = nums;
        var fallbackSource = nums.length > 0 ? 'sameLine' : null;
        // window: 最大6行先まで見るが、マスク後に空になる行(日付/電話)は budget を消費しない。
        // productive な行(候補ありの行)が 3 を越えたら打ち切り。
        // これにより「ご請求金額\nお支払期限:...\n2026年04月01日\n請求書番号: 20260403-2\n¥12,100」
        // のような間に挟まる日付行を飛ばして本来の ¥12,100 行まで到達できる。
        var productive = 0;
        for (var k = i + 1; k < Math.min(i + 7, lines.length) && productive < 3; k++) {
          var nextLine = lines[k];
          if (!nextLine.trim()) continue;
          if (isAnotherAmountLabel(nextLine)) break;
          var nextNums = collectAmountCandidates(nextLine);
          if (nextNums.length === 0) continue;
          productive++;
          if (hasYenValue(nextNums)) { nums = nextNums; source = 'nextLine'; break; }
          if (!hasYenValue(fallbackNums)) { fallbackNums = nextNums; fallbackSource = 'nextLine'; }
        }
        if (!hasYenValue(nums) && fallbackNums.length > 0) {
          nums = fallbackNums;
          source = fallbackSource || source;
        }
      }

      if (nums.length > 0) {
        var value = pickBestAmount(nums);
        var withYen = hasYenValue(nums);
        var confidence;
        if (source === 'sameLine' && withYen) confidence = 100;
        else if (source === 'sameLine') confidence = 60;
        else if (source === 'nextLine' && withYen) confidence = 50;
        else confidence = 30;
        return { value: value, confidence: confidence };
      }
    }
  }

  return { value: 0, confidence: 0 };
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
  // ¥記号(前)・円/JPYサフィックス(後)・%付き(税率)のいずれもキャプチャ
  // 円/JPY サフィックスがあれば「金額」と確信できるので hasYen 相当の信頼として扱う
  var re = /([¥￥])?\s*([\d,]+)\s*(円|JPY|jpy)?(\s*[%％])?/g;
  var m;
  while ((m = re.exec(line)) !== null) {
    if (m[4]) continue;  // % 付き → 税率
    var n = parseNumber(m[2]);
    // 1億円以上は実質的に書類番号(8桁以上の連続数字)等のノイズなので除外。
    // ¥付きでも除外する(¥20260403 のような OCR 連結ケース対策)
    if (n > 0 && n < 100000000) amounts.push({ value: n, hasYen: !!m[1] || !!m[3] });
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
  return /(小計|消費税|税額|合計|総額|総計|ご請求|御請求|請求金額|送料|運送費|配送料|振込手数料|手数料|割引|値引|Subtotal|Sub\s*Total|Grand\s*Total|Amount\s*Due|Total\b|Tax\b|VAT|GST|Shipping|Handling|Discount)/i.test(line);
}

/**
 * 多列ラベル(小計 消費税 合計 / 合計 税抜 消費税 総額 等)レイアウトから
 * subtotal / taxAmount / total を一括抽出する。
 *
 * 例: "小計 消費税 合計金額\n100,000 10,000 110,000"
 *     → { subtotal: 100000, taxAmount: 10000, total: 110000 }
 *
 * 通常の extractAmount は「ラベル毎に next-line を見る」ため、3つのラベルが
 * 同じ next-line を見て第1値を採用してしまい subtotal/tax/total が混同される。
 * 本関数は 1行のラベル並びと 1行の値並びを順番に対応させて解決する。
 *
 * @param {string} text
 * @return {{subtotal:number, taxAmount:number, total:number}|null}
 */
function extractAmountTriplet_(text) {
  if (!text) return null;
  var lines = text.split('\n');
  // ラベル単位で順序を保ったまま検出する正規表現
  var labelTokenRe = /(小計|税抜合計|税抜金額|税抜|消費税額等|消費税額|消費税等|消費税|税額|税込合計|税込|総額|総計|合計金額|合計|ご請求金額|御請求金額|ご請求額|請求金額)/g;

  function collectLabels(line) {
    var ls = [];
    labelTokenRe.lastIndex = 0;
    var lm;
    while ((lm = labelTokenRe.exec(line)) !== null) {
      // (税込) / （税抜) のように括弧内の修飾子はラベルではないので除外
      var prevCh = lm.index > 0 ? line.charAt(lm.index - 1) : '';
      if ((lm[1] === '税込' || lm[1] === '税抜') && (prevCh === '(' || prevCh === '（')) continue;
      ls.push(lm[1]);
    }
    return ls;
  }

  for (var i = 0; i < lines.length - 1; i++) {
    var labels = collectLabels(lines[i]);
    // 前行が「ラベルのみの短い行」(例: 単独 "小計")なら現行のラベルと結合する。
    // これにより "小計\n消費税 合計金額\n100,000 10,000 110,000" のような
    // 多行ラベルレイアウトも 1 セットとして扱える。
    if (i > 0) {
      var prevLabels = collectLabels(lines[i - 1]);
      // 前行を結合する条件: ラベルがあり、かつ短く、かつ値(数字)を含まない。
      // 「消費税額 ¥410」のような自己完結 label+value 行は結合しない(別個の集計行のため)
      var prevHasValue = collectAmountCandidates(lines[i - 1]).length > 0;
      if (prevLabels.length > 0 && !prevHasValue && lines[i - 1].replace(/[\s　]/g, '').length <= 10) {
        labels = prevLabels.concat(labels);
      }
    }
    if (labels.length < 2) continue;

    // 値の探索: まず同一行内(ラベル末尾位置以降)を試し、なければ次行を見る
    var valueLines = [];
    // 同一行: 最後のラベル末尾位置以降の文字列
    labelTokenRe.lastIndex = 0;
    var lastLabelEnd = 0;
    var ltm;
    while ((ltm = labelTokenRe.exec(lines[i])) !== null) {
      lastLabelEnd = ltm.index + ltm[0].length;
    }
    var sameLineRest = lines[i].substring(lastLabelEnd);
    if (sameLineRest.trim()) valueLines.push(sameLineRest);
    // 次行(最大2行先)
    for (var nj = i + 1; nj < Math.min(i + 3, lines.length); nj++) {
      if (lines[nj].trim()) { valueLines.push(lines[nj]); break; }
    }

    var matched = false;
    for (var vi = 0; vi < valueLines.length; vi++) {
      var nums = collectAmountCandidates(valueLines[vi]);
      if (nums.length < 2) continue;

      // ラベルが値より多い場合は「先頭ラベルが行ヘッダ(『合計:』等)」とみなして
      // 末尾から N 個のラベルだけを値と対応させる。
      // 例: labels=['合計','税抜','消費税','総額'], nums=3個 → ['税抜','消費税','総額'] と対応
      var pairCount = Math.min(labels.length, nums.length);
      var startLabel = labels.length - pairCount;

      var result = { subtotal: 0, taxAmount: 0, total: 0 };
      for (var k = 0; k < pairCount; k++) {
        var lab = labels[startLabel + k];
        var val = nums[k].value;
        if (lab === '小計' || lab === '税抜合計' || lab === '税抜金額' || lab === '税抜') {
          if (!result.subtotal) result.subtotal = val;
        } else if (lab === '消費税額等' || lab === '消費税額' || lab === '消費税等' || lab === '消費税' || lab === '税額') {
          if (!result.taxAmount) result.taxAmount = val;
        } else if (lab === '合計' || lab === '合計金額' || lab === '総額' || lab === '総計' ||
                   lab === '税込合計' || lab === '税込' ||
                   lab === 'ご請求金額' || lab === '御請求金額' || lab === 'ご請求額' || lab === '請求金額') {
          if (!result.total) result.total = val;
        }
      }
      // 少なくとも 2 値が取れていれば triplet として採用
      var filled = (result.subtotal ? 1 : 0) + (result.taxAmount ? 1 : 0) + (result.total ? 1 : 0);
      if (filled >= 2) { matched = true; return result; }
    }
    if (matched) break;
  }
  return null;
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

  // 各カテゴリのキーワードを集計し、最多マッチのカテゴリを採用
  // (順序依存をなくし、「請求書 + 注文番号」のような同居レイアウトでも
  //  メイン書類の方が正しく選ばれるようにする)
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

  var scores = {};
  var totalMatches = 0;
  for (var c = 0; c < categories.length; c++) {
    var type = categories[c][0];
    var keywords = categories[c][1];
    var hits = 0;
    for (var k = 0; k < keywords.length; k++) {
      if (matches(buildSpacedLabelRegex(keywords[k], 'i'))) hits++;
    }
    scores[type] = hits;
    totalMatches += hits;
  }

  // 「第◯条」は契約書の強いシグナル(タイトル無しの契約書もある)
  if (/第\s*[一二三四五六七八九十0-9]+\s*条/.test(text)) {
    scores.contract = (scores.contract || 0) + 2;
    totalMatches += 2;
  }

  if (totalMatches > 0) {
    var bestType = null;
    var bestHits = 0;
    // categories の順番(quote/order/delivery/receipt/invoice/...)を同点時のタイブレーカーに使う
    // → invoice が中位で、より具体的な種別(quote/order/delivery/receipt)を優先する
    for (var ci = 0; ci < categories.length; ci++) {
      var t = categories[ci][0];
      if (scores[t] > bestHits) {
        bestHits = scores[t];
        bestType = t;
      }
    }
    // contract は categories ループ後にも比較
    if (scores.contract > bestHits) bestType = 'contract';
    if (bestType) return bestType;
  }

  // テーブルっぽい構造を検出(タブ区切りや連続スペースの行が複数)
  // 閾値5: 請求書の明細部分(3行程度のテーブル)を 'table' と誤判定して
  // shouldCallGemini_ の対象外になる事故を防ぐ
  var lines = text.split('\n');
  var tableLineCount = 0;
  for (var j = 0; j < lines.length; j++) {
    if (/\t/.test(lines[j]) || /\S\s{2,}\S/.test(lines[j])) {
      tableLineCount++;
    }
  }
  if (tableLineCount >= 5) return 'table';

  return 'text';
}
