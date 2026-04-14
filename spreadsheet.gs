/**
 * スプレッドシート作成モジュール
 * パースされた請求書データからGoogle Spreadsheetを作成する
 */

/**
 * 取引台帳に1エントリを追記
 * @param {Object} entry - parseLedgerEntry()の戻り値
 */
function appendToLedger(entry) {
  if (!CFG.ledger.spreadsheetId) {
    throw new Error('CFG.ledger.spreadsheetId が未設定です。createLedgerSpreadsheet() を実行してください');
  }

  var ss = SpreadsheetApp.openById(CFG.ledger.spreadsheetId);
  var sheet = ss.getSheetByName(CFG.ledger.sheetName) || ss.getActiveSheet();

  var row = [
    Utilities.formatDate(entry.processedAt, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'),
    entry.fileName,
    entry.fileLink,
    entry.docType,
    entry.vendor || '',
    entry.issueDate || '',
    entry.docNumber || '',
    entry.total || '',
    entry.subtotal || '',
    entry.tax || '',
    entry.paymentDueDate || '',
    entry.contentSummary || '',
    entry.rawText || '',
    entry.status,
  ];

  sheet.appendRow(row);

  // 金額列（合計8, 小計9, 消費税10）をカンマ区切り表示に
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 8, 1, 3).setNumberFormat('#,##0');

  console.log('台帳に追記: ' + entry.fileName +
              ' (' + entry.docType + ' / ' + (entry.vendor || '取引先未検出') +
              ' / ' + (entry.total || '金額未検出') + ')');
}

/**
 * 請求書データからスプレッドシートを作成（後方互換ラッパー）
 * @param {Object} invoice - parseInvoice()の戻り値
 * @param {string} originalFileName - 元PDFのファイル名
 * @return {string} 作成されたスプレッドシートのID
 */
function createInvoiceSheet(invoice, originalFileName) {
  return createBusinessDocSheet(invoice, originalFileName, '請求書');
}

/**
 * ビジネス書類（請求書/領収書/見積書/注文書/納品書）からスプレッドシートを作成
 * @param {Object} invoice - parseInvoice()の戻り値
 * @param {string} originalFileName - 元ファイル名
 * @param {string} docType - 文書種別ラベル ('請求書' | '領収書' | '見積書' | '注文書' | '納品書')
 * @return {string} 作成されたスプレッドシートのID
 */
function createBusinessDocSheet(invoice, originalFileName, docType) {
  var baseName = originalFileName.replace(/\.[^.]+$/, '');
  var prefix = docType + '_';
  if (baseName.indexOf(prefix) === 0) {
    baseName = baseName.substring(prefix.length);
  }
  var parts = [docType, baseName];
  if (invoice.invoiceNumber) parts.push(invoice.invoiceNumber);
  var sheetName = parts.join('_');

  var ss = SpreadsheetApp.create(sheetName);
  var sheet = ss.getActiveSheet();
  sheet.setName(docType + 'データ');

  // ===== ヘッダー部 =====
  var row = 1;
  sheet.getRange(row, 1).setValue(docType + '番号').setFontWeight('bold');
  sheet.getRange(row, 2).setValue(invoice.invoiceNumber || '（未検出）');
  row++;

  sheet.getRange(row, 1).setValue('発行日').setFontWeight('bold');
  sheet.getRange(row, 2).setValue(invoice.issueDate || '（未検出）');
  row++;

  sheet.getRange(row, 1).setValue('発行元').setFontWeight('bold');
  sheet.getRange(row, 2).setValue(invoice.vendorName || '（未検出）');
  row++;

  sheet.getRange(row, 1).setValue('宛先').setFontWeight('bold');
  sheet.getRange(row, 2).setValue(invoice.recipientName || '（未検出）');
  row += 2; // 1行空ける

  // ===== 明細テーブル =====
  var tableHeaderRow = row;
  var headers = ['品名', '数量', '単価', '金額'];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#4a86c8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  row++;

  // 明細行
  var items = invoice.items || [];
  if (items.length > 0) {
    var itemData = items.map(function(item) {
      return [item.name, item.quantity, item.unitPrice, item.amount];
    });
    sheet.getRange(row, 1, itemData.length, 4).setValues(itemData);

    // 数値列のフォーマット
    sheet.getRange(row, 2, itemData.length, 1).setNumberFormat('#,##0');     // 数量
    sheet.getRange(row, 3, itemData.length, 1).setNumberFormat('#,##0');     // 単価
    sheet.getRange(row, 4, itemData.length, 1).setNumberFormat('#,##0');     // 金額

    row += itemData.length;
  } else {
    sheet.getRange(row, 1).setValue('（明細データ未検出）');
    row++;
  }

  row++; // 1行空ける

  // ===== 合計部 =====
  sheet.getRange(row, 3).setValue('小計').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange(row, 4).setValue(invoice.subtotal).setNumberFormat('#,##0');
  row++;

  sheet.getRange(row, 3).setValue('消費税').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange(row, 4).setValue(invoice.taxAmount).setNumberFormat('#,##0');
  row++;

  sheet.getRange(row, 3).setValue('合計').setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange(row, 4).setValue(invoice.total).setNumberFormat('#,##0')
    .setFontWeight('bold')
    .setFontSize(12);
  row++;

  // ===== 罫線設定 =====
  var tableEndRow = tableHeaderRow + (items.length > 0 ? items.length : 1);
  sheet.getRange(tableHeaderRow, 1, tableEndRow - tableHeaderRow + 1, 4)
    .setBorder(true, true, true, true, true, true);

  // ===== 列幅調整 =====
  sheet.setColumnWidth(1, 250); // 品名
  sheet.setColumnWidth(2, 80);  // 数量
  sheet.setColumnWidth(3, 120); // 単価
  sheet.setColumnWidth(4, 120); // 金額

  // ===== OCR原文シート =====
  var rawSheet = ss.insertSheet('OCR原文');
  rawSheet.getRange(1, 1).setValue('以下はOCRで抽出された原文テキストです：');
  rawSheet.getRange(3, 1).setValue(invoice.rawText);
  rawSheet.setColumnWidth(1, 600);

  // 出力フォルダへ移動
  if (CFG.folders.output) {
    var outputFolder = DriveApp.getFolderById(CFG.folders.output);
    DriveApp.getFileById(ss.getId()).moveTo(outputFolder);
  }

  console.log('スプレッドシート作成完了: ' + sheetName + ' (ID: ' + ss.getId() + ')');
  return ss.getId();
}

/**
 * 汎用テーブルデータからスプレッドシートを作成
 * （請求書以外の表形式PDF向け）
 * @param {string[][]} rows - 行データの2次元配列
 * @param {string} title - スプレッドシートのタイトル
 * @return {string} 作成されたスプレッドシートのID
 */
function createGenericTableSheet(rows, title) {
  var ss = SpreadsheetApp.create(title);
  var sheet = ss.getActiveSheet();
  sheet.setName('データ');

  if (rows.length > 0) {
    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

    // 1行目をヘッダーとして書式設定
    sheet.getRange(1, 1, 1, rows[0].length)
      .setFontWeight('bold')
      .setBackground('#4a86c8')
      .setFontColor('#ffffff');
  }

  // 出力フォルダへ移動
  if (CFG.folders.output) {
    var outputFolder = DriveApp.getFolderById(CFG.folders.output);
    DriveApp.getFileById(ss.getId()).moveTo(outputFolder);
  }

  return ss.getId();
}
