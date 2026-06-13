/**
 * 請求書 一括作成＆メール下書きツール（ママゴトラボ）
 *
 * スプレッドシートの請求データから請求書PDFを一括生成し、
 * 取引先ごとにGmailの下書きを作成します（添付：請求書PDF）。
 *
 * 安全のため、既定は「下書き作成」です。
 * 内容を目視で確認してから手動で送信できます。
 * まとめて自動送信したい場合のみメニューの「請求書を送信（自動）」を使います。
 *
 * 使い方の詳細は README.md を参照してください。
 */

// ===== 設定 =====
const SHEET_DATA = '請求データ';   // 請求明細を入力するシート
const SHEET_CONFIG = '設定';       // 自社情報を入力するシート
const TAX_RATE = 0.10;             // 消費税率（10%）

// 請求データシートの列番号（1始まり）
const COL = {
  invoiceNo: 1,  // 請求書番号（同じ番号の行は1枚の請求書にまとめます）
  issueDate: 2,  // 発行日
  client: 3,     // 取引先名
  email: 4,      // 宛先メールアドレス
  item: 5,       // 品目
  qty: 6,        // 数量
  unitPrice: 7,  // 単価（税抜）
  status: 8,     // ステータス（自動更新）
  processedAt: 9 // 処理日時（自動更新）
};

/**
 * スプレッドシートを開いたときにメニューを追加する
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('請求書ツール')
    .addItem('① 初期セットアップ（サンプル作成）', 'initSheets')
    .addSeparator()
    .addItem('② 請求書を作成（Gmail下書き）', 'createInvoiceDrafts')
    .addItem('③ 請求書を送信（自動）', 'sendInvoices')
    .addToUi();
}

/**
 * 設定シートと請求データシートを作成し、サンプルを入れる。
 * 初めて使う人がすぐ動かせるようにするための初期化処理。
 */
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 設定シート
  let config = ss.getSheetByName(SHEET_CONFIG);
  if (!config) {
    config = ss.insertSheet(SHEET_CONFIG);
    const rows = [
      ['項目', '値'],
      ['自社名', '株式会社サンプル'],
      ['住所', '東京都〇〇区〇〇1-2-3'],
      ['電話番号', '03-0000-0000'],
      ['メールアドレス', 'info@example.com'],
      ['振込先', '〇〇銀行 〇〇支店 普通 1234567 カ）サンプル'],
      ['支払期限（発行日から日数）', '30']
    ];
    config.getRange(1, 1, rows.length, 2).setValues(rows);
    config.getRange(1, 1, 1, 2).setFontWeight('bold');
    config.setColumnWidth(1, 200);
    config.setColumnWidth(2, 320);
  }

  // 請求データシート
  let data = ss.getSheetByName(SHEET_DATA);
  if (!data) {
    data = ss.insertSheet(SHEET_DATA);
    const header = ['請求書番号', '発行日', '取引先名', '宛先メール', '品目', '数量', '単価', 'ステータス', '処理日時'];
    const sample = [
      ['INV-001', '2026-06-30', '取引先A株式会社', 'a@example.com', 'Webサイト保守（6月分）', 1, 30000, '', ''],
      ['INV-001', '2026-06-30', '取引先A株式会社', 'a@example.com', 'サーバー利用料', 1, 5000, '', ''],
      ['INV-002', '2026-06-30', '取引先B商店', 'b@example.com', '記事執筆（5本）', 5, 8000, '', '']
    ];
    data.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    data.getRange(2, 1, sample.length, header.length).setValues(sample);
    data.setColumnWidth(COL.item, 220);
    data.setColumnWidth(COL.email, 180);
  }

  SpreadsheetApp.getUi().alert(
    'セットアップ完了',
    '「設定」シートに自社情報を、「請求データ」シートに請求内容を入力してください。\n' +
    'サンプルが入っているので、書き換えて試せます。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 請求書PDFを作り、取引先ごとにGmailの下書きを作成する（送信はしない）
 */
function createInvoiceDrafts() {
  processInvoices_(false);
}

/**
 * 請求書PDFを作り、取引先ごとにメールを自動送信する
 */
function sendInvoices() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    '確認',
    '請求書を取引先へ自動送信します。よろしいですか？\n（不安な場合は「請求書を作成（Gmail下書き）」をおすすめします）',
    ui.ButtonSet.OK_CANCEL
  );
  if (res !== ui.Button.OK) return;
  processInvoices_(true);
}

/**
 * 本体処理。請求データを請求書単位にまとめ、PDF化して下書き作成または送信する。
 * @param {boolean} send true=送信 / false=下書き作成
 */
function processInvoices_(send) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(SHEET_DATA);
  const ui = SpreadsheetApp.getUi();

  if (!dataSheet || !ss.getSheetByName(SHEET_CONFIG)) {
    ui.alert('先に「初期セットアップ」を実行してください。');
    return;
  }

  const config = readConfig_();
  const invoices = groupByInvoice_(dataSheet);

  // --- 処理前バリデーション（コストの高い処理を始める前にまとめて検証）---
  const errors = validateInvoices_(invoices, config);
  if (errors.length > 0) {
    ui.alert('入力エラー', '処理を中止しました。以下を修正してください。\n\n' + errors.join('\n'), ui.ButtonSet.OK);
    return;
  }

  // --- 請求書ごとに生成 ---
  let ok = 0;
  const now = new Date();
  Object.keys(invoices).forEach(function (no) {
    const inv = invoices[no];
    const html = buildInvoiceHtml_(config, inv);
    const pdf = Utilities.newBlob(html, 'text/html', '請求書_' + no + '.html')
      .getAs('application/pdf')
      .setName('請求書_' + no + '.pdf');

    const subject = '【請求書】' + config['自社名'] + ' / ' + inv.client + ' 様';
    const body =
      inv.client + ' 御中\n\n' +
      'いつもお世話になっております。' + config['自社名'] + 'です。\n' +
      '請求書（' + no + '）をお送りいたします。ご査収のほどよろしくお願いいたします。\n\n' +
      '------------------------------\n' +
      config['自社名'] + '\n' + config['住所'] + '\n' +
      'TEL: ' + config['電話番号'] + '\n' +
      '------------------------------';

    const options = { attachments: [pdf], name: config['自社名'] };
    if (send) {
      GmailApp.sendEmail(inv.email, subject, body, options);
    } else {
      GmailApp.createDraft(inv.email, subject, body, options);
    }

    // ステータス更新（請求書に含まれる全行に記録）
    const label = send ? '送信済み' : '下書き作成済み';
    inv.rows.forEach(function (r) {
      dataSheet.getRange(r, COL.status).setValue(label);
      dataSheet.getRange(r, COL.processedAt).setValue(now);
    });
    ok++;
  });

  ui.alert('完了', ok + ' 件の請求書を' + (send ? '送信' : '下書き作成') + 'しました。', ui.ButtonSet.OK);
}

/**
 * 設定シートを {項目: 値} の辞書として読み込む
 */
function readConfig_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const config = {};
  values.forEach(function (row) {
    if (row[0]) config[String(row[0]).trim()] = row[1];
  });
  return config;
}

/**
 * 請求データを請求書番号ごとにまとめる
 * @return {Object} { 請求書番号: {client, email, issueDate, items:[], rows:[行番号]} }
 */
function groupByInvoice_(sheet) {
  const last = sheet.getLastRow();
  const invoices = {};
  if (last < 2) return invoices;

  const values = sheet.getRange(2, 1, last - 1, COL.processedAt).getValues();
  values.forEach(function (row, i) {
    const no = String(row[COL.invoiceNo - 1]).trim();
    if (!no) return;
    if (!invoices[no]) {
      invoices[no] = {
        no: no,
        issueDate: row[COL.issueDate - 1],
        client: String(row[COL.client - 1]).trim(),
        email: String(row[COL.email - 1]).trim(),
        items: [],
        rows: []
      };
    }
    invoices[no].items.push({
      item: String(row[COL.item - 1]).trim(),
      qty: Number(row[COL.qty - 1]) || 0,
      unitPrice: Number(row[COL.unitPrice - 1]) || 0
    });
    invoices[no].rows.push(i + 2); // 実際の行番号
  });
  return invoices;
}

/**
 * 全請求書の入力チェック。エラーメッセージの配列を返す。
 */
function validateInvoices_(invoices, config) {
  const errors = [];
  const required = ['自社名', '住所', '振込先'];
  required.forEach(function (key) {
    if (!config[key]) errors.push('・「設定」シートの『' + key + '』が空です');
  });

  const emailRe = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;
  Object.keys(invoices).forEach(function (no) {
    const inv = invoices[no];
    if (!inv.client) errors.push('・' + no + '：取引先名が空です');
    if (!emailRe.test(inv.email)) errors.push('・' + no + '：宛先メールが不正です（' + inv.email + '）');
    inv.items.forEach(function (it) {
      if (!it.item) errors.push('・' + no + '：品目が空の行があります');
      if (it.qty <= 0) errors.push('・' + no + '：数量が0以下の行があります（' + it.item + '）');
    });
  });
  return errors;
}

/**
 * 請求書のHTMLを組み立てる（PDF化のもとになる）
 */
function buildInvoiceHtml_(config, inv) {
  let subtotal = 0;
  let rows = '';
  inv.items.forEach(function (it) {
    const amount = it.qty * it.unitPrice;
    subtotal += amount;
    rows +=
      '<tr>' +
      '<td>' + escapeHtml_(it.item) + '</td>' +
      '<td style="text-align:right">' + it.qty + '</td>' +
      '<td style="text-align:right">' + formatYen_(it.unitPrice) + '</td>' +
      '<td style="text-align:right">' + formatYen_(amount) + '</td>' +
      '</tr>';
  });
  const tax = Math.floor(subtotal * TAX_RATE);
  const total = subtotal + tax;
  const issue = formatDate_(inv.issueDate);
  const dueDays = Number(config['支払期限（発行日から日数）']) || 30;
  const due = formatDate_(new Date(new Date(inv.issueDate).getTime() + dueDays * 86400000));

  return '' +
    '<html><head><meta charset="utf-8"><style>' +
    'body{font-family:sans-serif;color:#333;padding:40px;}' +
    'h1{font-size:24px;border-bottom:2px solid #5a8060;padding-bottom:8px;}' +
    '.meta{text-align:right;font-size:12px;color:#666;}' +
    '.total{font-size:20px;font-weight:bold;color:#5a8060;}' +
    'table{width:100%;border-collapse:collapse;margin-top:20px;}' +
    'th,td{border:1px solid #ccc;padding:8px;font-size:13px;}' +
    'th{background:#f0f4f0;}' +
    '.box{margin-top:24px;font-size:13px;line-height:1.8;}' +
    '</style></head><body>' +
    '<div class="meta">請求書番号：' + escapeHtml_(inv.no) + '<br>発行日：' + issue + '</div>' +
    '<h1>請 求 書</h1>' +
    '<p><b>' + escapeHtml_(inv.client) + ' 御中</b></p>' +
    '<p>下記のとおりご請求申し上げます。</p>' +
    '<p class="total">ご請求金額：' + formatYen_(total) + '（税込）</p>' +
    '<p>お支払期限：' + due + '</p>' +
    '<table>' +
    '<tr><th>品目</th><th>数量</th><th>単価</th><th>金額</th></tr>' +
    rows +
    '<tr><td colspan="3" style="text-align:right">小計</td><td style="text-align:right">' + formatYen_(subtotal) + '</td></tr>' +
    '<tr><td colspan="3" style="text-align:right">消費税（' + (TAX_RATE * 100) + '%）</td><td style="text-align:right">' + formatYen_(tax) + '</td></tr>' +
    '<tr><td colspan="3" style="text-align:right"><b>合計</b></td><td style="text-align:right"><b>' + formatYen_(total) + '</b></td></tr>' +
    '</table>' +
    '<div class="box">' +
    '<b>お振込先</b><br>' + escapeHtml_(String(config['振込先'])) + '<br><br>' +
    '<b>' + escapeHtml_(String(config['自社名'])) + '</b><br>' +
    escapeHtml_(String(config['住所'] || '')) + '<br>' +
    'TEL：' + escapeHtml_(String(config['電話番号'] || '')) + '　' +
    'Mail：' + escapeHtml_(String(config['メールアドレス'] || '')) +
    '</div>' +
    '</body></html>';
}

// ===== 小さなユーティリティ =====

function formatYen_(n) {
  return '¥' + Number(n).toLocaleString('ja-JP');
}

function formatDate_(d) {
  const date = (d instanceof Date) ? d : new Date(d);
  return Utilities.formatDate(date, 'JST', 'yyyy年M月d日');
}

function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}
