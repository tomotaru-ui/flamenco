/**
 * フラメンコ公演 参加登録システム バックエンド
 * このファイルは okinawakarate2026@gmail.com で作成するスプレッドシートに
 * 紐づく Google Apps Script プロジェクトに貼り付けて使用します。
 * デプロイ・初期設定の手順は README.md を参照してください。
 */

var SHEET_NAME = '登録';
var CAPACITY = 220;
var ADULT_PRICE = 2500;
var CHILD_PRICE = 1000;
var HEADERS = ['登録番号', 'タイムスタンプ', '代表者氏名', 'フリガナ', 'メールアドレス', '電話番号',
  '大人人数', '子供人数', '合計金額', '支払方法', '支払状況', '入場チェック', '備考'];

function getProp_(key, fallback) {
  var v = PropertiesService.getScriptProperties().getProperty(key);
  return v ? v : fallback;
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 初回のみ手動で実行してください(シート作成 + 権限の許可ダイアログ表示のため)。
 */
function setup() {
  getSheet_();
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action || 'register';
    if (action === 'register') return handleRegister_(payload);
    if (action === 'lookup') return jsonOut_(handleLookup_(payload));
    if (action === 'markPaid') return jsonOut_(handleMarkPaid_(payload));
    if (action === 'checkin') return jsonOut_(handleCheckin_(payload));
    if (action === 'list') return jsonOut_(handleList_(payload));
    if (action === 'adminStats') return jsonOut_(handleAdminStats_(payload));
    if (action === 'reset') return jsonOut_(handleReset_(payload));
    if (action === 'sendQr') return jsonOut_(handleSendQr_(payload));
    return jsonOut_({ ok: false, error: '不明なアクションです' });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  var action = e.parameter.action;
  if (action === 'stats') return jsonOut_(getStats_());
  return ContentService.createTextOutput('Flamenco registration API is running.');
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getStats_() {
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  var total = 0;
  for (var i = 1; i < data.length; i++) {
    total += Number(data[i][6] || 0) + Number(data[i][7] || 0);
  }
  return { ok: true, capacity: CAPACITY, registered: total, remaining: Math.max(CAPACITY - total, 0) };
}

function handleRegister_(payload) {
  var name = String(payload.name || '').trim();
  var kana = String(payload.kana || '').trim();
  var email = String(payload.email || '').trim();
  var tel = String(payload.tel || '').trim();
  var adults = Math.max(0, parseInt(payload.adults, 10) || 0);
  var children = Math.max(0, parseInt(payload.children, 10) || 0);
  var paymentMethod = payload.paymentMethod === 'cash' ? 'cash' : 'transfer';

  if (!name || !email || (adults + children) < 1) {
    return jsonOut_({ ok: false, error: 'お名前・メールアドレス・人数をご確認のうえ、もう一度お送りください。' });
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return jsonOut_({ ok: false, error: 'メールアドレスの形式が正しくありません。' });
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_();
    var stats = getStats_();
    if (stats.remaining < (adults + children)) {
      return jsonOut_({
        ok: false,
        error: '大変申し訳ございません。現在お申込みいただける残席は ' + stats.remaining + ' 席のため、' +
          (adults + children) + ' 名でのお申込みを承ることができませんでした。お手数ですが主催者までお問い合わせください。'
      });
    }

    var seq = sheet.getLastRow(); // ヘッダー行を除いた既存件数 + 1 と一致する
    var regNumber = 'TF-' + ('0000' + seq).slice(-4);
    var amount = adults * ADULT_PRICE + children * CHILD_PRICE;
    var timestamp = new Date();
    var paymentMethodLabel = paymentMethod === 'cash' ? '当日現金' : '口座振込';

    sheet.appendRow([regNumber, timestamp, name, kana, email, tel, adults, children, amount,
      paymentMethodLabel, '未確認', '未', '']);

    var token = buildQrToken_(regNumber);
    var qrBlob = fetchQrImage_(token);
    sendConfirmationEmail_({
      regNumber: regNumber, name: name, email: email, adults: adults, children: children,
      amount: amount, paymentMethod: paymentMethod, qrBlob: qrBlob
    });

    return jsonOut_({ ok: true, regNumber: regNumber });
  } finally {
    lock.releaseLock();
  }
}

function buildQrToken_(regNumber) {
  var secret = getProp_('QR_SECRET', 'change-me-secret');
  var raw = Utilities.computeHmacSha256Signature(regNumber, secret);
  var hex = raw.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
  return regNumber + '.' + hex.substring(0, 8);
}

function verifyQrToken_(token) {
  var parts = String(token || '').split('.');
  if (parts.length !== 2) return null;
  var regNumber = parts[0];
  var sig = parts[1];
  var expected = buildQrToken_(regNumber).split('.')[1];
  if (sig !== expected) return null;
  return regNumber;
}

function fetchQrImage_(token) {
  var url = 'https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=' + encodeURIComponent(token);
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  return resp.getBlob().setName('qrcode.png');
}

function findRowByRegNumber_(sheet, regNumber) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === regNumber) return i + 1; // 1-indexed row number
  }
  return -1;
}

function requirePin_(payload) {
  var pin = getProp_('STAFF_PIN', '');
  if (!pin || String(payload.pin) !== String(pin)) {
    throw new Error('PINが正しくありません');
  }
}

function requireAdminPin_(payload) {
  var pin = getProp_('ADMIN_PIN', '');
  if (!pin || String(payload.pin) !== String(pin)) {
    throw new Error('管理者PINが正しくありません');
  }
}

function handleAdminStats_(payload) {
  requireAdminPin_(payload);
  var sheet = getSheet_();
  var stats = getStats_();
  return {
    ok: true,
    registrationCount: Math.max(sheet.getLastRow() - 1, 0),
    registeredSeats: stats.registered,
    capacity: stats.capacity,
    remaining: stats.remaining
  };
}

/**
 * 現在の登録データを別シートにアーカイブしてから、
 * 「登録」シートをヘッダーのみの状態に戻す(次回登録は TF-0001 から再開する)。
 */
function handleReset_(payload) {
  requireAdminPin_(payload);
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = getSheet_();
    var lastRow = sheet.getLastRow();
    var archivedCount = Math.max(lastRow - 1, 0);
    var archiveName = null;

    if (archivedCount > 0) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var label = String(payload.eventLabel || '').trim().replace(/[^\w぀-ヿ一-鿿-]/g, '');
      var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
      archiveName = ('archive_' + (label ? label + '_' : '') + stamp).substring(0, 90);
      var archiveSheet = sheet.copyTo(ss);
      archiveSheet.setName(archiveName);
      sheet.deleteRows(2, lastRow - 1);
    }

    return { ok: true, archivedCount: archivedCount, archiveName: archiveName };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 紙チケット・代理登録など、スプレッドシートに直接記入した行に対して、
 * QRコードを生成しなおして確認メールを送信する。
 */
function handleSendQr_(payload) {
  requireAdminPin_(payload);
  var regNumber = String(payload.regNumber || '').trim();
  if (!regNumber) return { ok: false, error: '登録番号を指定してください' };

  var sheet = getSheet_();
  var row = findRowByRegNumber_(sheet, regNumber);
  if (row < 0) return { ok: false, error: '登録が見つかりません: ' + regNumber };

  var v = sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0];
  var name = v[2];
  var email = String(v[4] || '').trim();
  var adults = Number(v[6] || 0);
  var children = Number(v[7] || 0);
  var amount = Number(v[8] || 0);
  var paymentMethod = v[9] === '当日現金' ? 'cash' : 'transfer';

  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return { ok: false, error: 'この登録にはメールアドレスが正しく設定されていません。スプレッドシートのメールアドレス列を確認してください。' };
  }
  if (!name) {
    return { ok: false, error: 'この登録には代表者氏名が設定されていません。スプレッドシートを確認してください。' };
  }

  var token = buildQrToken_(regNumber);
  var qrBlob = fetchQrImage_(token);
  sendConfirmationEmail_({
    regNumber: regNumber, name: name, email: email, adults: adults, children: children,
    amount: amount, paymentMethod: paymentMethod, qrBlob: qrBlob
  });

  return { ok: true, email: email };
}

function resolveRegNumber_(payload) {
  if (payload.code) return verifyQrToken_(payload.code);
  if (payload.regNumber) return String(payload.regNumber).trim();
  return null;
}

function handleLookup_(payload) {
  requirePin_(payload);
  var regNumber = resolveRegNumber_(payload);
  if (!regNumber) return { ok: false, error: 'QRコードまたは登録番号を指定してください' };
  var sheet = getSheet_();
  var row = findRowByRegNumber_(sheet, regNumber);
  if (row < 0) return { ok: false, error: '登録が見つかりません: ' + regNumber };
  var v = sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0];
  return {
    ok: true, regNumber: v[0], name: v[2], adults: v[6], children: v[7],
    amount: v[8], paymentMethod: v[9], paymentStatus: v[10], checkin: v[11]
  };
}

function handleMarkPaid_(payload) {
  requirePin_(payload);
  var regNumber = resolveRegNumber_(payload);
  if (!regNumber) return { ok: false, error: '登録番号を指定してください' };
  var sheet = getSheet_();
  var row = findRowByRegNumber_(sheet, regNumber);
  if (row < 0) return { ok: false, error: '登録が見つかりません' };
  sheet.getRange(row, 11).setValue('確認済');
  return { ok: true };
}

function handleCheckin_(payload) {
  requirePin_(payload);
  var regNumber = resolveRegNumber_(payload);
  if (!regNumber) return { ok: false, error: '登録番号を指定してください' };
  var sheet = getSheet_();
  var row = findRowByRegNumber_(sheet, regNumber);
  if (row < 0) return { ok: false, error: '登録が見つかりません' };
  sheet.getRange(row, 12).setValue('済');
  return { ok: true };
}

function handleList_(payload) {
  requirePin_(payload);
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    rows.push({
      regNumber: data[i][0], name: data[i][2], adults: data[i][6], children: data[i][7],
      amount: data[i][8], paymentMethod: data[i][9], paymentStatus: data[i][10], checkin: data[i][11]
    });
  }
  return { ok: true, rows: rows };
}

function sendConfirmationEmail_(info) {
  var eventName = getProp_('EVENT_NAME', 'フラメンコ公演');
  var eventDate = getProp_('EVENT_DATE', '2027年11月21日(日) 開場15:30 / 開演16:00 / 終演予定20:00');
  var venue = getProp_('EVENT_VENUE', '那覇市ぶんかテンプス館 4階テンプスホール');
  var deadline = getProp_('APPLY_DEADLINE', '2027年11月18日(木)');
  var contactEmail = getProp_('CONTACT_EMAIL', 'okinawaflamenco@gmail.com');
  var contactTel = getProp_('CONTACT_TEL', '090-9149-5228(広報:奥平)');
  var organizerUrl = getProp_('ORGANIZER_URL', 'https://carlosgomezflamencoshcool.jimdoweb.com/');
  var bankOkinawa = getProp_('BANK_OKINAWA', '沖縄銀行 ◯◯支店 普通 ◯◯◯◯◯◯◯ 口座名義:◯◯◯◯(確定次第ご連絡します)');
  var bankRyukyu = getProp_('BANK_RYUKYU', '琉球銀行 ◯◯支店 普通 ◯◯◯◯◯◯◯ 口座名義:◯◯◯◯(確定次第ご連絡します)');
  var last4 = info.regNumber.replace('TF-', '');

  var paymentBlock;
  if (info.paymentMethod === 'cash') {
    paymentBlock = '<p>お支払い方法:<b>当日現金払い</b><br>' +
      '受付にて合計金額 <b>' + info.amount.toLocaleString() + '円</b> をお支払いください。</p>';
  } else {
    paymentBlock =
      '<p>お支払い方法:<b>口座振込</b></p>' +
      '<p>下記いずれかの口座に、<b>' + deadline + '</b>までにお振込みください。<br>' +
      '<b style="color:#b00020">お振込みの際は、お名前の前に登録番号「' + last4 + '」を入力してください。</b><br>' +
      '(例:' + last4 + 'ヤマダハナコ)</p>' +
      '<p>【沖縄銀行】<br>' + bankOkinawa + '</p>' +
      '<p>【琉球銀行】<br>' + bankRyukyu + '</p>' +
      '<p>合計金額:<b>' + info.amount.toLocaleString() + '円</b></p>';
  }

  var html =
    '<div style="font-family:sans-serif;line-height:1.7;color:#222">' +
    '<h2>' + eventName + ' お申込み確認</h2>' +
    '<p>' + info.name + ' 様</p>' +
    '<p>この度はお申込みいただきありがとうございます。以下の内容で受け付けました。</p>' +
    '<table style="border-collapse:collapse">' +
    '<tr><td style="padding:4px 12px 4px 0">登録番号</td><td><b>' + info.regNumber + '</b></td></tr>' +
    '<tr><td style="padding:4px 12px 4px 0">日時</td><td>' + eventDate + '</td></tr>' +
    '<tr><td style="padding:4px 12px 4px 0">会場</td><td>' + venue + '</td></tr>' +
    '<tr><td style="padding:4px 12px 4px 0">人数</td><td>大人 ' + info.adults + ' 名 / 子供(中学生まで) ' + info.children + ' 名</td></tr>' +
    '</table>' +
    paymentBlock +
    '<p>当日は下記のQRコードを受付でご提示ください。</p>' +
    '<img src="cid:qrcode" width="220" height="220" alt="QRコード" />' +
    '<p>主催:Carlos Gomez Flamenco School (<a href="' + organizerUrl + '">' + organizerUrl + '</a>)</p>' +
    '<p>お問い合わせ:' + contactEmail + ' / TEL ' + contactTel + '</p>' +
    '</div>';

  GmailApp.sendEmail(info.email, '【' + eventName + '】お申込み確認(登録番号:' + info.regNumber + ')',
    'HTMLメール対応のメールソフトでご覧ください。', {
      htmlBody: html,
      inlineImages: { qrcode: info.qrBlob },
      name: eventName + ' 事務局'
    });
}
