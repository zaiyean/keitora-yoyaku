// ===== 定数設定 =====
var SPREADSHEET_ID = '1oYPRXas9W2YafJBRvzQ16h5CYsPYIuFboXMeP2S-d8g';
var SHEET_CONFIG = '設定';
// HTMLファイルのホスティングURL（管理者通知メールのリンクに使用）
var HTML_BASE_URL = 'https://zaiyean.github.io/keitora-yoyaku';

// ===== メインエントリーポイント =====
function doPost(e) {
  var result;
  try {
    var params = JSON.parse(e.postData.contents);
    var action = params.action;

    if (action === 'getConfig')          result = getConfigList();
    else if (action === 'getConfigById') result = getConfigById(params.jichikaiId);
    else if (action === 'checkPassword') result = checkPassword(params.jichikaiId, params.password);
    else if (action === 'submit')        result = submitReservation(params);
    else if (action === 'getReservations') result = getReservations(params.jichikaiId, params.year, params.month);
    else if (action === 'approve')       result = approveReservation(params.jichikaiId, params.rowNum);
    else if (action === 'reject')        result = rejectReservation(params.jichikaiId, params.rowNum);
    else if (action === 'delete')        result = deleteReservation(params.jichikaiId, params.rowNum);
    else if (action === 'addDirect')     result = addDirectReservation(params);
    else if (action === 'initSheet')     result = initSheet(params.jichikaiId);
    else result = { success: false, error: '不明なアクション: ' + action };

  } catch (err) {
    result = { success: false, error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== 全自治会リスト取得（ID・名前のみ） =====
function getConfigList() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_CONFIG);
  if (!sheet) return { success: false, error: '設定シートが見つかりません' };

  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) list.push({ id: String(data[i][0]), name: String(data[i][1]) });
  }
  return { success: true, data: list };
}

// ===== 特定自治会の設定取得 =====
// 設定シート列順: 自治会ID(0), 自治会名(1), パスワード(2), 質問項目(3), 用途選択肢(4), 通知メール(5), 区長電話番号(6)
function getConfigById(jichikaiId) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_CONFIG);
  if (!sheet) return { success: false, error: '設定シートが見つかりません' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === jichikaiId) {
      var questionItems = data[i][3] ? String(data[i][3]).split(',').map(function(s) { return s.trim(); }) : [];
      var purposeOptions = data[i][4] ? String(data[i][4]).split(',').map(function(s) { return s.trim(); }) : [];
      return {
        success: true,
        data: {
          id:           String(data[i][0]),
          name:         String(data[i][1]),
          questionItems: questionItems,
          purposeOptions: purposeOptions,
          notifyEmail:  String(data[i][5] || ''),
          choPhone:     String(data[i][6] || '') // 区長電話番号
        }
      };
    }
  }
  return { success: false, error: '自治会が見つかりません' };
}

// ===== パスワード検証 =====
function checkPassword(jichikaiId, password) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_CONFIG);
  if (!sheet) return { success: false, error: '設定シートが見つかりません' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === jichikaiId) {
      if (String(data[i][2]) === String(password)) {
        return { success: true };
      } else {
        return { success: false, error: 'パスワードが違います' };
      }
    }
  }
  return { success: false, error: '自治会が見つかりません' };
}

// ===== 予約シートの取得または作成 =====
// 列順: 受付番号,申請日時,利用日,開始時間,終了時間,自治会名,[追加項目],氏名,電話番号,メールアドレス,用途,ステータス,承認却下日時
function getOrCreateSheet(jichikaiId, jichikaiName, questionItems) {
  var sheetName = jichikaiId + '_予約';
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var stdSet = ['氏名', '電話番号', '用途'];
    var extraItems = (questionItems || []).filter(function(item) { return stdSet.indexOf(item) === -1; });
    var headers = ['受付番号', '申請日時', '利用日', '開始時間', '終了時間', '自治会名'];
    headers = headers.concat(extraItems);
    headers = headers.concat(['氏名', '電話番号', 'メールアドレス', '用途', 'ステータス', '承認却下日時']);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  return sheet;
}

// ===== シートヘッダー取得 =====
function getSheetHeaders(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

// ===== 受付番号採番 =====
function generateReceiptNo(sheet) {
  var now = new Date();
  var dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd');
  var data = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && String(data[i][0]).indexOf('#' + dateStr) === 0) count++;
  }
  return '#' + dateStr + String(count + 1).padStart(2, '0');
}

// ===== ダブルブッキングチェック（時間範囲の重複チェック） =====
function checkDoubleBooking(sheet, date, newStart, newEnd) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var iStatus = headers.indexOf('ステータス');

  var newStartH = parseInt(newStart.split(':')[0]);
  var newEndH   = parseInt(newEnd.split(':')[0]);

  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][2];
    var dateStr = '';
    if (rowDate instanceof Date) {
      dateStr = Utilities.formatDate(rowDate, 'Asia/Tokyo', 'yyyy/MM/dd');
    } else {
      dateStr = String(rowDate)
        .replace(/(\d{4})年(\d{1,2})月(\d{1,2})日/, function(_, y, m, d) {
          return y + '/' + ('0'+m).slice(-2) + '/' + ('0'+d).slice(-2);
        })
        .replace(/^(\d{4})-(\d{2})-(\d{2})$/, '$1/$2/$3');
    }
    var rowStatus = data[i][iStatus];
    if (dateStr === date && (rowStatus === '申請中' || rowStatus === '確定')) {
      var rawStart = data[i][3];
      var rawEnd   = data[i][4];
      var existStartStr = rawStart instanceof Date ? Utilities.formatDate(rawStart, 'Asia/Tokyo', 'H:mm') : String(rawStart);
      var existEndStr   = rawEnd   instanceof Date ? Utilities.formatDate(rawEnd,   'Asia/Tokyo', 'H:mm') : String(rawEnd);
      var existStartH = parseInt(existStartStr.split(':')[0]);
      var existEndH   = parseInt(existEndStr.split(':')[0]);
      if (!(newEndH <= existStartH || newStartH >= existEndH)) return true;
    }
  }
  return false;
}

// ===== 終了時間計算（単一スロット用フォールバック） =====
function calcEndTime(startTime) {
  return (parseInt(startTime.split(':')[0]) + 1) + ':00';
}

// ===== 予約申請 =====
function submitReservation(params) {
  var configResult = getConfigById(params.jichikaiId);
  if (!configResult.success) return configResult;
  var config = configResult.data;

  var sheet = getOrCreateSheet(params.jichikaiId, config.name, config.questionItems);

  var endTime = params.endTime || calcEndTime(params.startTime);

  if (checkDoubleBooking(sheet, params.date, params.startTime, endTime)) {
    return { success: false, error: 'その日時はすでに予約が入っています。別の日時をお選びください。' };
  }

  var receiptNo = generateReceiptNo(sheet);
  var now = new Date();
  var answers = params.answers || {};

  var stdSet = ['氏名', '電話番号', '用途'];
  var extraItems = config.questionItems.filter(function(item) { return stdSet.indexOf(item) === -1; });

  var rowData = [receiptNo, now, params.date, params.startTime, endTime, config.name];
  extraItems.forEach(function(item) { rowData.push(answers[item] || ''); });
  rowData.push(
    answers['氏名'] || '',
    answers['電話番号'] || '',
    params.email || '',
    answers['用途'] || '',
    '申請中', ''
  );

  var nextRow = sheet.getLastRow() + 1;
  var phoneCol = 8 + extraItems.length;
  sheet.getRange(nextRow, phoneCol).setNumberFormat('@');
  sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);

  sendTempMail(config, params, receiptNo, endTime);
  sendAdminNotify(config, params, receiptNo, endTime);

  return { success: true, receiptNo: receiptNo };
}

// ===== 仮受付メール =====
function sendTempMail(config, params, receiptNo, endTime) {
  var email = params.email;
  if (!email) return;
  var answers = params.answers || {};
  var name = answers['氏名'] || '';
  var dateObj = new Date(params.date.replace(/\//g, '-'));
  var youbi = ['日','月','火','水','木','金','土'][dateObj.getDay()];
  var phoneStr = config.choPhone ? config.choPhone : '';

  var body = name + ' 様\n\n'
    + '以下の内容で予約申請を受け付けました。\n'
    + '管理者が確認後、改めてご連絡いたします。\n\n'
    + '① 受付番号：' + receiptNo + '\n'
    + '② 利用日時：' + params.date + '（' + youbi + '）' + params.startTime + '〜' + endTime + '\n'
    + '③ 用途：' + (answers['用途'] || '') + '\n\n'
    + '※この時点では予約は確定していません。\n'
    + '確定メールが届くまで、お待ちください。\n\n'
    + 'お問合わせ等ございましたら、LINEまたは電話(区長:' + phoneStr + ') でご連絡よろしくお願いします。\n\n'
    + config.name;

  GmailApp.sendEmail(email, '【仮受付】' + config.name + '　軽トラック予約申請を受け付けました', body);
}

// ===== 管理者通知メール =====
function sendAdminNotify(config, params, receiptNo, endTime) {
  var adminEmail = config.notifyEmail;
  if (!adminEmail) return;
  var answers = params.answers || {};
  var dateObj = new Date(params.date.replace(/\//g, '-'));
  var youbi = ['日','月','火','水','木','金','土'][dateObj.getDay()];
  var adminUrl = HTML_BASE_URL + '/admin.html?jichikai=' + encodeURIComponent(config.id);

  var body = '新規予約申請が届きました。\n\n'
    + '① 受付番号：' + receiptNo + '\n'
    + '② 利用日時：' + params.date + '（' + youbi + '）' + params.startTime + '〜' + endTime + '\n'
    + '③ 氏名：' + (answers['氏名'] || '') + '\n'
    + '④ 用途：' + (answers['用途'] || '') + '\n'
    + '⑤ 電話番号：' + (answers['電話番号'] || '') + '\n\n'
    + '管理者ページ（承認・却下はこちら）:\n' + adminUrl;

  GmailApp.sendEmail(adminEmail, '【新規申請】' + config.name + '　軽トラック予約申請が入りました', body);
}

// ===== 予約一覧取得 =====
function getReservations(jichikaiId, year, month) {
  var sheetName = jichikaiId + '_予約';
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: true, data: [] };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, data: [] };

  var headers = data[0];
  var iName       = headers.indexOf('氏名');
  var iPhone      = headers.indexOf('電話番号');
  var iEmail      = headers.indexOf('メールアドレス');
  var iPurpose    = headers.indexOf('用途');
  var iStatus     = headers.indexOf('ステータス');
  var iApprovedAt = headers.indexOf('承認却下日時');

  var results = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;

    var rowDate = row[2];
    var dateStr = '';
    try {
      if (rowDate instanceof Date) {
        dateStr = Utilities.formatDate(rowDate, 'Asia/Tokyo', 'yyyy/MM/dd');
      } else {
        dateStr = String(rowDate).trim()
          .replace(/(\d{4})年(\d{1,2})月(\d{1,2})日/, function(_, y, m, d) {
            return y + '/' + ('0'+m).slice(-2) + '/' + ('0'+d).slice(-2);
          })
          .replace(/^(\d{4})-(\d{2})-(\d{2})$/, '$1/$2/$3');
      }
    } catch(e) { continue; }

    var parts = dateStr.split('/');
    if (parts.length < 3) continue;
    if (parseInt(parts[0]) !== year || parseInt(parts[1]) !== month) continue;

    var startTime = row[3] instanceof Date ? Utilities.formatDate(row[3], 'Asia/Tokyo', 'H:mm') : String(row[3]);
    var endTime   = row[4] instanceof Date ? Utilities.formatDate(row[4], 'Asia/Tokyo', 'H:mm') : String(row[4]);

    var extra = {};
    for (var c = 6; c < iName; c++) {
      if (headers[c]) extra[headers[c]] = row[c];
    }

    results.push({
      row:          i + 1,
      receiptNo:    row[0],
      timestamp:    row[1] ? Utilities.formatDate(new Date(row[1]), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') : '',
      date:         dateStr,
      startTime:    startTime,
      endTime:      endTime,
      jichikaiName: row[5],
      name:         iName       >= 0 ? row[iName]       : '',
      phone:        iPhone      >= 0 ? row[iPhone]      : '',
      email:        iEmail      >= 0 ? row[iEmail]      : '',
      purpose:      iPurpose    >= 0 ? row[iPurpose]    : '',
      status:       iStatus     >= 0 ? row[iStatus]     : '',
      approvedAt:   (iApprovedAt >= 0 && row[iApprovedAt]) ? Utilities.formatDate(new Date(row[iApprovedAt]), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') : '',
      extra:        extra
    });
  }

  return { success: true, data: results };
}

// ===== 予約承認 =====
function approveReservation(jichikaiId, rowNum) {
  var sheetName = jichikaiId + '_予約';
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'シートが見つかりません' };

  // 自治会設定を取得（電話番号・名前をメールで使用）
  var configResult = getConfigById(jichikaiId);
  var config = configResult.success ? configResult.data : { name: '', choPhone: '' };

  var headers = getSheetHeaders(sheet);
  var iName       = headers.indexOf('氏名');
  var iEmail      = headers.indexOf('メールアドレス');
  var iPurpose    = headers.indexOf('用途');
  var iStatus     = headers.indexOf('ステータス') + 1;
  var iApprovedAt = headers.indexOf('承認却下日時') + 1;

  var row = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  sheet.getRange(rowNum, iStatus).setValue('確定');
  sheet.getRange(rowNum, iApprovedAt).setValue(new Date());

  var email = iEmail >= 0 ? row[iEmail] : '';
  if (email) {
    var jichikaiName = config.name || String(row[5]);
    var phoneStr = config.choPhone || '';
    var dateStr = row[2] instanceof Date ? Utilities.formatDate(row[2], 'Asia/Tokyo', 'yyyy/MM/dd') : String(row[2]);
    var dateObj = new Date(dateStr.replace(/\//g, '-'));
    var youbi = ['日','月','火','水','木','金','土'][dateObj.getDay()];
    var startTimeStr = row[3] instanceof Date ? Utilities.formatDate(row[3], 'Asia/Tokyo', 'H:mm') : String(row[3]);
    var endTimeStr   = row[4] instanceof Date ? Utilities.formatDate(row[4], 'Asia/Tokyo', 'H:mm') : String(row[4]);
    var name    = iName    >= 0 ? row[iName]    : '';
    var purpose = iPurpose >= 0 ? row[iPurpose] : '';

    var body = name + ' 様\n\n'
      + '予約が確定しました。当日はお気をつけてご利用ください。\n\n'
      + '① 受付番号：' + row[0] + '\n'
      + '② 利用日時：' + dateStr + '（' + youbi + '）' + startTimeStr + '〜' + endTimeStr + '\n'
      + '③ 用途：' + purpose + '\n\n'
      + 'お問合わせ等ございましたら、LINEまたは電話(区長:' + phoneStr + ') でご連絡よろしくお願いします。\n\n'
      + jichikaiName;

    GmailApp.sendEmail(email, '【予約確定】' + jichikaiName + '　軽トラック予約が確定しました', body);
  }

  return { success: true };
}

// ===== 予約却下 =====
function rejectReservation(jichikaiId, rowNum) {
  var sheetName = jichikaiId + '_予約';
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'シートが見つかりません' };

  var configResult = getConfigById(jichikaiId);
  var config = configResult.success ? configResult.data : { name: '' };

  var headers = getSheetHeaders(sheet);
  var iName       = headers.indexOf('氏名');
  var iEmail      = headers.indexOf('メールアドレス');
  var iStatus     = headers.indexOf('ステータス') + 1;
  var iApprovedAt = headers.indexOf('承認却下日時') + 1;

  var row = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  sheet.getRange(rowNum, iStatus).setValue('却下');
  sheet.getRange(rowNum, iApprovedAt).setValue(new Date());

  var email = iEmail >= 0 ? row[iEmail] : '';
  if (email) {
    var jichikaiName = config.name || String(row[5]);
    var dateStr = row[2] instanceof Date ? Utilities.formatDate(row[2], 'Asia/Tokyo', 'yyyy/MM/dd') : String(row[2]);
    var dateObj = new Date(dateStr.replace(/\//g, '-'));
    var youbi = ['日','月','火','水','木','金','土'][dateObj.getDay()];
    var startTimeStr = row[3] instanceof Date ? Utilities.formatDate(row[3], 'Asia/Tokyo', 'H:mm') : String(row[3]);
    var endTimeStr   = row[4] instanceof Date ? Utilities.formatDate(row[4], 'Asia/Tokyo', 'H:mm') : String(row[4]);
    var name = iName >= 0 ? row[iName] : '';

    var body = name + ' 様\n\n'
      + '大変申し訳ありませんが、以下の予約申請はお断りとなりました。\n\n'
      + '① 受付番号：' + row[0] + '\n'
      + '② 利用日時：' + dateStr + '（' + youbi + '）' + startTimeStr + '〜' + endTimeStr + '\n\n'
      + '別の日時でのご予約をお待ちしております。';

    GmailApp.sendEmail(email, '【却下】' + jichikaiName + '　軽トラック予約がお断りとなりました', body);
  }

  return { success: true };
}

// ===== 予約削除（キャンセル） =====
function deleteReservation(jichikaiId, rowNum) {
  var sheetName = jichikaiId + '_予約';
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'シートが見つかりません' };

  var headers = getSheetHeaders(sheet);
  var iStatus     = headers.indexOf('ステータス') + 1;
  var iApprovedAt = headers.indexOf('承認却下日時') + 1;

  sheet.getRange(rowNum, iStatus).setValue('キャンセル');
  sheet.getRange(rowNum, iApprovedAt).setValue(new Date());
  return { success: true };
}

// ===== 管理者直接入力（確定済みとして登録） =====
function addDirectReservation(params) {
  var configResult = getConfigById(params.jichikaiId);
  if (!configResult.success) return configResult;
  var config = configResult.data;

  var sheet = getOrCreateSheet(params.jichikaiId, config.name, config.questionItems);

  var endTime = params.endTime || calcEndTime(params.startTime);

  if (checkDoubleBooking(sheet, params.date, params.startTime, endTime)) {
    return { success: false, error: 'その日時はすでに予約が入っています。別の日時をお選びください。' };
  }

  var receiptNo = generateReceiptNo(sheet);
  var now = new Date();
  var answers = params.answers || {};

  var stdSet = ['氏名', '電話番号', '用途'];
  var extraItems = config.questionItems.filter(function(item) { return stdSet.indexOf(item) === -1; });

  var rowData = [receiptNo, now, params.date, params.startTime, endTime, config.name];
  extraItems.forEach(function(item) { rowData.push(answers[item] || ''); });
  rowData.push(
    answers['氏名'] || '',
    answers['電話番号'] || '',
    params.email || '',
    answers['用途'] || '',
    '確定', now
  );

  var nextRow = sheet.getLastRow() + 1;
  var phoneCol = 8 + extraItems.length;
  sheet.getRange(nextRow, phoneCol).setNumberFormat('@');
  sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);

  return { success: true, receiptNo: receiptNo };
}

// ===== 新規自治会シート初期化 =====
function initSheet(jichikaiId) {
  var configResult = getConfigById(jichikaiId);
  if (!configResult.success) return configResult;
  var config = configResult.data;
  getOrCreateSheet(jichikaiId, config.name, config.questionItems);
  return { success: true };
}

// ===== 初回セットアップ（GASエディタから手動実行） =====
function setupSpreadsheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  var sheet = ss.getSheetByName(SHEET_CONFIG);
  if (!sheet) sheet = ss.insertSheet(SHEET_CONFIG);

  var headers = ['自治会ID', '自治会名', '管理者パスワード', '質問項目', '用途選択肢', '通知メールアドレス', '区長電話番号'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var testData = [
    ['higashi', '東発区会', '1234', '氏名,町内会,組,電話番号,用途', '引越し・荷物運搬,農作業・資材運搬,ゴミ・廃材処理,草刈り・剪定ゴミ,その他', 'test@example.com', '090-1234-5678'],
    ['nishi',   '西発区会', '5678', '氏名,町内会,組,電話番号,用途', '資材運搬,粗大ゴミ,その他', 'test2@example.com', '090-8765-4321']
  ];
  sheet.getRange(2, 1, testData.length, testData[0].length).setValues(testData);

  return 'セットアップ完了';
}
