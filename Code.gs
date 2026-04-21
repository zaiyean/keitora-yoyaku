// ===== 定数設定 =====
var SPREADSHEET_ID = '1oYPRXas9W2YafJBRvzQ16h5CYsPYIuFboXMeP2S-d8g';
var SHEET_CONFIG = '設定';

// ===== メインエントリーポイント =====
function doPost(e) {
  var result;
  try {
    var params = JSON.parse(e.postData.contents);
    var action = params.action;

    if (action === 'getConfig')         result = getConfigList();
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
          id: String(data[i][0]),
          name: String(data[i][1]),
          questionItems: questionItems,
          purposeOptions: purposeOptions,
          notifyEmail: String(data[i][5] || '')
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
function getOrCreateSheet(jichikaiId, jichikaiName, questionItems) {
  var sheetName = jichikaiId + '_予約';
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var stdSet = ['氏名', '電話番号', '用途'];
    var extraItems = (questionItems || []).filter(function(item) { return stdSet.indexOf(item) === -1; });
    var headers = ['受付番号', '申請日時', '利用日', '開始時間', '終了時間', '自治会名',
                   '氏名', '電話番号', 'メールアドレス', '用途', 'ステータス', '承認却下日時'];
    headers = headers.concat(extraItems);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  return sheet;
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

// ===== ダブルブッキングチェック =====
function checkDoubleBooking(sheet, date, startTime) {
  var data = sheet.getDataRange().getValues();
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
    var rowStatus = data[i][10];
    if (dateStr === date && String(data[i][3]) === startTime) {
      if (rowStatus === '申請中' || rowStatus === '確定') return true;
    }
  }
  return false;
}

// ===== 終了時間計算 =====
function calcEndTime(startTime) {
  return (parseInt(startTime.split(':')[0]) + 1) + ':00';
}

// ===== 予約申請 =====
function submitReservation(params) {
  var configResult = getConfigById(params.jichikaiId);
  if (!configResult.success) return configResult;
  var config = configResult.data;

  var sheet = getOrCreateSheet(params.jichikaiId, config.name, config.questionItems);

  if (checkDoubleBooking(sheet, params.date, params.startTime)) {
    return { success: false, error: 'その日時はすでに予約が入っています。別の日時をお選びください。' };
  }

  var receiptNo = generateReceiptNo(sheet);
  var now = new Date();
  var endTime = calcEndTime(params.startTime);
  var answers = params.answers || {};

  var stdSet = ['氏名', '電話番号', '用途'];
  var extraItems = config.questionItems.filter(function(item) { return stdSet.indexOf(item) === -1; });

  var rowData = [
    receiptNo, now, params.date, params.startTime, endTime, config.name,
    answers['氏名'] || '',
    answers['電話番号'] || '',
    params.email || '',
    answers['用途'] || '',
    '申請中', ''
  ];
  extraItems.forEach(function(item) { rowData.push(answers[item] || ''); });

  var nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 8).setNumberFormat('@');
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

  var body = name + ' 様\n\n'
    + '以下の内容で予約申請を受け付けました。\n'
    + '管理者が確認後、改めてご連絡いたします。\n\n'
    + '① 受付番号：' + receiptNo + '\n'
    + '② 利用日時：' + params.date + '（' + youbi + '）' + params.startTime + '〜' + endTime + '\n'
    + '③ 用途：' + (answers['用途'] || '') + '\n\n'
    + '※この時点では予約は確定していません。\n'
    + '確定メールが届くまで、お待ちください。';

  GmailApp.sendEmail(email, '【仮受付】軽トラック予約申請を受け付けました', body);
}

// ===== 管理者通知メール =====
function sendAdminNotify(config, params, receiptNo, endTime) {
  var adminEmail = config.notifyEmail;
  if (!adminEmail) return;
  var answers = params.answers || {};
  var dateObj = new Date(params.date.replace(/\//g, '-'));
  var youbi = ['日','月','火','水','木','金','土'][dateObj.getDay()];

  var body = '新規予約申請が届きました。\n\n'
    + '① 受付番号：' + receiptNo + '\n'
    + '② 利用日時：' + params.date + '（' + youbi + '）' + params.startTime + '〜' + endTime + '\n'
    + '③ 氏名：' + (answers['氏名'] || '') + '\n'
    + '④ 用途：' + (answers['用途'] || '') + '\n'
    + '⑤ 電話番号：' + (answers['電話番号'] || '');

  GmailApp.sendEmail(adminEmail, '【新規申請】軽トラック予約申請が入りました', body);
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
    for (var c = 12; c < headers.length; c++) {
      if (headers[c]) extra[headers[c]] = row[c];
    }

    results.push({
      row: i + 1,
      receiptNo:  row[0],
      timestamp:  row[1] ? Utilities.formatDate(new Date(row[1]), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') : '',
      date:       dateStr,
      startTime:  startTime,
      endTime:    endTime,
      jichikaiName: row[5],
      name:       row[6],
      phone:      row[7],
      email:      row[8],
      purpose:    row[9],
      status:     row[10],
      approvedAt: row[11] ? Utilities.formatDate(new Date(row[11]), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') : '',
      extra:      extra
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

  var row = sheet.getRange(rowNum, 1, 1, 12).getValues()[0];
  sheet.getRange(rowNum, 11).setValue('確定');
  sheet.getRange(rowNum, 12).setValue(new Date());

  var email = row[8];
  if (email) {
    var dateStr = row[2] instanceof Date ? Utilities.formatDate(row[2], 'Asia/Tokyo', 'yyyy/MM/dd') : String(row[2]);
    var dateObj = new Date(dateStr.replace(/\//g, '-'));
    var youbi = ['日','月','火','水','木','金','土'][dateObj.getDay()];
    var body = row[6] + ' 様\n\n'
      + '予約が確定しました。当日はお気をつけてご利用ください。\n\n'
      + '① 受付番号：' + row[0] + '\n'
      + '② 利用日時：' + dateStr + '（' + youbi + '）' + row[3] + '〜' + row[4] + '\n'
      + '③ 用途：' + row[9];
    GmailApp.sendEmail(email, '【予約確定】軽トラック予約が確定しました', body);
  }

  return { success: true };
}

// ===== 予約却下 =====
function rejectReservation(jichikaiId, rowNum) {
  var sheetName = jichikaiId + '_予約';
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'シートが見つかりません' };

  var row = sheet.getRange(rowNum, 1, 1, 12).getValues()[0];
  sheet.getRange(rowNum, 11).setValue('却下');
  sheet.getRange(rowNum, 12).setValue(new Date());

  var email = row[8];
  if (email) {
    var dateStr = row[2] instanceof Date ? Utilities.formatDate(row[2], 'Asia/Tokyo', 'yyyy/MM/dd') : String(row[2]);
    var dateObj = new Date(dateStr.replace(/\//g, '-'));
    var youbi = ['日','月','火','水','木','金','土'][dateObj.getDay()];
    var body = row[6] + ' 様\n\n'
      + '大変申し訳ありませんが、以下の予約申請はお断りとなりました。\n\n'
      + '① 受付番号：' + row[0] + '\n'
      + '② 利用日時：' + dateStr + '（' + youbi + '）' + row[3] + '〜' + row[4] + '\n\n'
      + '別の日時でのご予約をお待ちしております。';
    GmailApp.sendEmail(email, '【却下】軽トラック予約がお断りとなりました', body);
  }

  return { success: true };
}

// ===== 予約削除（キャンセル） =====
function deleteReservation(jichikaiId, rowNum) {
  var sheetName = jichikaiId + '_予約';
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'シートが見つかりません' };

  sheet.getRange(rowNum, 11).setValue('キャンセル');
  sheet.getRange(rowNum, 12).setValue(new Date());
  return { success: true };
}

// ===== 管理者直接入力（確定済みとして登録） =====
function addDirectReservation(params) {
  var configResult = getConfigById(params.jichikaiId);
  if (!configResult.success) return configResult;
  var config = configResult.data;

  var sheet = getOrCreateSheet(params.jichikaiId, config.name, config.questionItems);

  var receiptNo = generateReceiptNo(sheet);
  var now = new Date();
  var endTime = calcEndTime(params.startTime);
  var answers = params.answers || {};

  var stdSet = ['氏名', '電話番号', '用途'];
  var extraItems = config.questionItems.filter(function(item) { return stdSet.indexOf(item) === -1; });

  var rowData = [
    receiptNo, now, params.date, params.startTime, endTime, config.name,
    answers['氏名'] || '',
    answers['電話番号'] || '',
    params.email || '',
    answers['用途'] || '',
    '確定', now
  ];
  extraItems.forEach(function(item) { rowData.push(answers[item] || ''); });

  var nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 8).setNumberFormat('@');
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

  var headers = ['自治会ID', '自治会名', '管理者パスワード', '質問項目', '用途選択肢', '通知メールアドレス'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var testData = [
    ['higashi', '東発区会', '1234', '氏名,組,電話番号,用途', '引越し・荷物運搬,農作業・資材運搬,ゴミ・廃材処理,草刈り・剪定ゴミ,その他', 'test@example.com'],
    ['nishi',   '西発区会', '5678', '氏名,組,電話番号,用途,地区', '資材運搬,粗大ゴミ,その他', 'test2@example.com']
  ];
  sheet.getRange(2, 1, testData.length, testData[0].length).setValues(testData);

  return 'セットアップ完了';
}
