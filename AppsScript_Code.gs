// LIGHTMED 裝機報告 / 耗材訂購 — Google Apps Script 後端
// 功能：接收前端 index.html 送出的 JSON（依 formType 分流），寫入對應 Google Sheet 分頁並寄通知信。

var NOTIFY_EMAIL = 'service@lightmed.com';

var INSTALL_COLUMNS = [
  ['clinicName',          '醫院/診所名稱'],
  ['doctorName',          '醫生/聯繫窗口姓名'],
  ['contactTitle',        '職稱'],
  ['country',             '國家'],
  ['clinicAddress',       '醫院/診所 地址'],
  ['clinicPhone',         '醫院/診所 電話'],
  ['contactEmail',        '醫院/診所 聯繫電子郵件'],
  ['installDate',         '裝機時間'],
  ['mainSN',              '主機序號'],
  ['applicatorSN',        'Applicator序號'],
  ['dealer',              '經銷商'],
  ['mainAppearance',      '主機外觀檢查'],
  ['mainAppearanceNote',  '主機外觀檢查備註'],
  ['laserPower',          '雷射功率檢查'],
  ['laserPowerNote',      '雷射功率檢查備註'],
  ['handpieceQty',        '手機數量'],
  ['handpieceNote',       '手機備註'],
  ['fiberQty',            '光針數量'],
  ['fiberNote',           '光針備註'],
  ['mirrorQty',           '反射鏡數量'],
  ['mirrorNote',          '反射鏡備註'],
  ['remark',              '備註'],
  ['submitTime',          '填答時間']
];

// 與 index.html 內的 TIPS 陣列保持一致（新增/調整型號時兩邊都要改）
var TIPS = [
  'ST010', 'ST100', 'ST110', 'ST120',
  'ST200', 'ST201', 'ST205', 'ST214', 'ST215', 'ST220', 'ST240', 'ST250', 'ST270',
  'ST700', 'ST701', 'ST900'
];

var ORDER_INFO_COLUMNS = [
  ['oClinicName',    '醫院/診所名稱'],
  ['oContactName',   '聯繫人姓名'],
  ['oContactTitle',  '職稱'],
  ['oPhone',         '電話'],
  ['oShipAddress',   '收件地址'],
  ['oEmail',         '電子郵件'],
  ['oRemark',        '訂單備註'],
  ['oSubmitTime',    '填答時間']
];

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  if (data.formType === 'tipOrder') {
    handleTipOrder(data);
  } else {
    handleInstallReport(data);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleInstallReport(data) {
  var sheet = getOrCreateSheet('裝機報告', INSTALL_COLUMNS.map(function(c) { return c[1]; }));
  var row = INSTALL_COLUMNS.map(function(c) { return data[c[0]] || ''; });
  sheet.appendRow(row);

  var lines = INSTALL_COLUMNS.map(function(c) {
    var v = data[c[0]];
    return v ? (c[1] + '：' + v) : null;
  }).filter(Boolean);
  var subject = '【裝機報告】' + (data.clinicName || '未填診所名稱') + ' - ' + (data.mainSN || '');
  MailApp.sendEmail(NOTIFY_EMAIL, subject, lines.join('\n'));
}

function handleTipOrder(data) {
  var header = ORDER_INFO_COLUMNS.map(function(c) { return c[1]; })
    .concat(TIPS.map(function(sku) { return sku + ' 數量'; }))
    .concat(TIPS.map(function(sku) { return sku + ' 備註'; }));
  var sheet = getOrCreateSheet('耗材訂購', header);

  var row = ORDER_INFO_COLUMNS.map(function(c) { return data[c[0]] || ''; })
    .concat(TIPS.map(function(sku) { return data['tip_' + sku] || ''; }))
    .concat(TIPS.map(function(sku) { return data['tipNote_' + sku] || ''; }));
  sheet.appendRow(row);

  var lines = ORDER_INFO_COLUMNS.map(function(c) {
    var v = data[c[0]];
    return v ? (c[1] + '：' + v) : null;
  }).filter(Boolean);

  var itemLines = TIPS.map(function(sku) {
    var qty = data['tip_' + sku];
    if (!qty) return null;
    var note = data['tipNote_' + sku];
    return sku + ' x' + qty + (note ? '（備註：' + note + '）' : '');
  }).filter(Boolean);

  var body = lines.join('\n') + '\n\n訂購品項：\n' + itemLines.join('\n');
  var subject = '【耗材訂購】' + (data.oClinicName || '未填診所名稱') + ' - ' + itemLines.length + ' 項';
  MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
}

function getOrCreateSheet(name, headerRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headerRow);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ---- 測試用：在 Apps Script 編輯器手動執行此函式，確認寫入與寄信正常 ----
function test_doPost_installReport() {
  var fake = {
    postData: {
      contents: JSON.stringify({
        formType: 'installReport',
        clinicName: '測試診所', doctorName: '王醫師', contactTitle: 'Manager', country: '台灣',
        clinicAddress: '台北市...', clinicPhone: '0912345678',
        contactEmail: 'test@example.com',
        installDate: '2026-08-20', mainSN: 'EYSAA3-123456', applicatorSN: 'SAB0J5003',
        dealer: '鐳鼎科技', mainAppearance: '正常', laserPower: '正常',
        handpieceQty: '1', fiberQty: '5', mirrorQty: '0', remark: '',
        submitTime: new Date().toISOString()
      })
    }
  };
  doPost(fake);
}

function test_doPost_tipOrder() {
  var fake = {
    postData: {
      contents: JSON.stringify({
        formType: 'tipOrder',
        oClinicName: '測試診所', oContactName: '陳小姐', oContactTitle: 'Manager',
        oPhone: '0912345678', oShipAddress: '台北市信義區...', oEmail: 'test@example.com',
        oRemark: '', tip_ST100: '2', tip_ST200: '5', tip_ST700: '3',
        oSubmitTime: new Date().toISOString()
      })
    }
  };
  doPost(fake);
}
