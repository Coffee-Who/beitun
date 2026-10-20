/**
 * 北屯生活圈：一鍵建立「新增地點」Google 表單 + 試算表
 * 用法：到 https://script.google.com 新增專案 → 把這整份貼上 → 按「執行」→ 允許權限
 * 執行完在下方「執行記錄」會看到兩個網址，貼到 GitHub 的 config.json
 */
function 建立投稿表單() {
  const CATS = {
    '醫院': ['急診重症','內科','外科','婦幼','五官','皮膚身心','復健照護','檢查診斷'],
    '美食': ['中式','日式','歐美','韓式','東南亞','台灣小吃','飲料咖啡','甜點冰品'],
    '運動': ['健身房','游泳池','球場','運動公園'],
    '生活': ['衣著配件','書籍文具','居家修繕','日常採買','交通車輛','個人服務','3C專業']
  };
  const options = [];
  Object.keys(CATS).forEach(function (m) { CATS[m].forEach(function (s) { options.push(m + '・' + s); }); });

  const form = FormApp.create('北屯生活圈｜新增地點');
  form.setDescription('推薦台中捷運北屯總站附近的好地方，送出後經管理員審核，就會出現在網站上。');
  form.addTextItem().setTitle('名稱').setRequired(true);
  form.addListItem().setTitle('分類').setChoiceValues(options).setRequired(true);
  form.addTextItem().setTitle('細項').setHelpText('例如：牙科、拉麵、修改衣服（可空白）');
  form.addTextItem().setTitle('地址').setRequired(true);
  form.addTextItem().setTitle('電話');
  form.addTextItem().setTitle('營業時間');
  form.addParagraphTextItem().setTitle('推薦理由').setHelpText('例如：有兒童牙科、要提早 2 個月預約');
  form.addTextItem().setTitle('推薦人');
  form.addTextItem().setTitle('Google地圖連結').setHelpText('可空白；在 Google 地圖按「分享」複製連結');
  form.addTextItem().setTitle('緯度').setHelpText('可空白；Google 地圖在位置上按右鍵，第一個數字');
  form.addTextItem().setTitle('經度').setHelpText('可空白；第二個數字');

  const ss = SpreadsheetApp.create('北屯生活圈｜地點資料');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush();
  Utilities.sleep(3000);

  let sheet = null;
  SpreadsheetApp.openById(ss.getId()).getSheets().forEach(function (sh) { if (sh.getFormUrl()) sheet = sh; });
  if (sheet) {
    const c = sheet.getLastColumn() + 1;
    sheet.getRange(1, c).setValue('審核').setFontWeight('bold').setBackground('#FFF0DF');
    sheet.getRange(2, c, 999, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox('✓', '').build());
  }
  DriveApp.getFileById(ss.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const csv = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/gviz/tq?tqx=out:csv&gid=' + (sheet ? sheet.getSheetId() : 0);
  Logger.log('formUrl（投稿表單）：' + form.getPublishedUrl());
  Logger.log('sheetCsvUrl（資料）：' + csv);
  Logger.log('管理表單：' + form.getEditUrl());
  Logger.log('管理資料（可刪除不當投稿）：' + ss.getUrl());
}
