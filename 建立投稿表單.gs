/**
 * 北屯生活圈：一鍵設定
 * 1. 建立「新增地點」Google 表單 + 資料試算表（含審核欄）
 * 2. 把 GitHub 上 data.csv 的現有資料搬進試算表（預設已審核）
 * 3. 自動補經緯度（用 Google 地圖連結或地址），之後新投稿也會自動補
 *
 * 用法：https://script.google.com → 新專案 → 整份貼上 → 存檔 → 上方選「建立投稿表單」→ 執行 → 允許權限
 * 執行完看下方「執行記錄」的兩個網址，貼進 GitHub 的 config.json
 */
const DATA_URL = 'https://coffee-who.github.io/beitun/data.csv';
const CATS = {
  '美食': ['大地商場','總站夜市','小吃・麵飯','雞排・鹹酥雞・滷味','便當・外送','甜點・麵包・冰品','飲料・咖啡・酒吧','火鍋・燒肉・鍋物','早餐・早午餐','聚餐・合菜','異國料理','素食'],
  '生活': ['超市・賣場・哪裡買','補習・才藝','幼兒園・托嬰','汽機車・洗車','寵物美容・住宿・醫院','洗冷氣・水電・搬家','剪髮・美甲・美容','按摩・整復','電腦・手機','洗衣・改衣・洗鞋','拜拜・土地公','其他服務'],
  '運動': ['健身・皮拉提斯','兒童運動課','球場・揪團'],
  '醫院': ['小孩看病・耳鼻喉','感冒・疫苗・家醫','皮膚科・醫美','牙醫','中醫・針灸','骨科・復健','眼科・配眼鏡','胃腸鏡・外科專科','婦產科','醫院・藥局']
};

function 建立投稿表單() {
  const options = [];
  Object.keys(CATS).forEach(function (m) { CATS[m].forEach(function (s) { options.push(m + '・' + s); }); });

  const form = FormApp.create('北屯生活圈｜新增地點');
  form.setDescription('推薦台中捷運北屯總站附近的好地方，管理員審核後就會出現在網站上。');
  form.addTextItem().setTitle('名稱').setRequired(true);
  form.addListItem().setTitle('分類').setChoiceValues(options).setRequired(true);
  form.addTextItem().setTitle('細項').setHelpText('例如：牙科、拉麵、火鍋（可空白）');
  form.addTextItem().setTitle('地址');
  form.addTextItem().setTitle('電話');
  form.addTextItem().setTitle('營業時間');
  form.addParagraphTextItem().setTitle('推薦理由').setHelpText('例如：有兒童牙科、要提早 2 個月預約');
  form.addTextItem().setTitle('推薦人');
  form.addTextItem().setTitle('Google地圖連結').setHelpText('建議填：在 Google 地圖按「分享」→ 複製連結，位置會最準');
  form.addTextItem().setTitle('緯度').setHelpText('可空白，系統會自動補');
  form.addTextItem().setTitle('經度').setHelpText('可空白，系統會自動補');

  const ss = SpreadsheetApp.create('北屯生活圈｜地點資料');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush();
  Utilities.sleep(4000);

  const sheet = 取得資料表_(ss.getId());
  sheet.getRange('A:Z').setNumberFormat('@');
  const head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const c = head.length + 1;
  sheet.getRange(1, c).setValue('審核').setFontWeight('bold').setBackground('#FFF0DF');
  sheet.getRange(2, c, 1999, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox('✓', '').build());

  匯入現有資料_(sheet);
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());
  補座標();

  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('補座標').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('補座標').timeBased().everyHours(1).create();

  DriveApp.getFileById(ss.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const csv = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/gviz/tq?tqx=out:csv&gid=' + sheet.getSheetId();
  Logger.log('formUrl（投稿表單）：' + form.getPublishedUrl());
  Logger.log('sheetCsvUrl（資料）：' + csv);
  Logger.log('管理資料（審核、修改、刪除）：' + ss.getUrl());
}

function 取得資料表_(id) {
  let sheet = null;
  SpreadsheetApp.openById(id).getSheets().forEach(function (sh) { if (sh.getFormUrl()) sheet = sh; });
  return sheet || SpreadsheetApp.openById(id).getSheets()[0];
}

function 匯入現有資料_(sheet) {
  let text;
  try { text = UrlFetchApp.fetch(DATA_URL, { muteHttpExceptions: true }).getContentText('UTF-8'); } catch (e) { Logger.log('讀不到 data.csv，略過匯入'); return; }
  const rows = Utilities.parseCsv(text.replace(/^\uFEFF/, ''));
  const h = rows.shift().map(function (x) { return x.trim(); });
  const g = function (r, n) { const i = h.indexOf(n); return i < 0 ? '' : (r[i] || '').trim(); };
  const head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const out = rows.filter(function (r) { return g(r, '名稱'); }).map(function (r) {
    const v = {
      '名稱': g(r, '名稱'), '分類': g(r, '大分類') + '・' + g(r, '第二層'), '細項': g(r, '第三層'),
      '地址': g(r, '地址'), '電話': g(r, '電話'), '營業時間': g(r, '營業時間'), '推薦理由': g(r, '推薦理由'),
      '推薦人': g(r, '推薦人'), 'Google地圖連結': g(r, 'Google地圖連結'), '緯度': g(r, '緯度'), '經度': g(r, '經度'), '審核': '✓'
    };
    return head.map(function (k) { return v[k] !== undefined ? v[k] : ''; });
  });
  if (out.length) sheet.getRange(sheet.getLastRow() + 1, 1, out.length, head.length).setValues(out);
  Logger.log('已匯入 ' + out.length + ' 筆現有資料');
}

/** 把沒有經緯度的列補上；可隨時手動執行 */
function 補座標() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID'); if (!id) return;
  const sheet = 取得資料表_(id);
  const data = sheet.getDataRange().getValues();
  const head = data[0], col = function (n) { return head.indexOf(n); };
  const iN = col('名稱'), iA = col('地址'), iL = col('Google地圖連結'), iLat = col('緯度'), iLng = col('經度');
  const geo = Maps.newGeocoder().setRegion('tw').setLanguage('zh-TW');
  let n = 0;
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (!row[iN] || (row[iLat] && row[iLng])) continue;
    let p = row[iL] ? 從連結取座標_(String(row[iL])) : null;
    if (!p && row[iA]) p = 地址轉座標_(geo, String(row[iA]));
    if (!p) p = 地址轉座標_(geo, row[iN] + ' 台中市');
    if (p) { sheet.getRange(r + 1, iLat + 1).setValue(String(p[0])); sheet.getRange(r + 1, iLng + 1).setValue(String(p[1])); n++; }
    Utilities.sleep(200);
  }
  Logger.log('補上座標 ' + n + ' 筆');
}

function 從連結取座標_(url) {
  try {
    let u = url;
    for (let k = 0; k < 4; k++) {
      const m = u.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/) || u.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/) || u.match(/[?&]q=(-?\d+\.\d+),(-?\d+\.\d+)/);
      if (m) return 合理_([parseFloat(m[1]), parseFloat(m[2])]);
      const res = UrlFetchApp.fetch(u, { followRedirects: false, muteHttpExceptions: true });
      const loc = res.getHeaders()['Location'] || res.getHeaders()['location'];
      if (!loc) {
        const body = res.getContentText();
        const b = body.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/) || body.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
        return b ? 合理_([parseFloat(b[1]), parseFloat(b[2])]) : null;
      }
      u = decodeURIComponent(loc);
    }
  } catch (e) {}
  return null;
}

function 地址轉座標_(geo, q) {
  try {
    const res = geo.geocode(q);
    if (res.status === 'OK' && res.results.length) { const l = res.results[0].geometry.location; return 合理_([l.lat, l.lng]); }
  } catch (e) {}
  return null;
}

function 合理_(p) { return (p[0] > 23.9 && p[0] < 24.5 && p[1] > 120.4 && p[1] < 121.1) ? [Math.round(p[0] * 1e6) / 1e6, Math.round(p[1] * 1e6) / 1e6] : null; }
