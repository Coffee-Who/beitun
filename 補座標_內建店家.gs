/**
 * 機捷特區・單元12生活區：替內建的 648 個地點補上準確座標
 *
 * 用法（只要做一次）：
 * 1. 先把 stores.csv 上傳到 GitHub 的 beitun，確認 https://coffee-who.github.io/beitun/stores.csv 打得開
 * 2. 打開原本建立投稿表單的那個 Apps Script 專案（一定要同一個專案，才找得到試算表）
 * 3. 左邊「檔案」按＋ →「指令碼」→ 取名「補座標_內建店家」→ 把這整份貼上 → 存檔
 * 4. 上方函式選「匯入內建店家並補座標」→ 執行
 * 5. 看下方「執行記錄」：如果寫「還有 N 筆」，再按一次執行（每次最多跑 5 分鐘，會接著做）
 *
 * 做了什麼：
 * - 用地址（只寫到區域的就用「店名＋區」）向 Google 地圖查座標
 * - 把店家連同座標加進試算表，推薦人寫「內建資料」、審核打勾
 * - 網頁會依店名把座標合併回內建資料，公里數和地圖圖釘就會出現
 * - 查到的位置離總站超過 10 公里（多半是查到同名的別家），網頁會自動忽略
 *
 * 需要同一個專案裡原本的：取得資料表_、從連結取座標_、地址轉座標_
 */
const STORES_URL = 'https://coffee-who.github.io/beitun/stores.csv';
const TIME_LIMIT_MS = 5 * 60 * 1000;

function 匯入內建店家並補座標() {
  const start = Date.now();
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (!id) throw new Error('找不到試算表：請在原本建立投稿表單的那個 Apps Script 專案裡執行');
  const sheet = 取得資料表_(id);

  const res = UrlFetchApp.fetch(STORES_URL, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('讀不到 ' + STORES_URL + '，請先把 stores.csv 上傳到 GitHub');
  const rows = Utilities.parseCsv(res.getContentText('UTF-8').replace(/^﻿/, ''));
  const h = rows.shift().map(function (x) { return x.trim(); });
  const g = function (r, k) { const i = h.indexOf(k); return i < 0 ? '' : String(r[i] || '').trim(); };

  const head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const iName = head.indexOf('名稱');
  const last = sheet.getLastRow();
  const have = {};
  if (last > 1) sheet.getRange(2, iName + 1, last - 1, 1).getValues().forEach(function (r) { have[String(r[0]).trim()] = true; });

  const geo = Maps.newGeocoder().setRegion('tw').setLanguage('zh-TW');
  const out = [];
  let found = 0, notFound = 0, left = 0;
  rows.forEach(function (r) {
    const name = g(r, '名稱');
    if (!name || have[name]) return;
    have[name] = true;
    if (Date.now() - start > TIME_LIMIT_MS) { left++; return; }   // 下次執行再做

    const loc = g(r, '位置');
    const addr = 轉地址_(loc);
    const link = g(r, 'Google地圖連結');
    let p = link ? 從連結取座標_(link) : null;
    if (!p && addr) p = 地址轉座標_(geo, addr);
    if (!p) p = 地址轉座標_(geo, name.replace(/[（(].*?[）)]/g, '') + ' 台中市' + 區_(loc));
    if (p) found++; else notFound++;
    Utilities.sleep(120);

    const v = {
      '名稱': name,
      '分類': g(r, '大分類') + '・' + g(r, '第二層'),
      '細項': g(r, '第三層'),
      '地址': addr,
      '營業時間': g(r, '營業時間'),
      '推薦人': '內建資料',
      'Google地圖連結': link,
      '緯度': p ? String(p[0]) : '',
      '經度': p ? String(p[1]) : '',
      '審核': '✓'
    };
    out.push(head.map(function (k) { return v[k] !== undefined ? v[k] : ''; }));
  });

  if (out.length) sheet.getRange(sheet.getLastRow() + 1, 1, out.length, head.length).setValues(out);
  Logger.log('這次加入 ' + out.length + ' 筆：查到座標 ' + found + ' 筆、查不到 ' + notFound + ' 筆');
  if (left) Logger.log('還有 ' + left + ' 筆，請再按一次「執行」');
  else Logger.log('全部做完了，約 5 分鐘後網站就會更新');
}

/* 「北屯 敦富二街93號」→「台中市北屯區敦富二街93號」；只寫到區域的回傳空白，避免查到區中心 */
function 轉地址_(loc) {
  const s = String(loc || '').replace(/[（(].*?[）)]/g, '').trim();
  if (!/[路街巷道]|商場|夜市|市場|公園/.test(s)) return '';
  const m = s.match(/^(北屯|潭子|北區|西區|西屯|南屯|東區|中區|南區|太平|大雅)\s*(.*)$/);
  if (m) return '台中市' + m[1] + (m[1].slice(-1) === '區' ? '' : '區') + m[2];
  return '台中市' + s.replace(/^台中\s*/, '');
}

/* 位置開頭的區名，接在店名後面查得比較準 */
function 區_(loc) {
  const m = String(loc || '').match(/^(北屯|潭子|北區|西區|西屯|南屯|東區|中區|南區|太平|大雅)/);
  return m ? (m[1].slice(-1) === '區' ? m[1] : m[1] + '區') : '';
}

/**
 * 補地址：替「內建資料」裡已經有座標、但地址空白的店，用 Google 地圖反查地址
 * 用法：上方函式選「補地址」→ 執行；寫「還有 N 筆」就再按一次
 * 網頁會把查到的地址換掉原本只寫「北屯」「台中」的位置，商圈也會跟著重新歸類
 */
function 補地址() {
  const start = Date.now();
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (!id) throw new Error('找不到試算表：請在原本建立投稿表單的那個 Apps Script 專案裡執行');
  const sheet = 取得資料表_(id);
  const data = sheet.getDataRange().getValues();
  const head = data[0], col = function (n) { return head.indexOf(n); };
  const iA = col('地址'), iLat = col('緯度'), iLng = col('經度'), iBy = col('推薦人');
  const geo = Maps.newGeocoder().setRegion('tw').setLanguage('zh-TW');
  let done = 0, failed = 0, left = 0;
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (String(row[iBy]) !== '內建資料' || String(row[iA]).trim() || !row[iLat] || !row[iLng]) continue;
    if (Date.now() - start > TIME_LIMIT_MS) { left++; continue; }
    let addr = '';
    try {
      const res = geo.reverseGeocode(Number(row[iLat]), Number(row[iLng]));
      if (res.status === 'OK' && res.results.length) addr = 整理地址_(res.results[0].formatted_address);
    } catch (e) {}
    if (addr) { sheet.getRange(r + 1, iA + 1).setValue(addr); done++; } else failed++;
    Utilities.sleep(120);
  }
  Logger.log('補上地址 ' + done + ' 筆；查不到 ' + failed + ' 筆');
  if (left) Logger.log('還有 ' + left + ' 筆，請再按一次「執行」');
  else Logger.log('全部做完了，約 5 分鐘後網站就會更新');
}

/* 「406台灣台中市北屯區敦富路68號」→「台中市北屯區敦富路68號」 */
function 整理地址_(s) {
  return String(s || '').replace(/^\d{3,6}/, '').replace(/^台灣/, '').replace(/^臺中市/, '台中市').trim();
}
