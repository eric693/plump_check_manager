// config.js

// ── 日期工具 ───────────────────────────────────────────────────────────────
// toISOString() 回傳 UTC，在台灣（UTC+8）跨日時 split('T')[0] 會取到前一天
// 以下函式全部使用本地時區，供各頁面共用

function localDateStr(d) {
    return d.getFullYear() + '-' +
           String(d.getMonth() + 1).padStart(2, '0') + '-' +
           String(d.getDate()).padStart(2, '0');
}
function getTodayStr() { return localDateStr(new Date()); }

function formatLocalDateTime(d) {
    if (!d) return '';
    const dt = (d instanceof Date) ? d : new Date(d);
    if (isNaN(dt.getTime())) return String(d);
    return localDateStr(dt) + ' ' +
           String(dt.getHours()).padStart(2, '0') + ':' +
           String(dt.getMinutes()).padStart(2, '0');
}
// ────────────────────────────────────────────────────────────────────────────

const API_CONFIG = {
  // 正式環境的 API URL
  apiUrl: "https://script.google.com/macros/s/AKfycbw0sgWa8d4FTBg-PV6T6L_9qEt-Bdg8Lkc9jrFewdt_kKzBBdTtg1EOpyUfuciTfdI_Vg/exec",
  
  // 新增回呼網址
  redirectUrl: "https://eric693.github.io/an_check_manager/"
  // 你也可以在這裡加入其他設定，例如：
  // timeout: 5000,
  // version: 'v4.6.0'
};
// 👇 新增：為了兼容性，同時定義全域變數 apiUrl
const apiUrl = API_CONFIG.apiUrl;
