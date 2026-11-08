// ===== 機密設定：請到 Apps Script「專案設定 → 指令碼屬性」填寫，不要寫在程式碼裡 =====
// SPREADSHEET_ID            試算表 ID（測試環境與正式環境各填自己的）
// ADMIN_PASSWORD            後台密碼（未設定時後台無法登入）
// LINE_CHANNEL_ACCESS_TOKEN LINE 官方帳號金鑰（未設定時不會發送 LINE 通知）
// LINE_ADMIN_USER_ID        接收 LINE 通知的對象 ID
function getScriptProp_(key) {
  try { return (PropertiesService.getScriptProperties().getProperty(key) || "").toString().trim(); } catch (e) { return ""; }
}
// 不再寫死試算表 ID：避免測試環境忘了設定時，誤把測試訂單寫進正式試算表
const GLOBAL_SPREADSHEET_ID = getScriptProp_("SPREADSHEET_ID");
if (!GLOBAL_SPREADSHEET_ID) throw new Error("尚未設定 SPREADSHEET_ID，請到 Apps Script「專案設定 → 指令碼屬性」新增。");
const STORE_ADDRESS = "高雄市前鎮區景德路72巷12號";
const TIMEZONE = "GMT+8";
const ORDER_LOOKUP_URL = "https://www.lilou.food/";
const LILOU_EMAIL_FROM = "orders@lilou.food";
const LILOU_EMAIL_NAME = "Lilou 粒露";
const LILOU_EMAIL_REPLY_TO = "orders@lilou.food";
const LILOU_EMAIL_AUTO_NOTICE = "此信件由系統自動發送，請勿直接回覆此信。\n如有任何問題，歡迎透過 Lilou 官方 LINE 或 Instagram 與我們聯繫。";

const FORM_OPTIONS_CACHE_KEY = "lilou_form_options_v1";
const FORM_OPTIONS_CACHE_SECONDS = 60;

const ADMIN_TOKEN_CACHE_PREFIX = "lilou_admin_token_";
const ADMIN_TOKEN_SECONDS = 21600; // 6 小時

const LINE_CHANNEL_ACCESS_TOKEN = getScriptProp_("LINE_CHANNEL_ACCESS_TOKEN");
const LINE_ADMIN_USER_ID = getScriptProp_("LINE_ADMIN_USER_ID");

const ADMIN_LOGIN_FAIL_KEY = "lilou_admin_login_fail";
const ADMIN_LOGIN_MAX_FAILS = 5;          // 連續錯誤 5 次
const ADMIN_LOGIN_LOCK_SECONDS = 600;     // 鎖定 10 分鐘

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getTaiwanTodayTime() {
  const todayStr = Utilities.formatDate(new Date(), TIMEZONE, "yyyy/MM/dd");
  return new Date(todayStr + " 00:00:00").getTime();
}

function normalizeDateString(dateStr) {
  if (!dateStr) return "";
  let clean = dateStr.toString().trim().replace(/-/g, "/").replace(/\s+/g, "");
  if (!clean) return "";
  const parts = clean.split("/");
  if (parts.length !== 3) return "";
  const y = Number(parts[0]), m = Number(parts[1]), d = Number(parts[2]);
  if (!y || !m || !d) return "";
  return `${y}/${String(m).padStart(2, "0")}/${String(d).padStart(2, "0")}`;
}

function parseAvailableDates(value, minTime) {
  if (!value) return [];
  return value.toString().replace(/["'\r\n]/g, "").split(/[,，、\n]/).map(normalizeDateString).filter(Boolean).filter(dateStr => {
    const t = new Date(dateStr + " 00:00:00").getTime();
    return !isNaN(t) && t >= minTime;
  });
}

// 自動建立分類設定分頁
function ensureCategoriesSheet_(ss) {
  let sheet = ss.getSheetByName("Categories") || ss.getSheetByName("Categories 分類管理");
  if (!sheet) {
    sheet = ss.insertSheet("Categories");
    sheet.appendRow(["id", "name", "behavior", "display", "sort"]);
    sheet.appendRow(["combo", "迷你可麗露", "combo", "TRUE", "10"]);
    sheet.appendRow(["single", "限定甜點", "single", "TRUE", "20"]);
    sheet.appendRow(["giftbox", "中秋禮盒", "single", "TRUE", "30"]);
  }
  return sheet;
}

// ===== 第2階段：日期規則（全站 + 分類）=====
// Settings 新增的鍵：這些不是「日期清單」，解析時要排除
const LILOU_RULE_SETTING_KEYS = ["pickup_min_days", "pickup_max_days", "delivery_min_days", "delivery_max_days", "order_cutoff_time"];
const LILOU_RULE_DEFAULTS = { pickup_min_days: "3", pickup_max_days: "30", delivery_min_days: "7", delivery_max_days: "30", order_cutoff_time: "18:00" };
const CATEGORY_RULE_HEADERS = ["name_en", "pickup_min_days", "pickup_max_days", "delivery_min_days", "delivery_max_days", "pickup_date_rule", "delivery_date_rule", "date_notice", "daily_limit"];

// 自動在 Settings 補上日期規則的預設值（已存在就不動）
function ensureSettingsDateRules_(ss) {
  const sheet = ss.getSheetByName("Settings");
  if (!sheet) return;
  const keys = sheet.getDataRange().getValues().map(r => (r[0] || "").toString().trim());
  [["pickup_min_days", "3"], ["pickup_max_days", "30"], ["delivery_min_days", "7"], ["delivery_max_days", "30"], ["order_cutoff_time", "18:00"]].forEach(def => {
    if (keys.indexOf(def[0]) === -1) sheet.appendRow(def);
  });
}

// 規則字串 → 符合的日期清單判斷。格式：2026/09/11~2026/09/27, 2026/10/01（逗號或換行分隔，區間用 ~ 或 ～ 或「至」）
function lilouParseDateRule_(rule) {
  const text = (rule || "").toString().trim();
  if (!text) return null;
  const parts = text.split(/[,，、\n]+/).map(t => t.trim()).filter(Boolean);
  const ranges = [];
  parts.forEach(part => {
    const m = part.split(/\s*(?:~|～|至)\s*/);
    if (m.length === 2) {
      const a = normalizeDateString(m[0]) || "0000/00/00";
      const b = normalizeDateString(m[1]) || "9999/99/99";
      ranges.push([a, b]);
    } else {
      const d = normalizeDateString(part);
      if (d) ranges.push([d, d]);
    }
  });
  return ranges.length ? ranges : null;
}
function lilouDateInRule_(dateStr, ranges) {
  if (!ranges) return true;
  return ranges.some(r => dateStr >= r[0] && dateStr <= r[1]);
}
function lilouAddDays_(dateStr, days) {
  const p = dateStr.split("/").map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + days));
  return d.getUTCFullYear() + "/" + String(d.getUTCMonth() + 1).padStart(2, "0") + "/" + String(d.getUTCDate()).padStart(2, "0");
}
function lilouNum_(v) {
  const n = Number((v === undefined || v === null ? "" : v).toString().replace(/[^0-9.]/g, ""));
  return (v === undefined || v === null || v.toString().trim() === "" || isNaN(n)) ? null : n;
}
// 依台灣時間與截止時間，算出「下單基準日」：超過截止時間算隔天
function lilouOrderBaseDate_(cutoff) {
  const today = Utilities.formatDate(new Date(), TIMEZONE, "yyyy/MM/dd");
  const nowHm = Utilities.formatDate(new Date(), TIMEZONE, "HH:mm");
  const c = (cutoff || "").toString().trim();
  if (/^\d{1,2}:\d{2}$/.test(c) && nowHm >= c.padStart(5, "0")) return lilouAddDays_(today, 1);
  return today;
}
// 計算某配送方式下可選日期。kind = "pickup" | "delivery"；cats = 購物車內商品所屬分類物件
function lilouAllowedDates_(kind, openDates, settings, cats, baseDate) {
  const g = k => lilouNum_(settings[k] !== undefined ? settings[k] : LILOU_RULE_DEFAULTS[k]);
  // 分類有填就用分類的（可填 0＝當天），沒填才用全店規則；與前台算法相同
  const gMin = g(kind + "_min_days") || 0, gMax = g(kind + "_max_days") || 0;
  const list = cats || [];
  const effMin = c => { const v = lilouNum_(c[kind + "_min_days"]); return v === null ? gMin : v; };
  const effMax = c => { const v = lilouNum_(c[kind + "_max_days"]); return v === null ? (gMax > 0 ? gMax : Infinity) : v; };
  const minDays = list.length ? Math.max.apply(null, list.map(effMin)) : gMin;
  const maxDays = list.length ? Math.min.apply(null, list.map(effMax)) : (gMax > 0 ? gMax : Infinity);
  const from = lilouAddDays_(baseDate, minDays);
  const to = isFinite(maxDays) ? lilouAddDays_(baseDate, maxDays) : "9999/99/99";
  const rules = (cats || []).map(c => lilouParseDateRule_(c[kind + "_date_rule"]));
  return (openDates || []).map(normalizeDateString).filter(Boolean).filter(d => d >= from && d <= to && rules.every(r => lilouDateInRule_(d, r)));
}

// ===== 生日蛋糕 =====
const CAKE_OPTION_HEADERS = ["group", "name", "name_en", "price", "input", "included_qty", "max_qty", "display", "sort"];

// 依表頭位置寫入一列；缺少的表頭自動補在最右邊
function appendRowByHeader_(sheet, obj) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => h.toString().trim());
  const missing = Object.keys(obj).filter(k => headers.indexOf(k) === -1);
  if (missing.length) {
    sheet.getRange(1, headers.filter(Boolean).length + 1, 1, missing.length).setValues([missing]);
    headers = headers.filter(Boolean).concat(missing);
  }
  sheet.appendRow(headers.map(h => obj[h] !== undefined ? obj[h] : ""));
}

// 第一次使用：建立 Cake_Options 分頁、「生日蛋糕」分類與三款蛋糕（只做一次，之後店家刪改都不會再補回來）
function ensureCakeSetup_(ss) {
  let opt = ss.getSheetByName("Cake_Options");
  if (!opt) {
    opt = ss.insertSheet("Cake_Options");
    opt.appendRow(CAKE_OPTION_HEADERS);
    [
      ["base", "經典原味", "Original", "0", "", "", "", "TRUE", "10"],
      ["base", "台灣茶香", "Taiwanese Tea", "0", "", "", "", "TRUE", "20"],
      ["base", "濃醇巧克力", "Chocolate", "20", "", "", "", "TRUE", "30"],
      ["filling", "當季新鮮水果", "Fresh Fruits", "0", "", "", "", "TRUE", "10"],
      ["filling", "布丁夾餡", "Custard Pudding", "20", "", "", "", "TRUE", "20"],
      ["filling", "巧克力脆脆", "Chocolate Crunch", "20", "", "", "", "TRUE", "30"],
      ["addon", "蛋糕刀", "Cake Knife", "0", "check", "", "1", "TRUE", "10"],
      ["addon", "生日蠟燭", "Birthday Candle", "0", "check", "", "1", "TRUE", "20"],
      ["addon", "餐盤叉組", "Plates & Forks", "2", "qty", "8", "30", "TRUE", "30"],
      ["addon", "保冷袋", "Insulated Bag", "80", "check", "", "1", "TRUE", "40"]
    ].forEach(r => opt.appendRow(r));
  }
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty("CAKE_SEEDED") === "1") return;
  const cat = ss.getSheetByName("Categories") || ss.getSheetByName("Categories 分類管理");
  const prod = ss.getSheetByName("Products");
  if (!cat || !prod) return;
  const hasCake = cat.getDataRange().getDisplayValues().some(r => r.some(v => v.toString().trim().toLowerCase() === "cake"));
  if (!hasCake) {
    appendRowByHeader_(cat, { id: "cake", name: "生日蛋糕", name_en: "Birthday Cakes", behavior: "cake", display: "TRUE", sort: "25", pickup_min_days: "3", date_notice: "生日蛋糕需提前 3 天預訂。" });
    [["經典鮮奶油蛋糕｜4 吋", "Classic Cream Cake · 4 inch", "10"], ["經典鮮奶油蛋糕｜6 吋", "Classic Cream Cake · 6 inch", "20"], ["蘋果與花｜6 吋", "Apple & Flowers · 6 inch", "30"]].forEach(c => {
      appendRowByHeader_(prod, { name: c[0], name_en: c[1], price: "800", stock: "999", type: "cake", display: "TRUE", shipping_rule: "pickup", sort: c[2] });
    });
  }
  props.setProperty("CAKE_SEEDED", "1");
}

function readCakeOptions_(ss) {
  return readSheetObjects_(ss, "Cake_Options").filter(o => o.name && lilouIsOn_(o.display)).map(o => ({
    group: (o.group || "addon").trim().toLowerCase(), name: o.name.trim(), name_en: (o.name_en || "").trim(),
    price: Number((o.price || "0").toString().replace(/[^0-9.]/g, "")) || 0,
    input: (o.input || "check").trim().toLowerCase(), included_qty: Number(o.included_qty) || 0, max_qty: Number(o.max_qty) || 0, sort: Number(o.sort) || 999
  })).sort((a, b) => a.sort - b.sort);
}

// 每日限量：統計各分類每個取貨日已被訂走的數量（已取消不算）
function countCategoryBookings_(ss, categories, products) {
  const limited = categories.filter(c => Number(c.daily_limit) > 0);
  if (!limited.length) return {};
  const orders = readSheetObjects_(ss, "Orders");
  const result = {};
  limited.forEach(c => {
    const names = products.filter(p => p.type === c.id).map(p => p.name);
    const booked = {};
    orders.forEach(o => {
      if ((o.Order_Status || "").trim() === "已取消") return;
      const date = normalizeDateString((o.Date || "").toString().slice(0, 10));
      if (!date) return;
      const text = (o.Product_Name || "").toString();
      let qty = 0;
      names.forEach(n => {
        const re = new RegExp(n.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\((\\d+)", "g");
        let m; while ((m = re.exec(text)) !== null) qty += Number(m[1]) || 0;
      });
      if (qty) booked[date] = (booked[date] || 0) + qty;
    });
    result[c.id] = booked;
  });
  return result;
}

// ===== 自訂內容頁（例：客製服務、關於我們）=====
const PAGE_HEADERS = ["page_id", "nav_label", "nav_label_en", "title", "title_en", "intro", "display", "sort"];
const PAGE_BLOCK_HEADERS = ["block_id", "page_id", "title", "title_en", "body", "image", "button_text", "button_link", "display", "sort"];

// 第一次使用時自動建立 Pages / PageBlocks 分頁，並帶入原本放在商品裡的四項客製服務（預設不顯示，店家確認後再開啟）
function ensurePagesSheets_(ss) {
  let pages = ss.getSheetByName("Pages");
  if (!pages) {
    pages = ss.insertSheet("Pages");
    pages.appendRow(PAGE_HEADERS);
    pages.appendRow(["custom", "客製服務", "Custom Orders", "客製服務", "Custom Orders", "婚禮、彌月、活動與企業訂購，歡迎透過官方 LINE 與我們討論。", "FALSE", "10"]);
  }
  let blocks = ss.getSheetByName("PageBlocks");
  if (!blocks) {
    blocks = ss.insertSheet("PageBlocks");
    blocks.appendRow(PAGE_BLOCK_HEADERS);
    const line = "https://lin.ee/ane79Cw";
    [
      ["b1", "custom", "婚禮小物", "Wedding Favors", "單顆獨立包裝｜50 顆起訂\n迷你盒裝｜迷你露 2 顆或大粒露 1 顆\n客製婚禮貼紙｜可加購專屬設計貼紙\n\n## 訂購流程\n- 透過官方 LINE 聯繫，提供婚禮日期、預估份數與預算。\n- 確認甜點品項、包裝形式及客製需求。\n- 確認報價、交期與訂購細節後安排製作。", "", "LINE 洽詢", line, "TRUE", "10"],
      ["b2", "custom", "彌月禮盒", "Baby Celebration", "分享新生命到來的喜悅，讓法式甜點成為彌月贈禮的另一種選擇。\n可依送禮對象與預算，挑選可麗露、費南雪等甜點組合，並洽詢彌月專屬包裝及試吃安排。", "", "LINE 洽詢", line, "TRUE", "20"],
      ["b3", "custom", "活動甜點", "Dessert Catering", "適用於品牌活動、講座、開幕茶會、記者會與聚會。\n可依活動形式、預算與份數，規劃適合的甜點組合及包裝。\n\n**洽詢時請提供**：活動日期、預算、份數、配送地點與包裝需求。", "", "LINE 洽詢", line, "TRUE", "30"],
      ["b4", "custom", "企業訂購", "Corporate Gifting", "適用於客戶贈禮、員工節慶禮盒、品牌合作與商務拜訪。\n可依採購數量、單份預算、交期及品牌需求，討論禮盒搭配、數量級距與包裝客製的可行性。", "", "LINE 洽詢", line, "TRUE", "40"]
    ].forEach(r => blocks.appendRow(r));
  }
}

function lilouIsOn_(v) { return !/^(false|0|no|n|否|下架|隱藏)$/i.test((v === undefined || v === null ? "TRUE" : v).toString().trim() || "TRUE"); }

function readSheetObjects_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => h.toString().trim());
  return values.slice(1).filter(r => r.some(v => v.toString().trim() !== "")).map(r => {
    const o = {}; headers.forEach((h, i) => o[h] = (r[i] || "").toString()); return o;
  });
}

// 前台用：只輸出上架的頁面與區塊
function getPublicPages_(ss) {
  const blocks = readSheetObjects_(ss, "PageBlocks").filter(b => lilouIsOn_(b.display));
  return readSheetObjects_(ss, "Pages").filter(p => p.page_id && lilouIsOn_(p.display)).map(p => ({
    id: p.page_id.trim(), nav_label: p.nav_label || p.title, nav_label_en: p.nav_label_en || "", title: p.title || p.nav_label, title_en: p.title_en || "", intro: p.intro || "", sort: Number(p.sort) || 999,
    blocks: blocks.filter(b => (b.page_id || "").trim() === p.page_id.trim()).sort((a, b) => (Number(a.sort) || 999) - (Number(b.sort) || 999))
      .map(b => ({ title: b.title, title_en: b.title_en, body: b.body, image: b.image, button_text: b.button_text, button_link: b.button_link }))
  })).sort((a, b) => a.sort - b.sort);
}

// 自動建立活動宣傳設定
function ensureSettingsPromo_(ss) {
  let sheet = ss.getSheetByName("Settings");
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const keys = data.map(r => r[0].toString().trim());
  const promoDefaults = [
    ["promo_enabled", "TRUE"],
    ["promo_eyebrow", "中秋限定禮盒"],
    ["promo_title", "把團圓的心意裝進一盒甜點"],
    ["promo_desc", "限量製作，預購至 9 月 20 日，售完為止。"],
    ["promo_image", "images/seasonal-gift-box.jpg"],
    ["promo_btn_text", "立即預購"],
    ["promo_btn_link", "giftbox"],
    // 首頁小圖輪播（逗號分隔的圖片網址，可在後台上傳到雲端硬碟）
    ["gallery_images", [1, 2, 3, 4, 5, 6].map(n => "https://www.lilou.food/images/gift-" + n + ".jpg").join(",")],
    // 生日蛋糕頁的文字與店內慶生時段
    ["cake_intro", "為重要的日子，準備一份剛剛好的甜。"],
    ["cake_notice", "預訂時間｜生日蛋糕皆需提前 3 天預訂，最快可取貨日為下單日期起算第 3 天，實際取貨日期請依預約時段為準。\n壽星限定招待｜預訂生日蛋糕並預約於店內慶生，即招待當月壽星飲品 1 杯。\n水果備註｜若有不食用或需避開的水果，請於訂單備註欄註明，我們將盡可能協助調整。"],
    ["cake_outro", "謝謝你將重要的日子交給粒露，願每一份蛋糕都為日常留下值得紀念的甜。"],
    ["celebration_time_slots", "14:00,15:00,16:00,17:00"]
  ];
  promoDefaults.forEach(def => {
     if (!keys.includes(def[0])) { sheet.appendRow(def); }
  });
}

function getFormOptions() {
  const ss = SpreadsheetApp.openById(GLOBAL_SPREADSHEET_ID);

  ensureCategoriesSheet_(ss);
  ensureSettingsPromo_(ss);
  ensureSettingsDateRules_(ss);
  ensurePagesSheets_(ss);
  ensureCakeSetup_(ss);

  const catSheet = ss.getSheetByName("Categories") || ss.getSheetByName("Categories 分類管理");
  const prodSheet = ss.getSheetByName("Products");
  const packSheet = ss.getSheetByName("Packaging");
  const couponSheet = ss.getSheetByName("Coupons");
  const settingsSheet = ss.getSheetByName("Settings");

  if (!prodSheet) throw new Error("找不到 Products 分頁");
  if (!packSheet) throw new Error("找不到 Packaging 分頁");

  const catDisplay = catSheet.getDataRange().getDisplayValues();
  const prodDisplay = prodSheet.getDataRange().getDisplayValues();
  const packDisplay = packSheet.getDataRange().getDisplayValues();
  const couponDisplay = couponSheet ? couponSheet.getDataRange().getDisplayValues() : [];
  const settingsData = settingsSheet ? settingsSheet.getDataRange().getDisplayValues() : [];

  const catHeaders = catDisplay.length ? catDisplay[0].map(h => h.toString().trim().toLowerCase()) : [];
  const categories = [];
  for(let i=1; i<catDisplay.length; i++) {
    const row = catDisplay[i];
    const id = row[catHeaders.indexOf('id')] || '';
    if(!id) continue;
    const cat = {
      id: id.toString().trim(),
      name: row[catHeaders.indexOf('name')] || id,
      behavior: (row[catHeaders.indexOf('behavior')] || 'combo').toString().trim().toLowerCase(),
      display: (row[catHeaders.indexOf('display')] || 'TRUE').toString().toUpperCase() !== 'FALSE',
      sort: Number(row[catHeaders.indexOf('sort')]) || (i * 10)
    };
    CATEGORY_RULE_HEADERS.forEach(h => { const idx = catHeaders.indexOf(h); cat[h] = idx >= 0 ? (row[idx] || "").toString().trim() : ""; });
    categories.push(cat);
  }
  categories.sort((a,b) => a.sort - b.sort);

  const prodHeaders = prodDisplay.length ? prodDisplay[0].map(h => h.toString().trim()) : [];
  const prodNameIdx = findHeaderIndex_(prodHeaders, ["name", "Name", "Product_Name", "商品名稱"]);
  const prodPriceIdx = findHeaderIndex_(prodHeaders, ["price", "Price", "售價", "價格"]);
  const prodStockIdx = findHeaderIndex_(prodHeaders, ["stock", "Stock", "庫存"]);
  const prodTypeIdx = findHeaderIndex_(prodHeaders, ["type", "Type", "商品類型"]);
  const prodDisplayIdx = findHeaderIndex_(prodHeaders, ["display", "Display", "顯示", "是否顯示"]);
  const prodShippingIdx = findHeaderIndex_(prodHeaders, ["shipping_rule", "shipping", "配送規則"]);
  const prodSortIdx = findHeaderIndex_(prodHeaders, ["sort", "Sort", "排序"]);
  const prodDescriptionIdx = findHeaderIndex_(prodHeaders, ["description", "商品描述", "商品說明"]);
  const prodIngredientsIdx = findHeaderIndex_(prodHeaders, ["ingredients", "成分"]);
  const prodAllergensIdx = findHeaderIndex_(prodHeaders, ["allergens", "過敏原"]);
  const prodStorageIdx = findHeaderIndex_(prodHeaders, ["storage", "保存方式"]);
  const prodWeightIdx = findHeaderIndex_(prodHeaders, ["weight", "重量"]);
  const prodNutritionIdx = findHeaderIndex_(prodHeaders, ["nutrition", "營養標示"]);
  const prodOriginManufacturerIdx = findHeaderIndex_(prodHeaders, ["origin_manufacturer", "originManufacturer", "原產地與製造商"]);
  const prodNameEnIdx = findHeaderIndex_(prodHeaders, ["name_en", "英文名稱"]);
  const prodImageIdx = findHeaderIndex_(prodHeaders, ["image", "圖片"]);
  const prodMinIdx = findHeaderIndex_(prodHeaders, ["min_per_order", "每筆最少"]);
  const prodMaxIdx = findHeaderIndex_(prodHeaders, ["max_per_order", "每筆最多"]);

  const products = [];
  for (let i = 1; i < prodDisplay.length; i++) {
    const row = prodDisplay[i];
    const name = prodNameIdx >= 0 ? row[prodNameIdx].toString().trim() : (row[0] || "").toString().trim();
    if (!name) continue;

    const displayRaw = prodDisplayIdx >= 0 ? row[prodDisplayIdx].toString().trim() : "TRUE";
    if (!(/^(true|1|yes|y|顯示|上架|是)$/i.test(displayRaw || "TRUE"))) continue;

    products.push({
      name: name,
      price: Number((prodPriceIdx >= 0 ? row[prodPriceIdx] : row[1]).toString().replace(/[^0-9.-]/g, "")) || 0,
      stock: Number((prodStockIdx >= 0 ? row[prodStockIdx] : row[2]).toString().replace(/[^0-9.-]/g, "")) || 0,
      type: prodTypeIdx >= 0 ? (row[prodTypeIdx] || "combo").toString().trim().toLowerCase() : "combo",
      display: true,
      shipping_rule: prodShippingIdx >= 0 ? (row[prodShippingIdx] || "all").toString().trim().toLowerCase() : "all",
      sort: Number((prodSortIdx >= 0 ? row[prodSortIdx] : "").toString().replace(/[^0-9.-]/g, "")) || (i * 10),
      description: prodDescriptionIdx >= 0 ? row[prodDescriptionIdx].toString().trim() : "",
      ingredients: prodIngredientsIdx >= 0 ? row[prodIngredientsIdx].toString().trim() : "",
      allergens: prodAllergensIdx >= 0 ? row[prodAllergensIdx].toString().trim() : "",
      storage: prodStorageIdx >= 0 ? row[prodStorageIdx].toString().trim() : "",
      weight: prodWeightIdx >= 0 ? row[prodWeightIdx].toString().trim() : "",
      nutrition: prodNutritionIdx >= 0 ? row[prodNutritionIdx].toString().trim() : "",
      origin_manufacturer: prodOriginManufacturerIdx >= 0 ? row[prodOriginManufacturerIdx].toString().trim() : "",
      name_en: prodNameEnIdx >= 0 ? row[prodNameEnIdx].toString().trim() : "",
      image: prodImageIdx >= 0 ? row[prodImageIdx].toString().trim() : "",
      min_per_order: prodMinIdx >= 0 ? (lilouNum_(row[prodMinIdx]) || 0) : 0,
      max_per_order: prodMaxIdx >= 0 ? (lilouNum_(row[prodMaxIdx]) || 0) : 0
    });
  }

  const packHeaders = packDisplay.length ? packDisplay[0].map(h => h.toString().trim()) : [];
  const packNameIdx = findHeaderIndex_(packHeaders, ["name", "Name", "包裝名稱"]);
  const packExtraIdx = findHeaderIndex_(packHeaders, ["extra_price", "Extra_Price", "加價", "包裝加價"]);
  const packBoxSizeIdx = findHeaderIndex_(packHeaders, ["box_size", "Box_Size", "容量", "入數"]);
  const packTypeIdx = findHeaderIndex_(packHeaders, ["type", "Type", "包裝類型"]);
  const packDisplayIdx = findHeaderIndex_(packHeaders, ["display", "Display", "顯示", "是否顯示"]);
  const packSortIdx = findHeaderIndex_(packHeaders, ["sort", "Sort", "排序"]);
  const packNameEnIdx = findHeaderIndex_(packHeaders, ["name_en", "英文名稱"]);

  const packaging = [];
  for (let i = 1; i < packDisplay.length; i++) {
    const row = packDisplay[i];
    const name = packNameIdx >= 0 ? row[packNameIdx].toString().trim() : (row[0] || "").toString().trim();
    if (!name) continue;
    const displayRaw = packDisplayIdx >= 0 ? row[packDisplayIdx].toString().trim() : "TRUE";
    if (!(/^(true|1|yes|y|顯示|上架|是)$/i.test(displayRaw || "TRUE"))) continue;

    packaging.push({
      name: name,
      extra_price: Number((packExtraIdx >= 0 ? row[packExtraIdx] : row[1] || "0").toString().replace(/[^0-9.-]/g, "")) || 0,
      box_size: Number((packBoxSizeIdx >= 0 ? row[packBoxSizeIdx] : "0").toString().replace(/[^0-9.-]/g, "")) || 0,
      type: packTypeIdx >= 0 ? (row[packTypeIdx] || "combo").toString().trim().toLowerCase() : "combo",
      packaging_type: packTypeIdx >= 0 ? (row[packTypeIdx] || "combo").toString().trim().toLowerCase() : "combo",
      display: true,
      sort: Number((packSortIdx >= 0 ? row[packSortIdx] : "").toString().replace(/[^0-9.-]/g, "")) || (i * 10),
      name_en: packNameEnIdx >= 0 ? row[packNameEnIdx].toString().trim() : ""
    });
  }

  const couponHeaders = couponDisplay.length ? couponDisplay[0].map(h => h.toString().trim()) : [];
  const coupons = [];
  for (let i = 1; i < couponDisplay.length; i++) {
    const row = couponDisplay[i];
    const code = findHeaderIndex_(couponHeaders, ["code", "Code", "折扣碼"]) >= 0 ? row[findHeaderIndex_(couponHeaders, ["code", "Code", "折扣碼"])].toString().trim() : (row[0] || "").toString().trim();
    if (!code) continue;
    coupons.push({
      code: code.toLowerCase(),
      type: (findHeaderIndex_(couponHeaders, ["type", "類型"]) >= 0 ? row[findHeaderIndex_(couponHeaders, ["type", "類型"])] : row[1] || "").toString().trim().toLowerCase(),
      value: parseFloat(findHeaderIndex_(couponHeaders, ["value", "折扣值"]) >= 0 ? row[findHeaderIndex_(couponHeaders, ["value", "折扣值"])] : row[2]) || 0,
      min_spend: Number(findHeaderIndex_(couponHeaders, ["min_spend", "最低消費"]) >= 0 ? row[findHeaderIndex_(couponHeaders, ["min_spend", "最低消費"])] : row[3]) || 0,
      start_date: (findHeaderIndex_(couponHeaders, ["start_date", "開始日期"]) >= 0 ? row[findHeaderIndex_(couponHeaders, ["start_date", "開始日期"])] : (row[4]||"")).toString().trim(),
      end_date: (findHeaderIndex_(couponHeaders, ["end_date", "結束日期"]) >= 0 ? row[findHeaderIndex_(couponHeaders, ["end_date", "結束日期"])] : (row[5]||"")).toString().trim()
    });
  }

  const settings = {};
  const todayTime = getTaiwanTodayTime();
  let pickupDates = [], deliveryDates = [];

  for (let i = 1; i < settingsData.length; i++) {
    const rawKey = settingsData[i][0];
    const rawValue = settingsData[i][1];
    if (!rawKey || rawKey.toString().trim() === "") continue;

    const key = rawKey.toString().trim();
    const lowerKey = key.toLowerCase().trim();
    const value = rawValue ? rawValue.toString().trim() : "";

    // 日期規則的鍵直接存，不當成日期清單
    if (LILOU_RULE_SETTING_KEYS.indexOf(lowerKey) !== -1) { settings[lowerKey] = value; continue; }

    if (lowerKey === "pickup_available_dates" || lowerKey.includes("pickup") || lowerKey.includes("store") || lowerKey.includes("自取")) {
      // D+N 規則改由前台與送單驗證處理，這裡只排除已過去的日期
      pickupDates = pickupDates.concat(parseAvailableDates(value, todayTime));
      continue;
    }
    if (lowerKey === "delivery_available_dates" || lowerKey.includes("delivery") || lowerKey.includes("店到店") || lowerKey.includes("711") || lowerKey.includes("7-11")) {
      deliveryDates = deliveryDates.concat(parseAvailableDates(value, todayTime));
      continue;
    }
    settings[lowerKey] = value;
  }

  settings["pickup_available_dates"] = Array.from(new Set(pickupDates)).sort().join(",") || "目前無可選擇日期";
  settings["delivery_available_dates"] = Array.from(new Set(deliveryDates)).sort().join(",") || "目前無可選擇日期";
  Object.keys(LILOU_RULE_DEFAULTS).forEach(k => { if (settings[k] === undefined || settings[k] === "") settings[k] = LILOU_RULE_DEFAULTS[k]; });

  const pages = getPublicPages_(ss);
  const cake_options = readCakeOptions_(ss);
  const bookings = countCategoryBookings_(ss, categories, products);
  categories.forEach(c => { c.booked = bookings[c.id] || {}; });
  return { categories, products, packaging, coupons, settings, pages, cake_options };
}

function findHeaderIndex_(headers, names) {
  for (let i = 0; i < names.length; i++) {
    const idx = headers.indexOf(names[i]);
    if (idx !== -1) return idx;
  }
  return -1;
}

function normalizeCustomerStatus_(status) { return (status || "").toString().trim(); }

function queryOrders(phone) {
  if (!phone) return { status: "error", message: "請提供手機號碼。" };
  const searchPhone = phone.toString().replace(/[^0-9]/g, "").trim();
  // 安全性：只接受完整 10 碼手機號碼，避免輸入片段數字就查到其他客人的訂單
  if (!/^09\d{8}$/.test(searchPhone)) return { status: "error", message: "請輸入完整的 10 碼手機號碼。" };
  const ss = SpreadsheetApp.openById(GLOBAL_SPREADSHEET_ID);
  const orderSheet = ss.getSheetByName("Orders");
  if (!orderSheet) return { status: "error", message: "找不到 Orders 分頁" };

  // 加速：不再把整張 Orders 表讀進來。先只讀表頭和「電話」那一欄找出相符的列，再只讀那幾列
  const lastRow = orderSheet.getLastRow();
  const lastCol = orderSheet.getLastColumn();
  if (lastRow <= 1 || lastCol < 1) return { status: "success", data: [] };

  const headers = orderSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => h.toString().trim());
  const idxId = findHeaderIndex_(headers, ["ID", "Order_ID", "訂單編號"]);
  const idxTimestamp = findHeaderIndex_(headers, ["Timestamp", "Created_At", "下單日期"]);
  const idxName = findHeaderIndex_(headers, ["Name", "Customer_Name", "姓名"]);
  const idxPhone = findHeaderIndex_(headers, ["Phone", "Customer_Phone", "聯絡電話"]);
  const idxEmail = findHeaderIndex_(headers, ["Email", "E-mail"]);
  const idxDate = findHeaderIndex_(headers, ["Date", "Order_Date", "預定日期", "取貨日期"]);
  const idxMethod = findHeaderIndex_(headers, ["Shipping_Method", "Method", "配送方式"]);
  const idxAddress = findHeaderIndex_(headers, ["Address", "Store", "取貨門市", "地址"]);
  const idxDetails = findHeaderIndex_(headers, ["Product_Name", "Order_Details", "商品明細"]);
  const idxTotal = findHeaderIndex_(headers, ["Total_Price", "Total", "總金額"]);
  const idxNote = findHeaderIndex_(headers, ["Note", "備註"]);
  const idxStatus = findHeaderIndex_(headers, ["Order_Status", "Status", "訂單狀態"]);
  const idxPayment = findHeaderIndex_(headers, ["Payment_Link", "付款連結"]);
  const idxTracking = findHeaderIndex_(headers, ["Tracking_No", "貨運編號"]);
  const idxStatusNote = findHeaderIndex_(headers, ["Status_Note", "狀態備註"]);
  const idxLastUpdate = findHeaderIndex_(headers, ["Last_Update", "最後更新"]);
  const idxSubtotalAmount = findHeaderIndex_(headers, ["Subtotal_Amount", "商品小計"]);
  const idxDiscountAmount = findHeaderIndex_(headers, ["Discount_Amount", "折扣金額"]);
  const idxShippingFee = findHeaderIndex_(headers, ["Shipping_Fee", "運費"]);
  const idxItemsDetail = findHeaderIndex_(headers, ["Items_Detail", "商品單價明細"]);

  if (idxPhone === -1) return { status: "success", data: [] };

  const phoneCol = orderSheet.getRange(2, idxPhone + 1, lastRow - 1, 1).getDisplayValues();
  const matchedRowNums = [];
  for (let i = 0; i < phoneCol.length; i++) {
    if (phoneCol[i][0].toString().replace(/[^0-9]/g, "").trim() === searchPhone) matchedRowNums.push(i + 2);
  }
  if (!matchedRowNums.length) return { status: "success", data: [] };

  // 相符的列不多時逐列讀；較多時讀「第一筆到最後一筆」這一段再挑出來，減少讀取次數
  let rows;
  if (matchedRowNums.length <= 3) {
    rows = matchedRowNums.map(r => orderSheet.getRange(r, 1, 1, lastCol).getDisplayValues()[0]);
  } else {
    const firstRow = matchedRowNums[0];
    const lastMatchRow = matchedRowNums[matchedRowNums.length - 1];
    const block = orderSheet.getRange(firstRow, 1, lastMatchRow - firstRow + 1, lastCol).getDisplayValues();
    rows = matchedRowNums.map(r => block[r - firstRow]);
  }

  const matchedOrders = [];
  for (let i = 0; i < rows.length; i++) {
    matchedOrders.push({
        id: idxId !== -1 ? rows[i][idxId].toString().trim().replace(/^\'/, "") : "",
        orderDate: idxTimestamp !== -1 ? rows[i][idxTimestamp].toString().trim() : "",
        name: idxName !== -1 ? rows[i][idxName].toString().trim() : "",
        email: idxEmail !== -1 ? rows[i][idxEmail].toString().trim() : "",
        date: idxDate !== -1 ? rows[i][idxDate].toString().trim() : "",
        method: idxMethod !== -1 ? rows[i][idxMethod].toString().trim() : "",
        shipping_method: idxMethod !== -1 ? rows[i][idxMethod].toString().trim() : "",
        address: idxAddress !== -1 ? rows[i][idxAddress].toString().trim() : "",
        details: idxDetails !== -1 ? rows[i][idxDetails].toString().trim() : "",
        total: idxTotal !== -1 ? rows[i][idxTotal].toString().trim() : "0",
        note: idxNote !== -1 ? rows[i][idxNote].toString().trim() : "",
        status: normalizeCustomerStatus_(idxStatus !== -1 ? rows[i][idxStatus].toString().trim() : "訂單確認中"),
        paymentLink: idxPayment !== -1 ? rows[i][idxPayment].toString().trim() : "",
        trackingNo: idxTracking !== -1 ? rows[i][idxTracking].toString().trim() : "",
        statusNote: idxStatusNote !== -1 ? rows[i][idxStatusNote].toString().trim() : "",
        lastUpdate: idxLastUpdate !== -1 ? rows[i][idxLastUpdate].toString().trim() : "",
        subtotalAmount: idxSubtotalAmount !== -1 ? rows[i][idxSubtotalAmount].toString().trim() : "",
        discountAmount: idxDiscountAmount !== -1 ? rows[i][idxDiscountAmount].toString().trim() : "",
        shippingFee: idxShippingFee !== -1 ? rows[i][idxShippingFee].toString().trim() : "",
        itemsDetail: idxItemsDetail !== -1 ? rows[i][idxItemsDetail].toString().trim() : ""
      });
  }
  matchedOrders.reverse();
  return { status: "success", data: matchedOrders };
}

function doGet(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : "";
    if (action === "getFormOptions" || action === "getData") {
      const cache = CacheService.getScriptCache();
      const cached = cache.get(FORM_OPTIONS_CACHE_KEY);
      if (cached) return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
      const response = getFormOptions();
      const json = JSON.stringify(response);
      cache.put(FORM_OPTIONS_CACHE_KEY, json, FORM_OPTIONS_CACHE_SECONDS);
      return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
    }
    if (action === "queryOrders") return jsonOutput(queryOrders(e.parameter.phone));
    // 健康檢查：回報目前執行的程式版本、專案 ID、雲端硬碟權限（不含任何客人資料）
    // 健康檢查：回報目前執行的程式版本與專案 ID（不含任何客人資料）
    if (action === "diag") {
      let github = "ok";
      try { const c = githubConfig_(); github = "ok（" + c.repo + "／" + c.branch + "）"; } catch (err) { github = err.message; }
      return jsonOutput({ status: "success", codeVersion: "2026-11-08-orderdate", scriptId: ScriptApp.getScriptId(), github: github });
    }
    if (action === "adminLogin") return jsonOutput({ status: "error", message: "後台已更新，請重新整理後台頁面再登入。" });
    if (action === "adminGetAll") return jsonOutput(adminGetAll_(e.parameter.token));
    return jsonOutput({ status: "error", message: "無效的 action，請使用 getFormOptions 或 queryOrders。" });
  } catch (error) {
    return jsonOutput({ status: "error", message: error.toString() });
  }
}

function generateLilouOrderId_(orderSheet) {
  const today = Utilities.formatDate(new Date(), "GMT+8", "yyMMdd");
  const prefix = "Li-" + today + "-";
  const lastRow = orderSheet.getLastRow();
  if (lastRow <= 1) return prefix + "001";
  const ids = orderSheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat().filter(String);
  let maxSeq = 0;
  ids.forEach(id => {
    const idText = id.toString().trim();
    if (idText.indexOf(prefix) === 0) {
      const seq = parseInt(idText.replace(prefix, ""), 10);
      if (!isNaN(seq)) maxSeq = Math.max(maxSeq, seq);
    }
  });
  return prefix + String(maxSeq + 1).padStart(3, "0");
}

function sendLinePushMessage_(text) {
  if (!LINE_CHANNEL_ACCESS_TOKEN || !LINE_ADMIN_USER_ID) return;
  const options = {
    method: "post", contentType: "application/json",
    headers: { Authorization: "Bearer " + LINE_CHANNEL_ACCESS_TOKEN },
    payload: JSON.stringify({ to: LINE_ADMIN_USER_ID, messages: [{ type: "text", text: text }] }),
    muteHttpExceptions: true
  };
  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", options);
}

function isStorePickupMethod_(method) { return (method || "").toString().trim() === "Store Pickup"; }
function isSevenPickupCodMethod_(method) { return (method || "").toString().trim().toLowerCase().includes("cod"); }
function isSevenPickupMethod_(method) { return (method || "").toString().trim().toLowerCase().includes("7-11") || (method || "").toString().trim().toLowerCase().includes("711"); }
function isPrepaidSevenPickupMethod_(method) { return isSevenPickupMethod_(method) && !isSevenPickupCodMethod_(method); }

function getShippingMethodLabel_(method) {
  if (isStorePickupMethod_(method)) return "門市自取付款";
  if (isSevenPickupCodMethod_(method)) return "7-11 取貨付款";
  if (isSevenPickupMethod_(method)) return "7-11 店到店取貨";
  return method || "";
}

function buildNewOrderLineMessage_(orderId, data, finalAddress, timestamp) {
  return "🍮 Lilou 新訂單通知\n\n訂單編號：" + orderId + "\n成立時間：" + timestamp + "\n\n姓名：" + (data.name || "") + "\n電話：" + (data.phone || "") + "\nEmail：" + (data.email || "") + "\n\n配送方式：" + getShippingMethodLabel_(data.shipping_method) + "\n預定日期：" + (data.date || "") + "\n取貨／收件資訊：" + (finalAddress || "") + "\n\n訂單金額：$" + (data.total_price || "") + "\n折扣碼：" + (data.coupon_code || "無") + "\n\n商品明細：\n" + (data.product_name || "") + "\n\n查詢訂單：\n" + ORDER_LOOKUP_URL;
}

function sendOrderCreatedEmail_(data, orderId, finalAddress, timestamp) {
  const email = (data.email || "").toString().trim();
  if (!email) return;
  const isPrepaidDelivery = isPrepaidSevenPickupMethod_(data.shipping_method);
  let body = "您好，感謝您的訂購！\n\n我們已收到您的訂單，目前正在確認內容與庫存，完成後將更新訂單狀態。\n\n訂單編號：" + orderId + "\n訂單狀態：訂單確認中\n\n";
  if (isPrepaidDelivery) body += "若後續訂單狀態更新為「待付款」，請依付款通知完成付款。\n若未收到付款通知，也可至「查詢訂單」查看最新進度。\n\n";
  body += "您可隨時至「查詢訂單」查看最新進度：\n" + ORDER_LOOKUP_URL + "\n\n" + LILOU_EMAIL_AUTO_NOTICE + "\n\nLilou｜迷你可麗露專賣";
  GmailApp.sendEmail(email, "Lilou｜您的訂單已收到", body, { from: LILOU_EMAIL_FROM, name: LILOU_EMAIL_NAME, replyTo: LILOU_EMAIL_REPLY_TO });
}

function doPost(e) {
  // 後台登入不需要排隊鎖定，先處理（密碼改由 POST 傳送，不會出現在網址上）
  try {
    if (e && e.postData && e.postData.contents) {
      const pre = JSON.parse(e.postData.contents);
      if (pre && pre.action === "adminLogin") return jsonOutput(adminLogin_(pre.password));
    }
  } catch (preErr) {}

  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    lock.waitLock(15000);
    locked = true;
    if (!e || !e.postData || !e.postData.contents) return jsonOutput({ status: "error", message: "未收到訂單資料。" });

    const data = JSON.parse(e.postData.contents);
    if (data.action === "adminUpdateOrder") return jsonOutput(adminUpdateOrder_(data));
    if (data.action === "adminReplaceSheetRows") return jsonOutput(withSnapshot_(adminReplaceSheetRows_(data.token, data.sheetName, data.rows)));
    if (data.action === "adminPublishSnapshot") { adminRequire_(data.token); return jsonOutput({ status: "success", snapshot: publishSiteSnapshot_() }); }
    if (data.action === "adminUploadImage") return jsonOutput(adminUploadImageToGithub_(data));
    if (data.action === "adminSaveProduct") return jsonOutput(withSnapshot_(adminSaveProduct_(data)));
    if (data.action === "adminSaveProductOrder") return jsonOutput(withSnapshot_(adminSaveProductOrder_(data)));
    if (data.action === "adminDeleteProduct") return jsonOutput(withSnapshot_(adminDeleteProduct_(data)));
    if (data.action === "adminSaveRow") return jsonOutput(withSnapshot_(adminSaveRow_(data)));
    if (data.action === "adminDeleteRow") return jsonOutput(withSnapshot_(adminDeleteRow_(data)));
    if (data.action === "adminSaveRowOrder") return jsonOutput(withSnapshot_(adminSaveRowOrder_(data)));
    if (data.action !== "submitOrder") return jsonOutput({ status: "error", message: "無效的 action。" });

    const ss = SpreadsheetApp.openById(GLOBAL_SPREADSHEET_ID);
    const orderSheet = ss.getSheetByName("Orders");
    if (!orderSheet) return jsonOutput({ status: "error", message: "找不到 Orders 分頁" });

    // 第2階段：用與前台相同的規則再檢查一次日期與限購（避免繞過網頁直接送單）
    const formOpts = getFormOptions();
    const ruleError = validateOrderRules_(ss, data, formOpts);
    if (ruleError) return jsonOutput({ status: "error", message: ruleError });
    // 資安：用試算表上的售價重算金額，和客人送來的不一致就不成立；商品明細也改用後端重建的內容
    const priced = lilouVerifyOrderPricing_(formOpts, data);
    if (priced.error) return jsonOutput({ status: "error", message: priced.error });

    const orderId = generateLilouOrderId_(orderSheet);
    const timestamp = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm:ss");
    const shippingMethod = data.shipping_method || "";
    const finalAddress = isStorePickupMethod_(shippingMethod) ? STORE_ADDRESS : (data.address || "");

    appendOrderByHeader_(orderSheet, {
      ID: orderId, Timestamp: timestamp, Name: sheetSafeText_(data.name), Phone: "'" + (data.phone || "").toString().replace(/[^0-9]/g, ""), Email: sheetSafeText_(data.email),
      Product_Name: sheetSafeText_(priced.productName), Packaging_Name: sheetSafeText_(priced.packagingName), Quantity: priced.quantity,
      Total_Price: priced.total, Subtotal_Amount: priced.subtotal, Discount_Amount: priced.discount,
      Shipping_Fee: priced.shippingFee, Shipping_Method: sheetSafeText_(shippingMethod), Date: sheetSafeText_(data.date), Address: sheetSafeText_(finalAddress),
      Note: sheetSafeText_(data.note), Coupon_Code: sheetSafeText_(priced.couponCode), Order_Status: "訂單確認中",
      Payment_Link: "", Tracking_No: "", Status_Note: "", Internal_Note: "", Last_Update: timestamp, Items_Detail: sheetSafeText_(data.items_detail)
    });

    updateProductStock(ss, priced.productName);
    data.product_name = priced.productName; data.total_price = priced.total;
    sendLinePushMessage_(buildNewOrderLineMessage_(orderId, data, finalAddress, timestamp));
    sendOrderCreatedEmail_(data, orderId, finalAddress, timestamp);

    return jsonOutput({ status: "success", orderId: orderId, orderDate: timestamp });
  } catch (error) {
    return jsonOutput({ status: "error", message: error.toString() });
  } finally {
    if (locked) lock.releaseLock();
  }
}

// 送單驗證：日期是否在可選範圍、每筆限購。回傳錯誤訊息，沒問題回傳空字串
function validateOrderRules_(ss, data, opts) {
  let groups;
  try { groups = JSON.parse(data.items_detail || "[]"); } catch (e) { return "訂單內容有誤，請重新整理頁面後再下單。"; }
  if (!Array.isArray(groups) || !groups.length) return "訂單內容有誤，請重新整理頁面後再下單。";

  opts = opts || getFormOptions();
  const qtyByName = {};
  groups.forEach(g => (g.items || []).forEach(it => {
    const name = (it.name || "").toString().trim();
    if (name) qtyByName[name] = (qtyByName[name] || 0) + (Number(it.qty) || 0);
  }));

  // 限購
  for (const name in qtyByName) {
    const p = opts.products.find(x => x.name === name);
    if (!p) continue;
    if (p.max_per_order > 0 && qtyByName[name] > p.max_per_order) return "「" + name + "」每筆訂單最多 " + p.max_per_order + " 個，請調整數量。";
    if (p.min_per_order > 0 && qtyByName[name] < p.min_per_order) return "「" + name + "」每筆訂單最少 " + p.min_per_order + " 個，請調整數量。";
  }

  // 日期
  const date = normalizeDateString((data.date || "").toString().slice(0, 10));
  if (!date) return "請選擇取貨／到貨日期。";
  const kind = isStorePickupMethod_(data.shipping_method) ? "pickup" : "delivery";
  const catIds = {};
  Object.keys(qtyByName).forEach(name => { const p = opts.products.find(x => x.name === name); if (p) catIds[p.type] = true; });
  const cats = opts.categories.filter(c => catIds[c.id]);
  const openDates = (opts.settings[kind + "_available_dates"] || "").split(",");
  const allowed = lilouAllowedDates_(kind, openDates, opts.settings, cats, lilouOrderBaseDate_(opts.settings.order_cutoff_time));
  if (allowed.indexOf(date) === -1) return "您選的日期 " + date + " 目前已無法預約，請重新整理頁面後再選擇日期。";
  // 每日限量
  for (const c of cats) {
    const limit = Number(c.daily_limit) || 0;
    if (!limit) continue;
    let qtyInOrder = 0;
    Object.keys(qtyByName).forEach(name => { const p = opts.products.find(x => x.name === name); if (p && p.type === c.id) qtyInOrder += qtyByName[name]; });
    const booked = (c.booked && c.booked[date]) || 0;
    if (booked + qtyInOrder > limit) return "「" + c.name + "」" + date + " 已額滿（每日限量 " + limit + "，剩 " + Math.max(0, limit - booked) + "），請改選其他日期。";
  }
  return "";
}

// ===== 後端金額重算（資安）=====
// 規則必須和前台 index.html 相同：組合 =（口味單價×數量＋包裝加價）×組數；單品 = 單價×數量；
// 生日蛋糕 = 蛋糕售價＋蛋糕體＋2 款夾餡＋配件；宅配加運費；優惠碼依試算表設定。
const LILOU_SHIPPING_FEE = 129; // 預設運費；實際以後台「設定」的 shipping_fee 為準，規則和前台 lilouApplyShippingFee 相同
function lilouShippingFee_(settings) {
  const v = settings && settings.shipping_fee !== undefined && settings.shipping_fee !== null ? settings.shipping_fee.toString().trim() : "";
  const n = Number(v);
  return v !== "" && isFinite(n) && n >= 0 ? Math.round(n) : LILOU_SHIPPING_FEE;
}
const LILOU_PRICE_CHANGED_MSG = "商品價格或優惠內容已更新，請重新整理頁面後再下單。";

function lilouVerifyOrderPricing_(opts, data) {
  let groups;
  try { groups = JSON.parse(data.items_detail || "[]"); } catch (e) { return { error: "訂單內容有誤，請重新整理頁面後再下單。" }; }
  if (!Array.isArray(groups) || !groups.length) return { error: "訂單內容有誤，請重新整理頁面後再下單。" };

  const catById = {}; (opts.categories || []).forEach(c => catById[c.id] = c);
  const prodByName = {}; (opts.products || []).forEach(p => prodByName[p.name] = p);
  const packByName = {}; (opts.packaging || []).forEach(p => packByName[p.name] = p);
  const cakeOpts = opts.cake_options || [];
  const qtyByName = {};
  let subtotal = 0, totalQty = 0;
  const productParts = [], packParts = [];

  for (let gi = 0; gi < groups.length; gi++) {
    const g = groups[gi] || {};
    const repeat = Math.max(1, Math.floor(Number(g.repeatQty) || 1));
    if (repeat > 99) return { error: "訂單內容有誤，請重新整理頁面後再下單。" };
    const items = Array.isArray(g.items) ? g.items : [];
    if (!items.length) return { error: "訂單內容有誤，請重新整理頁面後再下單。" };
    const boxName = (g.packageName || "").toString().trim();
    let groupUnitCost = 0, isCombo = false, isCake = false, perBoxCount = 0;
    const flavorParts = [];
    let optionText = "";

    for (const it of items) {
      const name = (it.name || "").toString().trim();
      const p = prodByName[name];
      if (!p) return { error: "「" + name + "」目前已下架或名稱已變更，請重新整理頁面後再下單。" };
      const qty = Math.floor(Number(it.qty) || 0);
      if (qty <= 0 || qty % repeat !== 0) return { error: "訂單內容有誤，請重新整理頁面後再下單。" };
      const perBox = qty / repeat;
      const behavior = ((catById[p.type] || {}).behavior || "combo").toString().toLowerCase();
      if (behavior === "combo") isCombo = true;
      let unit = Number(p.price) || 0;
      if (behavior === "cake") {
        isCake = true;
        const cake = lilouCakeExtra_(cakeOpts, g.cake, opts.settings || {});
        if (cake.error) return { error: cake.error };
        unit += cake.extra;
        optionText = cake.text;
      }
      groupUnitCost += unit * perBox;
      perBoxCount += perBox;
      qtyByName[name] = (qtyByName[name] || 0) + qty;
      flavorParts.push(repeat > 1 ? name + "(" + qty + "個：" + perBox + "個×" + repeat + "組)" : name + "(" + perBox + "個)");
    }

    // 包裝加價：組合要對到包裝名稱；單品與蛋糕前台不加價
    let boxExtra = 0;
    if (isCombo) {
      const pack = packByName[boxName];
      if (!pack) return { error: "包裝「" + boxName + "」目前已下架或名稱已變更，請重新整理頁面後再下單。" };
      if (pack.box_size > 0 && perBoxCount !== pack.box_size) return { error: "「" + boxName + "」必須剛好裝 " + pack.box_size + " 個，請重新整理頁面後再下單。" };
      boxExtra = Number(pack.extra_price) || 0;
    }
    subtotal += (groupUnitCost + boxExtra) * repeat;
    totalQty += perBoxCount * repeat;
    productParts.push("【第" + (gi + 1) + "組: " + boxName + (repeat > 1 ? " ×" + repeat + "組" : "") + "】" + flavorParts.join("、") + (isCake && optionText ? "［" + optionText + "］" : ""));
    packParts.push("[組" + (gi + 1) + "]" + boxName + (repeat > 1 ? "×" + repeat + "組" : ""));
  }

  // 庫存（前台已檢查，這裡防止繞過網頁）
  for (const name in qtyByName) {
    const stock = Number(prodByName[name].stock) || 0;
    if (qtyByName[name] > stock) return { error: "「" + name + "」庫存不足（剩 " + Math.max(0, stock) + " 個），請調整數量。" };
  }

  const shippingFee = isStorePickupMethod_(data.shipping_method) ? 0 : lilouShippingFee_(opts.settings);
  let discount = 0, couponCode = "";
  const code = (data.coupon_code || "").toString().trim().toLowerCase();
  if (code) {
    const c = (opts.coupons || []).find(x => (x.code || "").toString().trim().toLowerCase() === code);
    if (!c || !lilouCouponActive_(c)) return { error: "折扣碼「" + code + "」無效或已過期，請重新整理頁面後再下單。" };
    if (subtotal > 0 && subtotal >= (Number(c.min_spend) || 0)) {
      if (c.type === "minus") discount = Number(c.value) || 0;
      else if (c.type === "discount") discount = Math.round(subtotal * (1 - Number(c.value)));
      if (discount > subtotal + shippingFee) discount = subtotal + shippingFee;
      couponCode = code;
    }
  }
  const total = subtotal + shippingFee - discount;

  const sent = { subtotal: Number(data.subtotal_amount), shipping: Number(data.shipping_fee), discount: Number(data.discount_amount), total: Number(data.total_price) };
  if (sent.total !== total || sent.subtotal !== subtotal || sent.shipping !== shippingFee || sent.discount !== discount) {
    console.warn("金額不符", JSON.stringify({ sent: sent, server: { subtotal: subtotal, shipping: shippingFee, discount: discount, total: total } }));
    return { error: LILOU_PRICE_CHANGED_MSG };
  }
  return { subtotal: subtotal, shippingFee: shippingFee, discount: discount, total: total, couponCode: couponCode, quantity: totalQty,
    productName: productParts.join(" ｜ "), packagingName: packParts.join("、") };
}

// 生日蛋糕加價：蛋糕體 1 款、夾餡 2 款、配件；文字格式和前台 cakeOptionText 相同
function lilouCakeExtra_(cakeOpts, cake, settings) {
  const bad = { error: "生日蛋糕的選項有誤，請重新整理頁面後再下單。" };
  if (!cake || typeof cake !== "object") return bad;
  const find = (group, name) => cakeOpts.find(o => o.group === group && o.name === (name || "").toString());
  const base = find("base", cake.base), f1 = find("filling", cake.fill1), f2 = find("filling", cake.fill2);
  if (!base || !f1 || !f2) return bad;
  const plus = v => v > 0 ? "＋$" + v : "";
  let extra = base.price + f1.price + f2.price;
  const parts = ["蛋糕體：" + base.name + plus(base.price), "夾餡：" + f1.name + plus(f1.price) + "＋" + f2.name + plus(f2.price)];
  const addons = cake.addons && typeof cake.addons === "object" ? cake.addons : {};
  for (const name in addons) { if (!find("addon", name)) return bad; }
  const addText = [];
  cakeOpts.filter(o => o.group === "addon").forEach(o => {
    const q = Math.floor(Number(addons[o.name]) || 0);
    if (q < 0 || (o.max_qty > 0 && q > o.max_qty)) { extra = NaN; return; }
    if (!q) return;
    if (o.input === "qty") {
      const over = Math.max(0, q - (o.included_qty || 0));
      extra += over * o.price;
      addText.push(o.name + "×" + q + (over ? "（加購 " + over + " 組＋$" + (over * o.price) + "）" : ""));
    } else {
      extra += o.price;
      addText.push(o.name + (o.max_qty === 1 && /蠟燭/.test(o.name) ? "×1" : "") + plus(o.price));
    }
  });
  if (isNaN(extra)) return bad;
  if (addText.length) parts.push("配件：" + addText.join("、"));
  if (cake.celebrate) {
    const people = Math.floor(Number(cake.people) || 0);
    if (!cake.slot || people < 1 || people > 50) return { error: "預約店內慶生請選擇到店時段並填寫人數。" };
    parts.push("店內慶生：" + cake.slot.toString().slice(0, 20) + "・" + people + " 人");
  }
  return { extra: extra, text: parts.join("／") };
}

function lilouCouponActive_(c) {
  const today = Utilities.formatDate(new Date(), TIMEZONE, "yyyy/MM/dd");
  const start = normalizeDateString(c.start_date), end = normalizeDateString(c.end_date);
  if (start && today < start) return false;
  if (end && today > end) return false;
  return true;
}

function updateProductStock(ss, orderDetails) {
  try {
    const prodSheet = ss.getSheetByName("Products");
    if (!prodSheet || !orderDetails) return;
    const prodData = prodSheet.getDataRange().getDisplayValues();
    if (prodData.length <= 1) return;

    const headers = prodData[0].map(h => h.toString().trim());
    const nameCol0 = findHeaderIndex_(headers, ["name", "Name", "Product_Name", "商品名稱"]);
    const stockCol0 = findHeaderIndex_(headers, ["stock", "Stock", "庫存"]);
    if (nameCol0 < 0 || stockCol0 < 0) return;

    const lowStockProducts = [];
    const soldOutProducts = [];

    for (let i = 1; i < prodData.length; i++) {
      const productName = (prodData[i][nameCol0] || "").toString().trim();
      if (!productName) continue;
      const escapedName = productName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const regex = new RegExp(escapedName + "\\((\\d+)[^)]*\\)", "g");
      let match, totalPurchasedQty = 0;
      while ((match = regex.exec(orderDetails)) !== null) totalPurchasedQty += parseInt(match[1], 10);

      if (totalPurchasedQty > 0) {
        const currentStock = Number((prodData[i][stockCol0] || "").toString().replace(/[^0-9.-]/g, "")) || 0;
        const newStock = Math.max(0, currentStock - totalPurchasedQty);
        prodSheet.getRange(i + 1, stockCol0 + 1).setValue(newStock);
        if (currentStock > 0 && newStock === 0) soldOutProducts.push(productName);
        else if (currentStock > 5 && newStock <= 5) lowStockProducts.push({ name: productName, stock: newStock });
      }
    }

    if (lowStockProducts.length > 0) sendLinePushMessage_("⚠️【Lilou 低庫存提醒】\n\n以下商品庫存已剩 5 顆以下：\n" + lowStockProducts.map(item => "・" + item.name + "：剩 " + item.stock + " 顆").join("\n"));
    if (soldOutProducts.length > 0) sendLinePushMessage_("【Lilou 完售通知】\n\n" + soldOutProducts.map(name => "・" + name + " 已完售").join("\n"));
  } catch (error) {}
}

function getHeaderIndexMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => map[h.toString().trim()] = i + 1);
  return map;
}

function ensureOrderHeaders_(sheet) {
  const ORDER_HEADERS_REQUIRED_ = ["ID", "Timestamp", "Name", "Phone", "Email", "Product_Name", "Packaging_Name", "Quantity", "Total_Price", "Subtotal_Amount", "Discount_Amount", "Shipping_Fee", "Shipping_Method", "Date", "Address", "Note", "Coupon_Code", "Order_Status", "Payment_Link", "Tracking_No", "Status_Note", "Internal_Note", "Last_Update", "Items_Detail"];
  const existingHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(h => h.toString().trim());
  const missingHeaders = ORDER_HEADERS_REQUIRED_.filter(h => existingHeaders.indexOf(h) === -1);
  if (missingHeaders.length > 0) {
    const startCol = existingHeaders.length + 1;
    sheet.getRange(1, startCol, 1, missingHeaders.length).setValues([missingHeaders]);
  }
  return getHeaderIndexMap_(sheet);
}

// 資安：客人填的文字若以 = + - @ 開頭，試算表會當成公式執行（可能被拿來外洩其他訂單資料），前面加 ' 一律當純文字
function sheetSafeText_(value) {
  const text = (value === undefined || value === null) ? "" : value.toString();
  return /^[=+\-@\t\r]/.test(text) ? "'" + text : text;
}

function appendOrderByHeader_(sheet, orderValues) {
  const headerMap = ensureOrderHeaders_(sheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
  const row = headers.map(header => orderValues[header] !== undefined && orderValues[header] !== null ? orderValues[header] : "");
  sheet.appendRow(row);
}

function getCellValueByHeader_(sheet, row, headerMap, headerName) {
  const col = headerMap[headerName];
  if (!col) return "";
  return sheet.getRange(row, col).getDisplayValue().toString().trim();
}

function sendStatusEmail_(sheet, row, headerMap, status) {
  const email = getCellValueByHeader_(sheet, row, headerMap, "Email");
  if (!email) return;
  const orderId = getCellValueByHeader_(sheet, row, headerMap, "ID");
  const trackingNo = getCellValueByHeader_(sheet, row, headerMap, "Tracking_No");
  const statusNote = getCellValueByHeader_(sheet, row, headerMap, "Status_Note");
  let subject = "", message = "";

  if (status === "可取貨") {
    subject = "Lilou｜您的甜點可以取貨囉！";
    message = "您預訂的甜點已製作完成，歡迎依照預約日期及時段前往門市取貨。\n\n門市地址：\n" + STORE_ADDRESS + "\n\n營業時間：\n14:00－18:00（固定週一、二公休）";
  } else if (status === "待付款") {
    subject = "Lilou｜請完成付款";
    message = "您的訂單已完成確認。\n\n請依付款通知完成信用卡付款，付款完成後我們將安排後續製作與出貨。\n\n若未收到付款通知，也可至「查詢訂單」查看最新進度。";
  } else if (status === "待出貨") {
    subject = "Lilou｜準備出貨中";
    message = "目前正安排製作與包裝，完成後將立即為您出貨。";
  } else if (status === "已出貨") {
    subject = "Lilou｜您的甜點已出貨";
    message = "您的訂單已完成出貨。\n\n可至「查詢訂單」查看配送資訊與貨運編號。";
  } else if (status === "已完成") {
    subject = "Lilou｜訂單已完成";
    message = "您的訂單已完成。\n\n非常感謝您的支持，希望 Lilou 的甜點能為您帶來美好的時光。";
  } else if (status === "已取消") {
    subject = "Lilou｜訂單已取消";
    message = "您的訂單已取消。\n\n若此次取消為系統或付款問題，或您仍有訂購需求，歡迎隨時與我們聯繫。";
  } else return;

  let body = "您好，\n\n" + message + "\n\n訂單編號：" + orderId + "\n訂單狀態：" + status;
  if (status === "已出貨" && trackingNo) body += "\n\n貨運編號：\n" + trackingNo;
  if (statusNote) body += "\n\n備註：\n" + statusNote;
  body += "\n\n您也可以隨時查詢訂單進度：\n" + ORDER_LOOKUP_URL + "\n\n" + LILOU_EMAIL_AUTO_NOTICE + "\n\nLilou｜迷你可麗露專賣";
  GmailApp.sendEmail(email, subject, body, { from: LILOU_EMAIL_FROM, name: LILOU_EMAIL_NAME, replyTo: LILOU_EMAIL_REPLY_TO });
}

function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "Orders") return;
  const headerMap = getHeaderIndexMap_(sheet);
  const statusCol = headerMap["Order_Status"], paymentCol = headerMap["Payment_Link"], trackingCol = headerMap["Tracking_No"], noteCol = headerMap["Status_Note"], updateCol = headerMap["Last_Update"];
  if (!updateCol || e.range.getRow() <= 1) return;
  const editedCol = e.range.getColumn();
  if ([statusCol, paymentCol, trackingCol, noteCol].includes(editedCol)) {
    sheet.getRange(e.range.getRow(), updateCol).setValue(Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm:ss"));
  }
  if (editedCol === statusCol) sendStatusEmail_(sheet, e.range.getRow(), headerMap, e.range.getDisplayValue().toString().trim());
}

function getAdminPassword_() { return getScriptProp_("ADMIN_PASSWORD"); }
function createAdminToken_() { return Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, ""); }
function validateAdminToken_(token) { return token && CacheService.getScriptCache().get(ADMIN_TOKEN_CACHE_PREFIX + token) === "1"; }
function adminRequire_(token) { if (!validateAdminToken_(token)) throw new Error("後台登入已失效，請重新登入。"); }

function adminLogin_(password) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get(ADMIN_LOGIN_FAIL_KEY) || 0);
  if (fails >= ADMIN_LOGIN_MAX_FAILS) return { status: "error", message: "密碼錯誤次數過多，請 10 分鐘後再試。" };
  const realPassword = getAdminPassword_();
  if (!realPassword) return { status: "error", message: "尚未設定後台密碼，請到 Apps Script 指令碼屬性設定 ADMIN_PASSWORD。" };
  if (!password || password.toString() !== realPassword) {
    cache.put(ADMIN_LOGIN_FAIL_KEY, String(fails + 1), ADMIN_LOGIN_LOCK_SECONDS);
    return { status: "error", message: "後台密碼錯誤。" };
  }
  cache.remove(ADMIN_LOGIN_FAIL_KEY);
  const token = createAdminToken_();
  cache.put(ADMIN_TOKEN_CACHE_PREFIX + token, "1", ADMIN_TOKEN_SECONDS);
  return { status: "success", token: token, expiresIn: ADMIN_TOKEN_SECONDS };
}

function adminGetSheetDisplayRows_(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet && (sheetName === "Categories" || sheetName === "Categories 分類管理")) {
    sheet = ss.getSheetByName("Categories") || ss.getSheetByName("Categories 分類管理");
  }
  if (!sheet) return { headers: [], rows: [] };
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length === 0) return { headers: [], rows: [] };
  const headers = values[0].map(h => h.toString().trim());
  const rows = values.slice(1).filter(row => row.some(v => v.toString().trim() !== "")).map((row, index) => {
    const obj = { rowNumber: index + 2 };
    headers.forEach((h, i) => obj[h] = row[i] ? row[i].toString().trim() : "");
    return obj;
  });
  return { headers: headers, rows: rows };
}
function adminGetSheetObjects_(ss, sheetName) { return adminGetSheetDisplayRows_(ss, sheetName).rows; }
function adminExtractProduction_(details) {
  const result = {};
  if (!details) return result;
  const regex = /([^【】｜、,，\n\r()（）]+)\((\d+)個[^)]*\)/g;
  let match;
  while ((match = regex.exec(details.toString())) !== null) {
    const name = match[1].trim(), qty = Number(match[2]) || 0;
    if (name && qty > 0) result[name] = (result[name] || 0) + qty;
  }
  return result;
}
function adminExtractPackaging_(details) {
  const result = {};
  if (!details) return result;
  const regex = /【第\d+組:\s*([^】×]+?)(?:\s*×\s*(\d+)組)?】/g;
  let match;
  while ((match = regex.exec(details.toString())) !== null) {
    const name = match[1].trim(), qty = Number(match[2]) || 1;
    if (name && qty > 0) result[name] = (result[name] || 0) + qty;
  }
  return result;
}
function adminGetRevenueSettings_(ss) {
  return adminGetSheetObjects_(ss, "Revenue_Settings").map(o => ({ month: o.Month || o.month || "", storeRate: o.Store_Rate || o.store_rate || "", deliveryRate: o.Delivery_Rate || o.delivery_rate || "" })).filter(o => o.month);
}

function adminGetAll_(token) {
  adminRequire_(token);
  const ss = SpreadsheetApp.openById(GLOBAL_SPREADSHEET_ID);
  
  ensureCategoriesSheet_(ss);
  ensureSettingsPromo_(ss);
  ensureSettingsDateRules_(ss);
  ensurePagesSheets_(ss);
  ensureCakeSetup_(ss);

  const orders = adminGetSheetObjects_(ss, "Orders").map(o => ({
    id: o.ID || "", timestamp: o.Timestamp || "", name: o.Name || "", phone: o.Phone || "", date: o.Date || "", shippingMethod: o.Shipping_Method || "",
    address: o.Address || "", details: o.Product_Name || "", total: o.Total_Price || "", note: o.Note || "", email: o.Email || "",
    orderStatus: o.Order_Status || "", paymentLink: o.Payment_Link || "", trackingNo: o.Tracking_No || "", statusNote: o.Status_Note || "",
    internalNote: o.Internal_Note || "", lastUpdate: o.Last_Update || "", subtotalAmount: o.Subtotal_Amount || "", discountAmount: o.Discount_Amount || "",
    shippingFee: o.Shipping_Fee || "", itemsDetail: o.Items_Detail || ""
  })).filter(o => o.id);

  const categories = adminGetSheetObjects_(ss, "Categories").map((o, index) => ({
    id: o.id || o.ID || "", name: o.name || o.Name || "", behavior: (o.behavior || o.Behavior || "combo").toString().trim().toLowerCase(),
    display: (o.display !== undefined && o.display !== "") ? o.display : "TRUE", sort: o.sort || o.Sort || ((index + 1) * 10),
    name_en: o.name_en || "", pickup_min_days: o.pickup_min_days || "", pickup_max_days: o.pickup_max_days || "",
    delivery_min_days: o.delivery_min_days || "", delivery_max_days: o.delivery_max_days || "",
    pickup_date_rule: o.pickup_date_rule || "", delivery_date_rule: o.delivery_date_rule || "", date_notice: o.date_notice || "", daily_limit: o.daily_limit || ""
  })).filter(o => o.id).sort((a,b) => (Number(a.sort) || 999999) - (Number(b.sort) || 999999));

  const products = adminGetSheetObjects_(ss, "Products").map((o, index) => ({
    name: o.name || o.Name || "", price: o.price || o.Price || "", stock: o.stock || o.Stock || "",
    type: (o.type || o.Type || "combo").toString().trim().toLowerCase(), display: (o.display !== undefined && o.display !== "") ? o.display : "TRUE",
    shipping_rule: (o.shipping_rule || o.Shipping_Rule || o.shipping || "all").toString().trim().toLowerCase(), sort: o.sort || o.Sort || ((index + 1) * 10),
    name_en: o.name_en || "", image: o.image || "", min_per_order: o.min_per_order || "", max_per_order: o.max_per_order || "",
    description: o.description || o.Description || o.product_description || o["商品描述"] || "", ingredients: o.ingredients || o.Ingredients || "", allergens: o.allergens || o.Allergens || "", storage: o.storage || o.Storage || "", weight: o.weight || o.Weight || "", nutrition: o.nutrition || o.Nutrition || "", origin_manufacturer: o.origin_manufacturer || o.Origin_Manufacturer || o["原產地與製造商"] || ""
  })).filter(o => o.name).sort((a, b) => (Number(a.sort) || 999999) - (Number(b.sort) || 999999) || a.name.localeCompare(b.name, 'zh-Hant'));

  const packaging = adminGetSheetObjects_(ss, "Packaging").map((o, index) => ({
    name: o.name || o.Name || "", extra_price: o.extra_price || o.Extra_Price || "", box_size: o.box_size || o.Box_Size || "",
    type: (o.type || o.Type || o.packaging_type || "combo").toString().trim().toLowerCase(), display: (o.display !== undefined && o.display !== "") ? o.display : "TRUE", sort: o.sort || o.Sort || ((index + 1) * 10),
    name_en: o.name_en || ""
  })).filter(o => o.name).sort((a, b) => (Number(a.sort) || 999999) - (Number(b.sort) || 999999) || a.name.localeCompare(b.name, 'zh-Hant'));

  const coupons = adminGetSheetObjects_(ss, "Coupons").map(o => ({
    code: o.code || "", type: o.type || "", value: o.value || "", min_spend: o.min_spend || "", start_date: o.start_date || "", end_date: o.end_date || ""
  })).filter(o => o.code);

  const settings = adminGetSheetObjects_(ss, "Settings").map(o => ({ setting_name: o.setting_name || o.Setting_Name || "", value: o.value || o.Value || "" })).filter(o => o.setting_name);
  const revenueSettings = adminGetRevenueSettings_(ss);

  const dashboard = {
    totalOrders: orders.length, pendingPayment: orders.filter(o => o.orderStatus === "待付款").length, paid: orders.filter(o => o.orderStatus === "待出貨").length,
    readyPickup: orders.filter(o => o.orderStatus === "可取貨").length, shipped: orders.filter(o => o.orderStatus === "已出貨").length, pickedUp: orders.filter(o => o.orderStatus === "已完成").length,
    inProgress: orders.filter(o => o.orderStatus !== "已完成").length, completed: orders.filter(o => o.orderStatus === "已完成").length, revenue: orders.reduce((sum, o) => sum + (Number(String(o.total).replace(/[^0-9.-]/g, "")) || 0), 0)
  };

  const productionSummary = {}, packagingSummary = {};
  orders.forEach(o => {
    Object.entries(adminExtractProduction_(o.details)).forEach(([name, qty]) => productionSummary[name] = (productionSummary[name] || 0) + qty);
    Object.entries(adminExtractPackaging_(o.details)).forEach(([name, qty]) => packagingSummary[name] = (packagingSummary[name] || 0) + qty);
  });

  const pages = readSheetObjects_(ss, "Pages").filter(p => p.page_id).sort((a, b) => (Number(a.sort) || 999) - (Number(b.sort) || 999));
  const pageBlocks = readSheetObjects_(ss, "PageBlocks").filter(b => b.page_id).sort((a, b) => (Number(a.sort) || 999) - (Number(b.sort) || 999));
  const cakeOptions = readSheetObjects_(ss, "Cake_Options").filter(o => o.name);
  return { status: "success", dashboard, orders, categories, products, packaging, coupons, settings, revenueSettings, productionSummary, packagingSummary, pages, pageBlocks, cakeOptions };
}

// 🔽 已修正：支援「Categories 分類管理」別名與自動分頁對齊
function adminReplaceSheetRows_(token, sheetName, rows) {
  adminRequire_(token);
  const categoryConfig = { coreHeaders: ["id", "name", "behavior", "display", "sort"].concat(CATEGORY_RULE_HEADERS), keyHeader: "id" };
  const config = {
    Categories: categoryConfig,
    "Categories 分類管理": categoryConfig,
    Products: { coreHeaders: ["name", "price", "stock", "type", "display", "shipping_rule", "sort", "description", "ingredients", "allergens", "storage", "weight", "nutrition", "origin_manufacturer", "name_en", "image", "min_per_order", "max_per_order"], keyHeader: "name" },
    Packaging: { coreHeaders: ["name", "extra_price", "box_size", "type", "display", "sort", "name_en"], keyHeader: "name" },
    Coupons: { coreHeaders: ["code", "type", "value", "min_spend", "start_date", "end_date"], keyHeader: "code" },
    Settings: { coreHeaders: ["setting_name", "value"], keyHeader: "setting_name" },
    Pages: { coreHeaders: PAGE_HEADERS, keyHeader: "page_id" },
    PageBlocks: { coreHeaders: PAGE_BLOCK_HEADERS, keyHeader: "block_id" },
    Cake_Options: { coreHeaders: CAKE_OPTION_HEADERS, keyHeader: "name" }
  };
  if (!config[sheetName]) throw new Error("不允許更新此分頁：" + sheetName);

  const ss = SpreadsheetApp.openById(GLOBAL_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  
  // 兼顧分頁命名為 Categories 或 Categories 分類管理
  if (!sheet && (sheetName === "Categories" || sheetName === "Categories 分類管理")) {
    sheet = ss.getSheetByName("Categories") || ss.getSheetByName("Categories 分類管理");
  }
  if (!sheet) throw new Error("找不到 " + sheetName + " 分頁");

  const existing = adminGetSheetDisplayRows_(ss, sheet.getName());
  const existingHeaders = existing.headers.filter(Boolean);
  const coreHeaders = config[sheetName].coreHeaders;
  const keyHeader = config[sheetName].keyHeader;
  const headers = existingHeaders.length ? existingHeaders.concat(coreHeaders.filter(h => existingHeaders.indexOf(h) === -1)) : coreHeaders;

  const oldRowMap = {};
  existing.rows.forEach(row => {
    const key = (row[keyHeader] || "").toString().trim();
    if (key) oldRowMap[key] = row;
  });

  const cleanRows = (rows || []).filter(row => coreHeaders.some(h => row[h] !== undefined && row[h] !== null && row[h].toString().trim() !== ""));
  
  sheet.getDataRange().clearDataValidations();
  sheet.clearContents();
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (cleanRows.length > 0) {
    const values = cleanRows.map(row => {
      const key = (row[keyHeader] || "").toString().trim();
      const oldRow = oldRowMap[key] || {};
      return headers.map(h => {
        if (row[h] !== undefined && row[h] !== null) return row[h].toString().trim();
        if (oldRow[h] !== undefined && oldRow[h] !== null) return oldRow[h].toString().trim();
        return "";
      });
    });
    sheet.getRange(2, 1, values.length, headers.length).setValues(values);
  }
  try { CacheService.getScriptCache().remove(FORM_OPTIONS_CACHE_KEY); } catch (cacheError) {}
  return { status: "success", message: sheetName + " 已更新。", count: cleanRows.length };
}

// 商品圖片上傳：存到雲端硬碟「Lilou 商品圖片」資料夾，設為知道連結者可檢視，回傳圖片網址
// ===== 商品單項儲存（只更新那一列，多人同時編輯不會互相覆蓋）=====
const PRODUCT_FIELDS = ["name", "name_en", "image", "price", "stock", "min_per_order", "max_per_order", "type", "display", "shipping_rule", "sort", "description", "ingredients", "allergens", "storage", "weight", "nutrition", "origin_manufacturer"];

function productSheetInfo_() {
  const ss = SpreadsheetApp.openById(GLOBAL_SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Products");
  if (!sheet) throw new Error("找不到 Products 分頁");
  let headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(h => h.toString().trim());
  const missing = PRODUCT_FIELDS.filter(f => headers.indexOf(f) === -1);
  if (missing.length) {
    const start = headers.filter(Boolean).length + 1;
    sheet.getRange(1, start, 1, missing.length).setValues([missing]);
    headers = headers.filter(Boolean).concat(missing);
  }
  const nameCol = headers.indexOf("name");
  const lastRow = sheet.getLastRow();
  const names = lastRow > 1 ? sheet.getRange(2, nameCol + 1, lastRow - 1, 1).getDisplayValues().map(r => r[0].toString().trim()) : [];
  return { sheet, headers, names };
}
function clearFormCache_() { try { CacheService.getScriptCache().remove(FORM_OPTIONS_CACHE_KEY); } catch (e) {} }

// data.originalName：原本的名稱（新增商品時留空）；data.product：要寫入的欄位（只寫有給的欄位）
function adminSaveProduct_(data) {
  adminRequire_(data.token);
  const p = data.product || {};
  const name = (p.name !== undefined ? p.name : data.originalName || "").toString().trim();
  if (!name) throw new Error("商品名稱不能空白。");
  const info = productSheetInfo_();
  const original = (data.originalName || "").toString().trim();
  let idx = original ? info.names.indexOf(original) : (data.fillBlank ? info.names.indexOf("") : -1);
  if (original && idx === -1) throw new Error("找不到商品「" + original + "」，可能已被其他人改名或刪除，請重新整理。");
  const dup = info.names.indexOf(name);
  if (dup !== -1 && dup !== idx) throw new Error("已經有同名的商品「" + name + "」。");
  if (idx === -1) {
    const row = info.headers.map(h => p[h] !== undefined && p[h] !== null ? p[h].toString() : (h === "display" ? "TRUE" : ""));
    row[info.headers.indexOf("name")] = name;
    info.sheet.appendRow(row);
  } else {
    const rowNo = idx + 2;
    const range = info.sheet.getRange(rowNo, 1, 1, info.headers.length);
    const values = range.getValues()[0];
    info.headers.forEach((h, i) => { if (PRODUCT_FIELDS.indexOf(h) !== -1 && p[h] !== undefined && p[h] !== null) values[i] = p[h].toString(); });
    range.setValues([values]);
  }
  clearFormCache_();
  return { status: "success", message: "已儲存「" + name + "」。", name: name };
}

// data.order：[{ name, sort }]，只更新排序欄
function adminSaveProductOrder_(data) {
  adminRequire_(data.token);
  const info = productSheetInfo_();
  const sortCol = info.headers.indexOf("sort") + 1;
  let count = 0;
  (data.order || []).forEach(o => {
    const idx = info.names.indexOf((o.name || "").toString().trim());
    if (idx === -1) return;
    info.sheet.getRange(idx + 2, sortCol).setValue(Number(o.sort) || 0);
    count++;
  });
  clearFormCache_();
  return { status: "success", message: "排序已儲存。", count: count };
}

function adminDeleteProduct_(data) {
  adminRequire_(data.token);
  const info = productSheetInfo_();
  const idx = info.names.indexOf((data.name || "").toString().trim());
  if (idx === -1) throw new Error("找不到這項商品，可能已被刪除，請重新整理。");
  info.sheet.deleteRow(idx + 2);
  clearFormCache_();
  return { status: "success", message: "已刪除「" + data.name + "」。" };
}

// ===== 分類／包裝／優惠券／蛋糕選項：單項儲存、刪除、排序（只動那一列）=====
function rowSheetConfig_(kind) {
  const cfgs = {
    categories: { sheets: ["Categories", "Categories 分類管理"], fields: ["id", "name", "behavior", "display", "sort"].concat(CATEGORY_RULE_HEADERS), key: ["id"], label: "分類" },
    packaging: { sheets: ["Packaging"], fields: ["name", "extra_price", "box_size", "type", "display", "sort", "name_en"], key: ["name"], label: "包裝" },
    coupons: { sheets: ["Coupons"], fields: ["code", "type", "value", "min_spend", "start_date", "end_date"], key: ["code"], label: "優惠券" },
    cakeoptions: { sheets: ["Cake_Options"], fields: CAKE_OPTION_HEADERS, key: ["group", "name"], label: "蛋糕選項" }
  };
  const cfg = cfgs[kind];
  if (!cfg) throw new Error("不支援的資料類型：" + kind);
  return cfg;
}
function rowSheetInfo_(kind) {
  const cfg = rowSheetConfig_(kind);
  const ss = SpreadsheetApp.openById(GLOBAL_SPREADSHEET_ID);
  let sheet = null;
  cfg.sheets.forEach(n => { if (!sheet) sheet = ss.getSheetByName(n); });
  if (!sheet) { sheet = ss.insertSheet(cfg.sheets[0]); sheet.getRange(1, 1, 1, cfg.fields.length).setValues([cfg.fields]); }
  let headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(h => h.toString().trim());
  const missing = cfg.fields.filter(f => headers.indexOf(f) === -1);
  if (missing.length) {
    const start = headers.filter(Boolean).length + 1;
    sheet.getRange(1, start, 1, missing.length).setValues([missing]);
    headers = headers.filter(Boolean).concat(missing);
  }
  const lastRow = sheet.getLastRow();
  const values = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, headers.length).getDisplayValues() : [];
  const keys = values.map(r => rowKey_(cfg, h => r[headers.indexOf(h)]));
  return { cfg, sheet, headers, keys };
}
function rowKey_(cfg, get) {
  return cfg.key.map(k => { const v = (get(k) || "").toString().trim(); return k === "group" ? v.toLowerCase() : v; }).join("|");
}

function adminSaveRow_(data) {
  adminRequire_(data.token);
  const info = rowSheetInfo_(data.kind);
  const cfg = info.cfg;
  const row = data.row || {};
  const original = (data.originalKey || "").toString();
  const idx = original ? info.keys.indexOf(original) : -1;
  if (original && idx === -1) throw new Error("找不到這筆" + cfg.label + "，可能已被其他人修改或刪除，請重新整理。");
  const merged = {};
  if (idx !== -1) {
    const cur = info.sheet.getRange(idx + 2, 1, 1, info.headers.length).getDisplayValues()[0];
    info.headers.forEach((h, i) => { merged[h] = cur[i]; });
  }
  cfg.fields.forEach(f => { if (row[f] !== undefined && row[f] !== null) merged[f] = row[f].toString().trim(); });
  const newKey = rowKey_(cfg, k => merged[k]);
  if (cfg.key.some(k => !(merged[k] || "").toString().trim())) throw new Error(cfg.label + "名稱不能空白。");
  const dup = info.keys.indexOf(newKey);
  if (dup !== -1 && dup !== idx) throw new Error("已經有同名的" + cfg.label + "。");
  if (idx === -1) {
    if (cfg.fields.indexOf("display") !== -1 && !merged.display) merged.display = "TRUE";
    info.sheet.appendRow(info.headers.map(h => merged[h] !== undefined ? merged[h] : ""));
  } else {
    const range = info.sheet.getRange(idx + 2, 1, 1, info.headers.length);
    const values = range.getValues()[0];
    info.headers.forEach((h, i) => { if (cfg.fields.indexOf(h) !== -1 && row[h] !== undefined && row[h] !== null) values[i] = row[h].toString().trim(); });
    range.setValues([values]);
  }
  clearFormCache_();
  return { status: "success", message: cfg.label + "已儲存。", key: newKey };
}

function adminDeleteRow_(data) {
  adminRequire_(data.token);
  const info = rowSheetInfo_(data.kind);
  const idx = info.keys.indexOf((data.key || "").toString());
  if (idx === -1) throw new Error("找不到這筆" + info.cfg.label + "，可能已被刪除，請重新整理。");
  info.sheet.deleteRow(idx + 2);
  clearFormCache_();
  return { status: "success", message: info.cfg.label + "已刪除。" };
}

function adminSaveRowOrder_(data) {
  adminRequire_(data.token);
  const info = rowSheetInfo_(data.kind);
  const sortCol = info.headers.indexOf("sort") + 1;
  if (!sortCol) throw new Error("這個分頁沒有排序欄。");
  let count = 0;
  (data.order || []).forEach(o => {
    const idx = info.keys.indexOf((o.key || "").toString());
    if (idx === -1) return;
    info.sheet.getRange(idx + 2, sortCol).setValue(Number(o.sort) || 0);
    count++;
  });
  clearFormCache_();
  return { status: "success", message: "排序已儲存。", count: count };
}

// ===== 前台資料快照：後台改了商品／分類／設定等內容後，把前台要用的資料存成 GitHub 上的 data/site.json =====
// 第一次造訪的客人先讀這份靜態檔（很快），不必等 Apps Script 回應；之後前台仍會向 Apps Script 取得最新資料
function withSnapshot_(res) {
  if (res && res.status === "success") res.snapshot = publishSiteSnapshot_();
  return res;
}
function publishSiteSnapshot_() {
  try {
    if (!getScriptProp_("GITHUB_TOKEN")) return false;
    clearFormCache_();
    const data = getFormOptions();
    data.snapshotAt = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
    const cfg = githubConfig_();
    const path = "data/site.json";
    const url = "https://api.github.com/repos/" + cfg.repo + "/contents/" + path;
    let sha = null;
    const got = UrlFetchApp.fetch(url + "?ref=" + encodeURIComponent(cfg.branch), { headers: cfg.headers, muteHttpExceptions: true });
    if (got.getResponseCode() === 200) sha = JSON.parse(got.getContentText()).sha;
    const body = { message: "更新前台資料快照", content: Utilities.base64Encode(JSON.stringify(data), Utilities.Charset.UTF_8), branch: cfg.branch };
    if (sha) body.sha = sha;
    const res = UrlFetchApp.fetch(url, { method: "put", contentType: "application/json", headers: cfg.headers, payload: JSON.stringify(body), muteHttpExceptions: true });
    return res.getResponseCode() === 200 || res.getResponseCode() === 201;
  } catch (e) {
    return false;
  }
}

// ===== 後台上傳照片到 GitHub（存進 repo 的 images/uploads 資料夾）=====
// 指令碼屬性：GITHUB_TOKEN（必填，只能存取這個 repo 的 Contents 讀寫金鑰）、GITHUB_REPO（預設 liloucanele/online-shop）、GITHUB_BRANCH（選填，預設自動偵測）
function githubConfig_() {
  const token = getScriptProp_("GITHUB_TOKEN");
  if (!token) throw new Error("尚未設定 GITHUB_TOKEN，請到 Apps Script 指令碼屬性新增。");
  const repo = getScriptProp_("GITHUB_REPO") || "liloucanele/online-shop";
  const headers = { Authorization: "Bearer " + token, Accept: "application/vnd.github+json", "X-GitHub-Api-Version": "2022-11-28" };
  let branch = getScriptProp_("GITHUB_BRANCH");
  if (!branch) {
    const res = UrlFetchApp.fetch("https://api.github.com/repos/" + repo, { headers: headers, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error(githubErrorText_(res.getResponseCode()));
    branch = JSON.parse(res.getContentText()).default_branch || "main";
    PropertiesService.getScriptProperties().setProperty("GITHUB_BRANCH", branch);
  }
  return { repo: repo, branch: branch, headers: headers };
}
function githubErrorText_(code) {
  if (code === 401) return "GitHub 金鑰無效或已過期，請重新產生後更新 GITHUB_TOKEN。";
  if (code === 403) return "GitHub 金鑰沒有寫入權限，請確認 Contents 設為 Read and write。";
  if (code === 404) return "找不到 GitHub repo，或金鑰沒有這個 repo 的權限，請確認 GITHUB_REPO 與金鑰設定。";
  return "GitHub 回應錯誤（" + code + "），請稍後再試。";
}
function adminUploadImageToGithub_(data) {
  adminRequire_(data.token);
  const base64 = (data.base64 || "").toString().replace(/^data:[^,]+,/, "");
  const thumb64 = (data.thumbBase64 || "").toString().replace(/^data:[^,]+,/, "");
  if (!base64) throw new Error("沒有收到圖片。");
  if (base64.length > 4 * 1024 * 1024 || thumb64.length > 1024 * 1024) throw new Error("圖片太大，請換一張較小的照片。");
  const cfg = githubConfig_();
  const stamp = Utilities.formatDate(new Date(), TIMEZONE, "yyyyMMdd-HHmmss");
  const rand = Utilities.getUuid().slice(0, 6);
  const path = "images/uploads/" + stamp + "-" + rand + ".jpg";
  // 小圖（縮圖用）先傳，檔名在大圖後面加 -s；小圖失敗不影響大圖，前台會自動改用大圖
  if (thumb64) {
    try { githubPutFile_(cfg, path.replace(/\.jpg$/, "-s.jpg"), thumb64); } catch (e) {}
  }
  githubPutFile_(cfg, path, base64);
  return { status: "success", url: path, preview: "https://raw.githubusercontent.com/" + cfg.repo + "/" + cfg.branch + "/" + path };
}
function githubPutFile_(cfg, path, base64) {
  const res = UrlFetchApp.fetch("https://api.github.com/repos/" + cfg.repo + "/contents/" + path, {
    method: "put", contentType: "application/json", headers: cfg.headers, muteHttpExceptions: true,
    payload: JSON.stringify({ message: "後台上傳照片 " + path, content: base64, branch: cfg.branch })
  });
  const code = res.getResponseCode();
  if (code !== 201 && code !== 200) throw new Error(githubErrorText_(code));
}

function adminUpdateOrder_(data) {
  adminRequire_(data.token);
  const orderId = (data.orderId || "").toString().trim();
  if (!orderId) throw new Error("缺少訂單編號。");
  const newStatus = (data.orderStatus || "").toString().trim();

  const ss = SpreadsheetApp.openById(GLOBAL_SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Orders");
  const headerMap = getHeaderIndexMap_(sheet);
  const idCol = headerMap["ID"];
  const ids = sheet.getRange(2, idCol, Math.max(sheet.getLastRow() - 1, 1), 1).getDisplayValues().flat();

  let targetRow = -1;
  for (let i = 0; i < ids.length; i++) { if (ids[i].toString().trim() === orderId) { targetRow = i + 2; break; } }
  if (targetRow === -1) throw new Error("找不到此訂單：" + orderId);

  const oldStatus = getCellValueByHeader_(sheet, targetRow, headerMap, "Order_Status");
  if (newStatus === "已取消" && (oldStatus === "已出貨" || oldStatus === "已完成")) throw new Error("已出貨或已完成的訂單不可改為已取消。");

  if (newStatus && headerMap["Order_Status"]) sheet.getRange(targetRow, headerMap["Order_Status"]).setValue(newStatus);
  if (headerMap["Payment_Link"]) sheet.getRange(targetRow, headerMap["Payment_Link"]).setValue((data.paymentLink || "").toString().trim());
  if (headerMap["Tracking_No"]) sheet.getRange(targetRow, headerMap["Tracking_No"]).setValue((data.trackingNo || "").toString().trim());
  if (headerMap["Status_Note"]) sheet.getRange(targetRow, headerMap["Status_Note"]).setValue((data.statusNote || "").toString().trim());
  if (headerMap["Internal_Note"]) sheet.getRange(targetRow, headerMap["Internal_Note"]).setValue((data.internalNote || "").toString().trim());
  // 後台改取貨／到貨日期（客人改期用）：只有後台有送 date 才改，不檢查公休日、營業日與每日限量
  const newDate = (data.date || "").toString().trim();
  if (newDate && headerMap["Date"]) {
    if (!/^\d{4}\/\d{2}\/\d{2}/.test(newDate)) throw new Error("日期格式不正確。");
    sheet.getRange(targetRow, headerMap["Date"]).setNumberFormat("@").setValue(newDate);
  }

  const nowText = Utilities.formatDate(new Date(), TIMEZONE, "yyyy/MM/dd HH:mm:ss");
  if (headerMap["Last_Update"]) sheet.getRange(targetRow, headerMap["Last_Update"]).setValue(nowText);
  if (newStatus && newStatus !== oldStatus) try { sendStatusEmail_(sheet, targetRow, headerMap, newStatus); } catch (e) {}
  return { status: "success", message: "訂單已更新。", orderId: orderId, lastUpdate: nowText, date: newDate || undefined };
}
