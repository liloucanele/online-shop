// ===== 訂單 90 天後移到封存分頁 =====
// 每天凌晨 3 點自動執行：「已完成」或「已取消」滿 90 天的訂單，從 Orders 搬到「Orders_封存」分頁。
// 搬過去後客人用手機號碼查不到、後台訂單列表也不再顯示；資料原封不動保留在封存分頁，沒有刪除或遮蔽。
// 90 天的起算日：訂單最後一次更新狀態的時間（Last_Update，改成已完成／已取消時會記錄）；
//   沒有這個時間的舊訂單，改用取貨／到貨日（Date），再沒有就用下單時間（Timestamp）。
// 第一次使用：在 Apps Script 編輯器執行一次 installArchiveTrigger()。
// 想先看會搬哪些、不真的搬：執行 previewArchiveOrders()，結果在「執行記錄」。

const ARCHIVE_RETENTION_DAYS_ = 90;
const ARCHIVE_SHEET_NAME_ = "Orders_封存";
const ARCHIVE_STATUSES_ = ["已完成", "已取消"];

function archiveDateTime_(text) {
  const d = normalizeDateString((text || "").toString().trim().split(/[ T(（]/)[0]);
  return d ? new Date(d + " 00:00:00").getTime() : NaN;
}

function archiveOldOrders_(dryRun) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { status: "busy" };
  try {
    const ss = SpreadsheetApp.openById(GLOBAL_SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Orders");
    if (!sheet || sheet.getLastRow() < 2) return { status: "success", moved: 0 };
    const range = sheet.getDataRange();
    const values = range.getValues();          // 原始值：搬過去格式不變
    const shown = range.getDisplayValues();    // 顯示文字：判斷日期與狀態
    const headers = shown[0].map(h => h.toString().trim());
    const col = n => headers.indexOf(n);
    const iId = col("ID"), iStatus = col("Order_Status"), iUpd = col("Last_Update"), iDate = col("Date"), iTs = col("Timestamp");
    if (iStatus === -1) return { status: "error", message: "Orders 找不到 Order_Status 欄位" };
    const cutoff = getTaiwanTodayTime() - ARCHIVE_RETENTION_DAYS_ * 86400000;

    const moveRows = [], moveRowNums = [], ids = [];
    for (let r = 1; r < shown.length; r++) {
      const row = shown[r];
      if (ARCHIVE_STATUSES_.indexOf(row[iStatus].toString().trim()) === -1) continue;
      let t = iUpd > -1 ? archiveDateTime_(row[iUpd]) : NaN;
      if (isNaN(t) && iDate > -1) t = archiveDateTime_(row[iDate]);
      if (isNaN(t) && iTs > -1) t = archiveDateTime_(row[iTs]);
      if (isNaN(t) || t >= cutoff) continue;
      moveRows.push(values[r]);
      moveRowNums.push(r + 1);
      ids.push(iId > -1 ? row[iId] : "第" + (r + 1) + "列");
    }

    if (dryRun || !moveRows.length) {
      Logger.log((dryRun ? "預覽：將封存 " : "沒有需要封存的訂單 ") + moveRows.length + " 筆 " + ids.join(", "));
      return { status: "success", moved: 0, wouldMove: moveRows.length, ids: ids };
    }

    let archive = ss.getSheetByName(ARCHIVE_SHEET_NAME_);
    if (!archive) {
      archive = ss.insertSheet(ARCHIVE_SHEET_NAME_);
      archive.getRange(1, 1, 1, headers.length + 1).setValues([headers.concat(["Archived_At"])]);
      archive.setFrozenRows(1);
    }
    // 封存分頁的欄位順序以它自己的表頭為準，Orders 之後新增欄位也會補上
    const archHeaders = archive.getRange(1, 1, 1, archive.getLastColumn()).getValues()[0].map(h => h.toString().trim());
    headers.concat(["Archived_At"]).forEach(h => { if (h && archHeaders.indexOf(h) === -1) { archHeaders.push(h); archive.getRange(1, archHeaders.length).setValue(h); } });
    const archivedAt = Utilities.formatDate(new Date(), TIMEZONE, "yyyy/MM/dd HH:mm:ss");
    const out = moveRows.map(r => archHeaders.map(h => h === "Archived_At" ? archivedAt : (col(h) > -1 ? r[col(h)] : "")));

    // 先寫進封存分頁、確認寫入後才從 Orders 刪除，中途出錯也不會遺失訂單
    archive.getRange(archive.getLastRow() + 1, 1, out.length, archHeaders.length).setValues(out);
    SpreadsheetApp.flush();
    for (let k = moveRowNums.length - 1; k >= 0;) {   // 由下往上刪，連續的列一次刪
      const end = moveRowNums[k]; let start = end;
      while (k - 1 >= 0 && moveRowNums[k - 1] === start - 1) { start--; k--; }
      sheet.deleteRows(start, end - start + 1);
      k--;
    }
    try { CacheService.getScriptCache().remove(FORM_OPTIONS_CACHE_KEY); } catch (e) {}
    Logger.log("封存完成：搬移 " + moveRows.length + " 筆 " + ids.join(", "));
    return { status: "success", moved: moveRows.length, ids: ids };
  } finally {
    lock.releaseLock();
  }
}

function previewArchiveOrders() { return archiveOldOrders_(true); }
function runArchiveOrdersDaily() { return archiveOldOrders_(false); }

function installArchiveTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "runArchiveOrdersDaily").forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("runArchiveOrdersDaily").timeBased().everyDays(1).atHour(3).inTimezone("Asia/Taipei").create();
  Logger.log("已設定：每天凌晨 3 點自動封存");
}
