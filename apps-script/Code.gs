// 掐指一算 的資料接收程式
// 貼在「掐指一算｜數據後台」試算表的 擴充功能 > Apps Script
// 網頁送來的每個事件，會在 events 分頁新增一列

const SHEET_NAME = 'events';
const HEADERS = ['時間', '訪客', 'session', '事件', '數值', '方向', '模式', '裝置', '來源', '有寫問題'];
const EVENTS = ['visit', 'start', 'finish', 'skip', 'card', 'share', 'save_card',
                'copy_result', 'copy_link', 'again', 'repeat'];

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (EVENTS.indexOf(d.event) === -1) return out('ignored');

    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      sheet().appendRow([
        new Date(),
        clip(d.visitor), clip(d.session), d.event, clip(d.value),
        clip(d.topic), clip(d.mode), clip(d.device), clip(d.source), clip(d.asked)
      ]);
    } finally {
      lock.releaseLock();
    }
    return out('ok');
  } catch (err) {
    return out('error');
  }
}

function doGet() {
  return out('alive');
}

function sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

// 每格最多 60 字，避免有人亂送很長的資料
function clip(v) {
  return v === undefined || v === null ? '' : String(v).slice(0, 60);
}

function out(text) {
  return ContentService.createTextOutput(text);
}
