// Приёмник анкеты → Google Таблица «Анкета — бронь места (карусели через AI)»
const SHEET_ID = '1muaJVeg3mq3BA0Uvua5TFq-9aktyUGC1PegEZPfa_u0';
const SHEET_NAME = 'Анкеты';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const p = (e && e.parameter) || {};
    const clean = v => {
      v = String(v || '').slice(0, 2000);
      return /^[=+\-@]/.test(v) ? "'" + v : v; // телефон «+7…» и защита от формул
    };
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    sheet.appendRow([
      new Date(),
      p.username ? '@' + clean(p.username).replace(/^'?@?/, '') : '',
      clean(p.name),
      clean(p.role),
      clean(p.followers),
      clean(p.goals),
      clean(p.pains),
      clean(p.formats),
      clean(p.carousel_time),
      clean(p.need),
      clean(p.device),
      clean(p.email),
      clean(p.phone),
      clean(p.clientId)
    ]);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('Анкета: приёмник работает');
}
