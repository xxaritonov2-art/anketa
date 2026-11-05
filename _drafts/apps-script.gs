// Приёмник анкеты → Google Таблица «Анкета — бронь места (карусели через AI)»
const SHEET_ID = '1muaJVeg3mq3BA0Uvua5TFq-9aktyUGC1PegEZPfa_u0';
const SHEET_NAME = 'Анкеты';
const COL_INTERVIEW = 20; // T: Интервью (Да / Нет / не ответил)
const COL_SID = 21;       // U: ID анкеты

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // анкета шлёт ответы простым текстом вида key=value&key2=value2 — разбираем их сами
    const p = Object.assign({}, (e && e.parameter) || {});
    const raw = e && e.postData && e.postData.contents;
    if (raw && raw.indexOf('=') > 0) {
      raw.split('&').forEach(pair => {
        const k = pair.indexOf('=');
        if (k < 1) return;
        const key = decodeURIComponent(pair.slice(0, k).replace(/\+/g, ' '));
        const val = decodeURIComponent(pair.slice(k + 1).replace(/\+/g, ' '));
        if (!(key in p) || p[key] === '') p[key] = val;
      });
    }
    const clean = v => {
      v = String(v || '').slice(0, 2000);
      return /^[=+\-@]/.test(v) ? "'" + v : v; // телефон «+7…» и защита от формул
    };
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);

    // Ответ про интервью: дописываем в строку этой же анкеты
    if (p.update === 'interview') {
      const sid = String(p.sid || '');
      const last = sheet.getLastRow();
      if (sid && last > 1) {
        const ids = sheet.getRange(2, COL_SID, last - 1, 1).getValues();
        for (let r = ids.length - 1; r >= 0; r--) {
          if (String(ids[r][0]) === sid) {
            sheet.getRange(r + 2, COL_INTERVIEW).setValue(clean(p.interview));
            return ContentService.createTextOutput('ok');
          }
        }
      }
      // строка не нашлась (основная отправка потерялась) — пишем короткую строку, чтобы ответ не пропал
      const row = new Array(COL_SID).fill('');
      row[0] = new Date();
      row[1] = p.username ? '@' + clean(p.username).replace(/^'?@?/, '') : '';
      row[13] = clean(p.clientId);
      row[COL_INTERVIEW - 1] = clean(p.interview);
      row[COL_SID - 1] = clean(sid);
      sheet.appendRow(row);
      return ContentService.createTextOutput('ok');
    }

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
      clean(p.clientId),
      clean(p.ai_level),
      clean(p.ai_tools),
      clean(p.life),
      clean(p.price),
      clean(p.pay_format),
      clean(p.interview),
      clean(p.sid)
    ]);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('Анкета: приёмник работает');
}
