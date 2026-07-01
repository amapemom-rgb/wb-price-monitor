/**
 * WB Price Monitor — мониторинг цен конкурентов на Wildberries.
 *
 * Что делает:
 *  1. Читает артикулы из листа «Товары».
 *  2. Одним батч-запросом берёт текущие цены у публичного API WB.
 *  3. Сравнивает с предыдущей ценой; при изменении пишет в лист «Лог»
 *     и шлёт уведомление в Telegram.
 *
 * БЕЗОПАСНОСТЬ:
 *  - Токен бота и chat_id хранятся в Script Properties, НЕ в коде.
 *  - Артикулы конкурентов хранятся в таблице, НЕ в коде
 *    (репозиторий публичный).
 */

// ===== Настройки =====
var SHEET_PRODUCTS = 'Товары';
var SHEET_LOG      = 'Лог';
var WB_DEST        = '-1257786'; // геозона (Москва по умолчанию)
var REQUEST_DELAY_MS = 400;      // пауза между батч-запросами
var CHUNK_SIZE       = 50;       // сколько артикулов в одном запросе

// Столбцы листа «Товары» (в массиве — 0-based):
// A=Артикул(0) B=Метка(1) C=Название(2) D=Пред.цена(3)
// E=Тек.цена(4) F=Изменение(5) G=Обновлено(6)

/**
 * Главная функция — её вызывает триггер по расписанию.
 */
function checkPrices() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_PRODUCTS);
  if (!sheet) {
    throw new Error('Не найден лист «' + SHEET_PRODUCTS + '»');
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // только заголовок, данных нет

  var range = sheet.getRange(2, 1, lastRow - 1, 7);
  var rows = range.getValues();

  // Собираем список артикулов
  var nmList = [];
  for (var i = 0; i < rows.length; i++) {
    var nm = String(rows[i][0]).trim();
    if (nm) nmList.push(nm);
  }
  if (nmList.length === 0) return;

  // Запрашиваем цены батчами
  var prices = {}; // nm -> {price, name}
  for (var c = 0; c < nmList.length; c += CHUNK_SIZE) {
    var chunk = nmList.slice(c, c + CHUNK_SIZE);
    var data = fetchWbPrices(chunk);
    for (var key in data) prices[key] = data[key];
    Utilities.sleep(REQUEST_DELAY_MS);
  }

  var now = new Date();
  var changes = [];

  for (var j = 0; j < rows.length; j++) {
    var artikul = String(rows[j][0]).trim();
    if (!artikul) continue;

    var info = prices[artikul];
    if (!info) {
      // товар не найден — помечаем, но не роняем весь прогон
      rows[j][6] = 'нет данных ' + formatDate(now);
      continue;
    }

    var oldPrice = Number(rows[j][4]) || 0; // тек.цена прошлого прогона
    var newPrice = info.price;

    rows[j][2] = info.name;       // Название
    rows[j][3] = oldPrice;        // Предыдущая цена
    rows[j][4] = newPrice;        // Текущая цена
    rows[j][6] = formatDate(now); // Обновлено

    if (oldPrice > 0 && newPrice !== oldPrice) {
      var delta = Math.round((newPrice - oldPrice) * 100) / 100;
      var pct = Math.round((delta / oldPrice) * 1000) / 10;
      rows[j][5] = (delta > 0 ? '+' : '') + delta + ' ₽ (' + (pct > 0 ? '+' : '') + pct + '%)';
      changes.push({
        nm: artikul, name: info.name,
        oldPrice: oldPrice, newPrice: newPrice, delta: delta, pct: pct
      });
    } else if (oldPrice === 0) {
      rows[j][5] = 'базовая'; // первый прогон — фиксируем базу, алерт не шлём
    }
  }

  range.setValues(rows);

  if (changes.length > 0) {
    writeLog(ss, changes, now);
    sendTelegram(buildMessage(changes));
  }
}

/**
 * Запрос цен у публичного API WB для списка артикулов (батч).
 * Возвращает объект nm -> {price, name}.
 */
function fetchWbPrices(nmArray) {
  var url = 'https://card.wb.ru/cards/v2/detail?appType=1&curr=rub&dest=' + WB_DEST +
            '&nm=' + nmArray.join(';');
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var code = resp.getResponseCode();
  if (code !== 200) {
    // роняем прогон → Google пришлёт письмо об ошибке триггера
    throw new Error('WB API вернул код ' + code);
  }
  var json = JSON.parse(resp.getContentText());
  var products = (json.data && json.data.products) ? json.data.products : [];
  var result = {};
  for (var i = 0; i < products.length; i++) {
    var p = products[i];
    var priceU = extractPriceU(p);
    if (priceU == null) continue;
    result[String(p.id)] = {
      price: Math.round(priceU) / 100, // копейки -> рубли
      name: p.name || ''
    };
  }
  return result;
}

/**
 * Достаёт цену в копейках, устойчиво к смене формата ответа WB.
 * Новый формат: sizes[0].price.product / total / basic.
 * Старый формат: salePriceU / priceU.
 */
function extractPriceU(p) {
  if (p.sizes && p.sizes.length) {
    var pr = p.sizes[0].price;
    if (pr) {
      if (pr.product != null) return pr.product;
      if (pr.total   != null) return pr.total;
      if (pr.basic   != null) return pr.basic;
    }
  }
  if (p.salePriceU != null) return p.salePriceU;
  if (p.priceU != null) return p.priceU;
  return null;
}

/**
 * Запись изменений в лист «Лог» (только факты изменения цены).
 */
function writeLog(ss, changes, now) {
  var log = ss.getSheetByName(SHEET_LOG);
  if (!log) log = ss.insertSheet(SHEET_LOG);
  if (log.getLastRow() === 0) {
    log.appendRow(['Дата/время', 'Артикул', 'Название', 'Старая цена', 'Новая цена', 'Изменение']);
  }
  for (var i = 0; i < changes.length; i++) {
    var ch = changes[i];
    log.appendRow([
      formatDate(now), ch.nm, ch.name, ch.oldPrice, ch.newPrice,
      (ch.delta > 0 ? '+' : '') + ch.delta + ' ₽'
    ]);
  }
}

/**
 * Сборка текста уведомления для Telegram.
 */
function buildMessage(changes) {
  var lines = ['📊 Изменение цен на Wildberries:', ''];
  for (var i = 0; i < changes.length; i++) {
    var ch = changes[i];
    var arrow = ch.delta > 0 ? '🔺' : '🔻';
    lines.push(arrow + ' ' + (ch.name || ch.nm) + ' (арт. ' + ch.nm + ')');
    lines.push('   ' + ch.oldPrice + ' ₽ → ' + ch.newPrice + ' ₽  (' +
               (ch.pct > 0 ? '+' : '') + ch.pct + '%)');
  }
  return lines.join('\n');
}

/**
 * Отправка сообщения в Telegram. Токен и chat_id — из Script Properties.
 */
function sendTelegram(text) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('TELEGRAM_TOKEN');
  var chatId = props.getProperty('TELEGRAM_CHAT_ID');
  if (!token || !chatId) {
    throw new Error('Не заданы TELEGRAM_TOKEN / TELEGRAM_CHAT_ID в Script Properties');
  }
  var url = 'https://api.telegram.org/bot' + token + '/sendMessage';
  UrlFetchApp.fetch(url, {
    method: 'post',
    muteHttpExceptions: true,
    payload: { chat_id: chatId, text: text }
  });
}

function formatDate(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

/**
 * Разовая функция: создаёт триггер запуска каждые 2 часа.
 * Запусти вручную ОДИН раз из редактора.
 */
function createTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkPrices') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('checkPrices').timeBased().everyHours(2).create();
}

/**
 * Вспомогательная функция: проверить связку с Telegram (тестовое сообщение).
 */
function testTelegram() {
  sendTelegram('✅ WB Price Monitor: связь с Telegram работает.');
}
