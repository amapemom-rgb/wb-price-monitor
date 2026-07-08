// @OnlyCurrentDoc
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
// A=Комментарии(0) B=Артикул(1) C=Метка(2) D=Название(3)
// E=Пред.цена(4) F=Тек.цена(5) G=Изменение(6) H=Обновлено(7)

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

  var range = sheet.getRange(2, 1, lastRow - 1, 11);
  var rows = range.getValues();

  // Собираем список артикулов
  var nmList = [];
  for (var i = 0; i < rows.length; i++) {
    var nm = String(rows[i][1]).trim();
    if (nm) nmList.push(nm);
  }
  if (nmList.length === 0) return;

  // Запрашиваем цены из официального API (в изолированном блоке try/catch)
  var officialPrices = fetchOfficialWbPrices();

  // Запрашиваем публичные цены батчами
  var prices = {}; // nm -> {price, name}
  for (var c = 0; c < nmList.length; c += CHUNK_SIZE) {
    var chunk = nmList.slice(c, c + CHUNK_SIZE);
    var data = fetchWbPrices(chunk);
    for (var key in data) prices[key] = data[key];
    Utilities.sleep(REQUEST_DELAY_MS);
  }

  // Защита: если список артикулов не пуст, но от API не получили ни одного результата —
  // скорее всего, изменился формат API Wildberries. Прерываемся с ошибкой.
  if (Object.keys(prices).length === 0) {
    throw new Error('Не удалось получить данные ни по одному артикулу. Возможно, изменилась структура ответа API Wildberries.');
  }

  var now = new Date();
  var changes = [];

  for (var j = 0; j < rows.length; j++) {
    var artikul = String(rows[j][1]).trim();
    if (!artikul) continue;

    var info = prices[artikul];
    if (!info) {
      // товар не найден — помечаем, но не роняем весь прогон
      rows[j][7] = 'нет данных ' + formatDate(now);
      continue;
    }

    // Автоматическое определение меток
    if (officialPrices) {
      var isOurProduct = officialPrices.hasOwnProperty(artikul);
      rows[j][2] = isOurProduct ? 'моё' : 'конкурент';
    }
    var isOur = (rows[j][2] === 'моё');

    var oldPrice = Number(rows[j][5]) || 0; // тек.цена прошлого прогона
    var newPrice = info.price;

    rows[j][3] = info.name;       // Название
    rows[j][4] = oldPrice;        // Предыдущая цена
    rows[j][5] = newPrice;        // Текущая цена
    rows[j][7] = formatDate(now); // Обновлено

    // Заполнение ЛК-цен и расчет СПП для наших товаров
    if (isOur) {
      if (officialPrices && officialPrices.hasOwnProperty(artikul)) {
        rows[j][8] = officialPrices[artikul];
      }
      var priceLk = Number(rows[j][8]) || 0;
      var tekPrice = newPrice;
      
      if (priceLk <= 0 || tekPrice <= 0) {
        rows[j][9] = 'уточняется';
      } else {
        var spp = (priceLk - tekPrice) / priceLk * 100;
        spp = Math.round(spp * 10) / 10;
        if (spp < 0 || spp > 95) {
          rows[j][9] = 'уточняется';
        } else {
          rows[j][10] = rows[j][9] || '';
          rows[j][9] = spp;
        }
      }
    } else {
      rows[j][8] = '';
      rows[j][9] = '';
      rows[j][10] = '';
    }

    if (oldPrice > 0 && newPrice !== oldPrice) {
      var delta = Math.round((newPrice - oldPrice) * 100) / 100;
      var pct = Math.round((delta / oldPrice) * 1000) / 10;
      rows[j][6] = "'" + (delta > 0 ? '+' : '') + delta + ' ₽ (' + (pct > 0 ? '+' : '') + pct + '%)';
      var comment = String(rows[j][0]).trim();
      changes.push({
        nm: artikul, name: info.name, comment: comment,
        oldPrice: oldPrice, newPrice: newPrice, delta: delta, pct: pct,
        isOur: isOur, spp: rows[j][9], prevSpp: rows[j][10]
      });
    } else if (oldPrice === 0) {
      rows[j][6] = 'базовая'; // первый прогон — фиксируем базу, алерт не шлём
    } else {
      rows[j][6] = 'без изменений';
    }
  }

  // Реконсиляция (согласование цен) для наших товаров перед отправкой
  var hasOurChange = false;
  var ourChangedNms = [];
  for (var i = 0; i < changes.length; i++) {
    if (changes[i].isOur) {
      hasOurChange = true;
      ourChangedNms.push(changes[i].nm);
    }
  }

  if (hasOurChange) {
    Logger.log('Обнаружено изменение по нашим товарам. Задержка 60 сек перед перепроверкой...');
    Utilities.sleep(60000);
    
    // Перезапрашиваем официальные и публичные цены
    var officialPrices2 = fetchOfficialWbPrices();
    var prices2 = {};
    for (var c = 0; c < ourChangedNms.length; c += CHUNK_SIZE) {
      var chunk = ourChangedNms.slice(c, c + CHUNK_SIZE);
      var data = fetchWbPrices(chunk);
      for (var key in data) prices2[key] = data[key];
      Utilities.sleep(REQUEST_DELAY_MS);
    }
    
    // Обновляем измененные строки
    for (var i = 0; i < changes.length; i++) {
      var ch = changes[i];
      if (ch.isOur) {
        var artikul = ch.nm;
        var rowIndex = -1;
        for (var j = 0; j < rows.length; j++) {
          if (String(rows[j][1]).trim() === artikul) {
            rowIndex = j;
            break;
          }
        }
        
        if (rowIndex !== -1) {
          var info2 = prices2[artikul];
          if (info2) {
            ch.newPrice = info2.price;
            rows[rowIndex][5] = info2.price; // Текущая цена
            
            var oldPrice = ch.oldPrice;
            var newPrice = info2.price;
            var delta = Math.round((newPrice - oldPrice) * 100) / 100;
            var pct = Math.round((delta / oldPrice) * 1000) / 10;
            rows[rowIndex][6] = "'" + (delta > 0 ? '+' : '') + delta + ' ₽ (' + (pct > 0 ? '+' : '') + pct + '%)';
            ch.delta = delta;
            ch.pct = pct;
          }
          
          if (officialPrices2 && officialPrices2.hasOwnProperty(artikul)) {
            rows[rowIndex][8] = officialPrices2[artikul]; // Цена ЛК
          }
          
          var priceLk = Number(rows[rowIndex][8]) || 0;
          var tekPrice = Number(rows[rowIndex][5]) || 0;
          
          if (priceLk <= 0 || tekPrice <= 0) {
            rows[rowIndex][9] = 'уточняется';
          } else {
            var spp = (priceLk - tekPrice) / priceLk * 100;
            spp = Math.round(spp * 10) / 10;
            if (spp < 0 || spp > 95) {
              rows[rowIndex][9] = 'уточняется';
            } else {
              rows[rowIndex][9] = spp;
            }
          }
          ch.spp = rows[rowIndex][9];
        }
      }
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
  var result = {};
  
  // Отфильтруем пустые артикулы
  var cleanNms = [];
  for (var i = 0; i < nmArray.length; i++) {
    var sku = nmArray[i];
    if (sku) {
      cleanNms.push(sku);
    }
  }
  
  if (cleanNms.length === 0) return result;
  
  // Wildberries позволяет отправлять в одном запросе до 100+ артикулов,
  // разделяя их точкой с запятой (;).
  var url = 'https://www.wildberries.ru/__internal/u-card/cards/v4/detail?appType=1&curr=rub&dest=' + WB_DEST +
            '&nm=' + cleanNms.join(';');
            
  var attempts = 3;
  for (var attempt = 1; attempt <= attempts; attempt++) {
    try {
      var resp = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }
      });
      var code = resp.getResponseCode();
      if (code === 200) {
        var data = JSON.parse(resp.getContentText());
        var products = data.products || [];
        for (var j = 0; j < products.length; j++) {
          var p = products[j];
          var priceU = extractPriceU(p);
          if (priceU != null) {
            result[String(p.id)] = {
              price: Math.round(priceU) / 100, // копейки -> рубли
              name: p.name || ''
            };
          }
        }
        break; // Успех, выходим из цикла попыток
      } else if (code === 429) {
        Logger.log('Получен код 429 для пакетного запроса. Попытка ' + attempt + ' из ' + attempts + '. Ожидание 5 сек...');
        Utilities.sleep(5000);
      } else {
        Logger.log('Batch API вернул код ' + code);
        break;
      }
    } catch (e) {
      Logger.log('Ошибка пакетного запроса: ' + e.message);
      break;
    }
  }
  return result;
}

/**
 * Запрос цен у официального API WB категории «Цены и скидки».
 * Возвращает объект nmID -> discountedPrice.
 */
function fetchOfficialWbPrices() {
  var apiKey = PropertiesService.getScriptProperties().getProperty('WB_API_KEY');
  if (!apiKey || apiKey === 'YOUR_WB_API_KEY_HERE') {
    Logger.log('Официальный токен WB_API_KEY не задан или содержит плейсхолдер.');
    return null;
  }
  
  var url = 'https://discounts-prices-api.wildberries.ru/api/v2/list/goods/filter?limit=1000&offset=0';
  try {
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': apiKey
      },
      muteHttpExceptions: true
    });
    var code = response.getResponseCode();
    if (code !== 200) {
      Logger.log('Официальный API вернул код ' + code + ': ' + response.getContentText());
      return null;
    }
    var json = JSON.parse(response.getContentText());
    if (!json.data || !json.data.listGoods) {
      Logger.log('Официальный API вернул неожиданный формат ответа.');
      return null;
    }
    
    var goods = json.data.listGoods;
    var result = {};
    for (var i = 0; i < goods.length; i++) {
      var item = goods[i];
      var nmId = String(item.nmID);
      var price = 0;
      if (item.sizes && item.sizes.length > 0) {
        price = Number(item.sizes[0].discountedPrice) || 0;
      }
      result[nmId] = price;
    }
    return result;
  } catch (e) {
    Logger.log('Изолированная ошибка при обращении к официальному API: ' + e.message);
    return null;
  }
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
      "'" + (ch.delta > 0 ? '+' : '') + ch.delta + ' ₽'
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
    var title = ch.comment || ch.name || ch.nm;
    lines.push(arrow + ' ' + title + ' (арт. ' + ch.nm + ')');
    if (ch.comment && ch.name) {
      lines.push('   ' + ch.name);
    }
    lines.push('   ' + ch.oldPrice + ' ₽ → ' + ch.newPrice + ' ₽  (' +
               (ch.pct > 0 ? '+' : '') + ch.pct + '%)');
    if (ch.isOur && ch.spp !== 'уточняется' && ch.spp !== '' && ch.spp !== undefined) {
      var prevSppValid = (ch.prevSpp !== 'уточняется' && ch.prevSpp !== '' && ch.prevSpp !== undefined);
      if (prevSppValid && ch.prevSpp !== ch.spp) {
        var arrowSpp = Number(ch.spp) > Number(ch.prevSpp) ? '▲' : '▼';
        lines.push('   СПП: ' + ch.prevSpp + '% → ' + ch.spp + '% (' + arrowSpp + ')');
      } else {
        lines.push('   СПП: ' + ch.spp + '%');
      }
    }
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
  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    muteHttpExceptions: true,
    payload: { chat_id: chatId, text: text }
  });
  
  var code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Ошибка отправки в Telegram (Код ' + code + '): ' + resp.getContentText());
  }
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
function testFetch() {
  Logger.log('Running testFetch');
}

function testOfficialApi() {
  var apiKey = PropertiesService.getScriptProperties().getProperty('WB_API_KEY');
  Logger.log('Using API Key: ' + (apiKey ? apiKey.substring(0, 10) + '...' : 'null'));
  var url = 'https://discounts-prices-api.wildberries.ru/api/v2/list/goods/filter?limit=10&offset=0';
  var response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'Authorization': apiKey
    },
    muteHttpExceptions: true
  });
  Logger.log('Response code: ' + response.getResponseCode());
  Logger.log('Response content: ' + response.getContentText().substring(0, 1000));
}

function testTelegram() {
  sendTelegram('✅ WB Price Monitor: связь с Telegram работает.');
}

/**
 * Инициализация структуры таблицы: создание листов «Товары» и «Лог» с заголовками.
 */
function initSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_PRODUCTS);
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName(SHEET_PRODUCTS);
  }
  // Записываем заголовки в первую строку (A1:K1)
  sheet.getRange(1, 1, 1, 11).setValues([[
    'Комментарии', 'Артикул', 'Метка', 'Название', 'Пред.цена', 'Тек.цена', 'Изменение', 'Обновлено', 'Цена ЛК', 'СПП %', 'Пред. СПП %'
  ]]);
  
  // Добавляем тестовый артикул, если таблица пуста
  if (sheet.getLastRow() < 2) {
    sheet.getRange('A2').setValue('Тест');
    sheet.getRange('B2').setValue('211690022');
    sheet.getRange('C2').setValue('Тест');
  }

  // Создаем лист Лог, если его нет
  var log = ss.getSheetByName(SHEET_LOG);
  if (!log) {
    log = ss.insertSheet(SHEET_LOG);
  }
  if (log.getLastRow() === 0) {
    log.appendRow(['Дата/время', 'Артикул', 'Название', 'Старая цена', 'Новая цена', 'Изменение']);
  }
}



