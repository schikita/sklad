var BOOKS = {
  counterparties: "1PBYZcU1TNnkeYfueQLoNDOY0Owou6crqEEIsSVhSKZg",
  cash: "1SlIUkncDyrhpvs7FPmtWY6zvX684nbiO-s0bVrzFOGQ",
  hr: "1Dy844ebEUk-H-MYeKTQpWN6Wsh-UIT05HaoxtfABSh8",
  products: "1mMZlLOu9XP2BTzs_frWq6uAKzoqC_TC3bBJ1JmWzlyY",
  client: "1s9Dx7eiCx9gsFeVAlqIEGN2pIgdDqqHBjnldKnXquIs"
};

var VITRINA = ["franchise_id", "source", "sheet", "name", "report_date", "money_in", "money_out", "campaign_id", "link", "payload"];
var RAW = ["franchise_id", "source", "sheet", "loaded_at", "entity", "payload"];
var ETL_STARTED = 0;

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function doGet(e) {
  var requested = "";
  if (e && e.parameter && e.parameter.franchise) requested = String(e.parameter.franchise);
  if (!/^[A-Za-z0-9_-]{0,40}$/.test(requested)) requested = "";
  var page = HtmlService.createTemplateFromFile("Index");
  page.franchise = requested;
  return page.evaluate()
    .setTitle("Пульт РемМонтаж")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  var book = registryBook();
  var body = e && e.postData ? String(e.postData.contents || "") : "";
  book.getSheetByName("Вход 1С").appendRow([new Date(), body.substring(0, 45000)]);
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function setupMasterRegistry() {
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty("REGISTRY_ID");
  var ss = null;
  if (existing) {
    try { ss = SpreadsheetApp.openById(existing); } catch (err) { existing = ""; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create("Мастер-реестр РемМонтаж");
    props.setProperty("REGISTRY_ID", ss.getId());
    ss.setSpreadsheetTimeZone("Europe/Moscow");
  }
  ensureShape(ss);
  seed(ss);
  try { runETL(); } catch (err) { Logger.log(err); }
  try { resetTriggers(); } catch (err) { Logger.log("триггер: " + err); }
  Logger.log(ss.getUrl());
  return ss.getUrl();
}

function resetTriggers() {
  var all = ScriptApp.getProjectTriggers();
  for (var i = 0; i < all.length; i++) {
    var fn = all[i].getHandlerFunction();
    if (fn === "runETL" || fn === "continueETL") ScriptApp.deleteTrigger(all[i]);
  }
  ScriptApp.newTrigger("runETL").timeBased().everyDays(1).atHour(6).create();
}

function runETL() {
  try { backupRegistry(); } catch (err) { Logger.log("бэкап: " + err); }
  PropertiesService.getScriptProperties().deleteProperty("ETL_CURSOR");
  return continueETL();
}

function continueETL() {
  ETL_STARTED = Date.now();
  var book = registryBook();
  var props = PropertiesService.getScriptProperties();
  var cursor = Number(props.getProperty("ETL_CURSOR") || "0");
  var franchises = indexFranchises(readFranchiseList(book));
  var sources = readSourceList(book);
  for (var i = cursor; i < sources.length; i++) {
    if (Date.now() - ETL_STARTED > 4.5 * 60 * 1000) {
      props.setProperty("ETL_CURSOR", String(i));
      scheduleContinue();
      logLine(book, "", "робот", "заметка", "Пауза на строке реестра " + (i + 1) + ", продолжение через минуту");
      return;
    }
    var src = sources[i];
    var fr = franchises[src.franchise];
    if (!fr || fr.status !== "АКТИВЕН") {
      logLine(book, src.franchise, src.name, "заметка", "Франшиза не активна, лист пропущен");
      continue;
    }
    if (src.status !== "Включать") {
      logLine(book, src.franchise, src.name, "заметка", "Статус «Исключить», лист пропущен");
      continue;
    }
    try { syncSheet(book, fr, src); }
    catch (err) { logLine(book, src.franchise, src.name, "ошибка", String(err)); }
  }
  var list = readFranchiseList(book);
  for (var f = 0; f < list.length; f++) {
    if (list[f].status === "АКТИВЕН") pullMoySklad(book, list[f]);
  }
  noteWarehouse(book);
  props.deleteProperty("ETL_CURSOR");
}

function scheduleContinue() {
  var all = ScriptApp.getProjectTriggers();
  for (var i = 0; i < all.length; i++) {
    if (all[i].getHandlerFunction() === "continueETL") ScriptApp.deleteTrigger(all[i]);
  }
  ScriptApp.newTrigger("continueETL").timeBased().after(60 * 1000).create();
}

function syncSheet(book, fr, src) {
  var ss = openBook(src.id);
  var sheet = ss.getSheetByName(src.sheet);
  if (!sheet) {
    logLine(book, fr.id, src.name, "предупреждение", "Лист «" + src.sheet + "» не найден");
    return;
  }
  var lastRow = Math.min(Math.max(sheet.getLastRow(), 1), 5000);
  var lastCol = Math.min(Math.max(sheet.getLastColumn(), 1), 30);
  var range = sheet.getRange(1, 1, lastRow, lastCol);
  var values = range.getValues();
  var formulas = range.getFormulas();
  var headerAt = src.headerRow ? Number(src.headerRow) - 1 : findHeader(values);
  var headers = values[headerAt].map(function (cell) { return String(cell || "").trim(); });
  var required = requiredHeader(book, src.blank);
  if (required && headers.indexOf(required) === -1) {
    logLine(book, fr.id, src.name, "предупреждение", "Нет столбца «" + required + "». Лист пропущен, остальные источники продолжают сбор");
    return;
  }
  var dict = loadDictionary(book);
  var fixes = correctionMap(book);
  var unknown = {};
  var badNumber = {};
  var rows = [];
  var rawRows = [];
  var loadedAt = Utilities.formatDate(new Date(), "Europe/Moscow", "yyyy-MM-dd HH:mm:ss");
  var emptyRun = 0;
  for (var r = headerAt + 1; r < values.length; r++) {
    var payload = {};
    var any = false;
    var firstUrl = "";
    for (var c = 0; c < headers.length; c++) {
      if (!headers[c]) continue;
      var parts = splitCell(formulas[r][c], values[r][c]);
      if (String(parts.text || "").trim() !== "" || parts.url) any = true;
      payload[headers[c]] = parts.text !== "" && parts.text !== null && parts.text !== undefined ? parts.text : parts.url;
      if (parts.url && !firstUrl) firstUrl = parts.url;
    }
    if (!any) {
      emptyRun++;
      if (emptyRun > 20) break;
      continue;
    }
    emptyRun = 0;
    var mapped = mapPayload(payload, dict, unknown, badNumber);
    if (firstUrl) mapped.link = firstUrl;
    var name = pickName(mapped, payload);
    var key = fr.id + "|" + src.id + "|" + src.sheet + "|" + name + "|" + (mapped.report_date || "");
    if (fixes[key]) {
      if (fixes[key].money_in !== undefined) mapped.money_in = fixes[key].money_in;
      if (fixes[key].money_out !== undefined) mapped.money_out = fixes[key].money_out;
    }
    rows.push([
      fr.id, src.name, src.sheet, name, mapped.report_date || "",
      mapped.money_in === undefined ? "" : mapped.money_in,
      mapped.money_out === undefined ? "" : mapped.money_out,
      mapped.campaign_id || "", mapped.link || "", JSON.stringify(payload)
    ]);
    rawRows.push([fr.id, src.name, src.sheet, loadedAt, "sheet", JSON.stringify(payload)]);
  }
  replacePairs(book.getSheetByName("Витрина"), VITRINA.length, src.name, src.sheet, rows);
  replacePairs(book.getSheetByName("Сырьё"), RAW.length, src.name, src.sheet, rawRows);
  var unknownNames = Object.keys(unknown);
  if (unknownNames.length) {
    logLine(book, fr.id, src.name, "предупреждение", "Незнакомые столбцы пропущены в витрине: " + unknownNames.join(", "));
  }
  var badNames = Object.keys(badNumber);
  if (badNames.length) {
    logLine(book, fr.id, src.name, "предупреждение", "Текст вместо числа очищен: " + badNames.join(", "));
  }
  var note = rows.length + " строк";
  if (!rows.length && src.blank === "Товары") note += ". Файл карточек ещё не передан";
  logLine(book, fr.id, src.name, "ок", note);
}

function pullMoySklad(book, fr) {
  var token = String(fr.token || "").replace(/^\s+|\s+$/g, "");
  if (!token) {
    logLine(book, fr.id, "МойСклад", "заметка", "Токен не задан, API пропущен");
    return;
  }
  if (!/^[A-Za-z0-9._\-]+$/.test(token) || token.length < 8) {
    logLine(book, fr.id, "МойСклад", "ошибка", "Токен не похож на ключ API");
    return;
  }
  var ping = msFetch(token, "/entity/organization?limit=1");
  if (ping.code === 401 || ping.code === 403) {
    book.getSheetByName("Управление Франшизами").getRange(fr.sheetRow, 4).setValue("ОШИБКА ТОКЕНА");
    logLine(book, fr.id, "МойСклад", "ошибка", "Ключ не принят");
    return;
  }
  if (ping.code !== 200) {
    logLine(book, fr.id, "МойСклад", "ошибка", "Ответ " + ping.code);
    return;
  }
  var org = {};
  try { org = (JSON.parse(ping.text).rows || [])[0] || {}; } catch (err) { org = {}; }
  book.getSheetByName("Управление Франшизами").getRange(fr.sheetRow, 5, 1, 2).setValues([[org.name || "", org.id || ""]]);
  var props = PropertiesService.getScriptProperties();
  var since = props.getProperty("MS_SINCE_" + fr.id) || "";
  var entities = ["counterparty", "product", "paymentin", "paymentout", "customerorder", "demand", "employee"];
  for (var i = 0; i < entities.length; i++) {
    var entity = entities[i];
    var collected = [];
    var offset = 0;
    var failed = false;
    while (offset <= 20000) {
      var path = "/entity/" + entity + "?limit=1000&offset=" + offset;
      if (entity === "paymentin" || entity === "paymentout" || entity === "demand" || entity === "customerorder") {
        path += "&expand=agent";
      }
      if (since) path += "&filter=" + encodeURIComponent("updated>" + since);
      var page = msFetch(token, path);
      if (page.code === 429) {
        Utilities.sleep(2000);
        page = msFetch(token, path);
      }
      if (page.code !== 200) {
        logLine(book, fr.id, entity, "ошибка", "Ответ " + page.code);
        failed = true;
        break;
      }
      var batch = [];
      try { batch = JSON.parse(page.text).rows || []; } catch (err) { batch = []; }
      for (var b = 0; b < batch.length; b++) collected.push(batch[b]);
      if (batch.length < 1000) break;
      offset += 1000;
    }
    if (!failed) writeMoySklad(book, fr, entity, collected, !since);
  }
  props.setProperty("MS_SINCE_" + fr.id, Utilities.formatDate(new Date(), "Europe/Moscow", "yyyy-MM-dd HH:mm:ss"));
  logLine(book, fr.id, "МойСклад", "ок", "API опрошен");
}

function writeMoySklad(book, fr, entity, items, replace) {
  var loadedAt = Utilities.formatDate(new Date(), "Europe/Moscow", "yyyy-MM-dd HH:mm:ss");
  var rawRows = [];
  var rows = [];
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var payload = JSON.stringify(item);
    if (payload.length > 45000) payload = JSON.stringify({ id: item.id, name: item.name, truncated: true });
    var agent = item.agent && item.agent.name ? item.agent.name : "";
    var name = item.name || agent || "";
    var moment = item.moment ? String(item.moment).substring(0, 10) : "";
    var sum = item.sum === undefined || item.sum === null ? "" : item.sum / 100;
    var moneyIn = entity === "paymentout" ? "" : sum;
    var moneyOut = entity === "paymentout" ? sum : "";
    if (entity !== "paymentin" && entity !== "paymentout") {
      moneyIn = "";
      moneyOut = "";
    }
    var link = item.meta && item.meta.uuidHref ? item.meta.uuidHref : "";
    rawRows.push([fr.id, "МойСклад", entity, loadedAt, entity, payload]);
    rows.push([fr.id, "МойСклад", entity, name, moment, moneyIn, moneyOut, "", link, payload]);
  }
  if (replace) {
    replacePairs(book.getSheetByName("Сырьё"), RAW.length, "МойСклад", entity, rawRows);
    replacePairs(book.getSheetByName("Витрина"), VITRINA.length, "МойСклад", entity, rows);
  } else if (rawRows.length) {
    appendRows(book.getSheetByName("Сырьё"), rawRows);
    appendRows(book.getSheetByName("Витрина"), rows);
  }
}

function noteWarehouse(book) {
  var project = String(book.getSheetByName("Служебное").getRange("B3").getValue() || "").replace(/^\s+|\s+$/g, "");
  var ads = String(book.getSheetByName("Служебное").getRange("B5").getValue() || "").replace(/^\s+|\s+$/g, "");
  if (!project) logLine(book, "", "BigQuery", "заметка", "Проект Google Cloud не задан. Свод лежит в листе «Витрина»");
  if (!ads) logLine(book, "", "Google Ads", "заметка", "Customer ID нет, расход рекламы не подтягивается");
}

function apiState() {
  var book = registryBook();
  var franchises = readFranchiseList(book);
  var byId = indexFranchises(franchises);
  var sources = [];
  readSourceList(book).forEach(function (src) {
    var fr = byId[src.franchise];
    if (!fr || fr.status === "ОТКЛЮЧЕН") return;
    sources.push({
      sheetRow: src.sheetRow,
      franchise: src.franchise,
      name: src.name,
      sheet: src.sheet,
      status: src.status,
      url: src.url,
      franchiseStatus: fr.status,
      token: fr.token ? "задан" : "пусто"
    });
  });
  franchises.forEach(function (fr) {
    if (fr.status === "ОТКЛЮЧЕН") return;
    sources.push({
      sheetRow: 0,
      franchise: fr.id,
      name: "МойСклад",
      sheet: "API",
      status: fr.token ? "Включать" : "Исключить",
      url: "https://online.moysklad.ru/app/",
      franchiseStatus: fr.status,
      token: fr.token ? "задан" : "пусто",
      virtual: true
    });
  });
  return {
    sources: sources,
    summary: summarize(book),
    journal: readJournal(book),
    dictionary: readDictionaryRows(book),
    registryUrl: book.getUrl(),
    adsCustomerId: String(book.getSheetByName("Служебное").getRange("B5").getValue() || "").replace(/^\s+|\s+$/g, "")
  };
}

function apiSourceStats(req) {
  var book = registryBook();
  var src = sourceByRow(book, req.sheetRow);
  return summarize(book, src);
}

function apiRows(req) {
  var book = registryBook();
  var src = sourceByRow(book, req.sheetRow);
  var values = book.getSheetByName("Витрина").getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    if (src.virtual) {
      if (values[i][1] !== "МойСклад") continue;
    } else if (values[i][1] !== src.name || values[i][2] !== src.sheet) continue;
    out.push({
      key: src.franchise + "|" + (src.id || values[i][2]) + "|" + values[i][2] + "|" + values[i][3] + "|" + values[i][4],
      name: values[i][3],
      report_date: values[i][4] instanceof Date
        ? Utilities.formatDate(values[i][4], "Europe/Moscow", "yyyy-MM-dd")
        : values[i][4],
      money_in: values[i][5],
      money_out: values[i][6],
      link: values[i][8],
      vitrinaRow: i + 1
    });
  }
  return out;
}

function apiSetSourceStatus(req) {
  if (!req || Number(req.sheetRow) < 2) return apiState();
  var allowed = { "Включать": true, "Исключить": true };
  if (!allowed[req.status]) return apiState();
  registryBook().getSheetByName("Источники").getRange(Number(req.sheetRow), 5).setValue(req.status);
  return apiState();
}

function apiSetFranchiseStatus(req) {
  var allowed = { "АКТИВЕН": true, "ЗАБЛОКИРОВАН": true, "ОТКЛЮЧЕН": true };
  if (!req || !allowed[req.status]) return apiState();
  var book = registryBook();
  var list = readFranchiseList(book);
  for (var i = 0; i < list.length; i++) {
    if (list[i].id === req.franchise) {
      book.getSheetByName("Управление Франшизами").getRange(list[i].sheetRow, 4).setValue(req.status);
    }
  }
  return apiState();
}

function apiSyncOne(req) {
  var book = registryBook();
  if (!req || !req.sheetRow) {
    readFranchiseList(book).forEach(function (fr) {
      if (fr.status === "АКТИВЕН") pullMoySklad(book, fr);
    });
    noteWarehouse(book);
    return apiState();
  }
  var src = sourceByRow(book, req.sheetRow);
  var fr = franchiseById(book, src.franchise);
  if (!fr || fr.status !== "АКТИВЕН" || src.status !== "Включать") {
    logLine(book, src.franchise, src.name, "заметка", "Сбор пропущен из-за статуса");
    return apiState();
  }
  syncSheet(book, fr, src);
  return apiState();
}

function apiAddSource(req) {
  var book = registryBook();
  var id = sheetIdFromUrl(req.url);
  var franchises = readFranchiseList(book);
  var franchise = franchises.length ? franchises[0].id : "FR-001";
  book.getSheetByName("Источники").appendRow([
    franchise,
    String(req.name || "Новый источник"),
    id,
    String(req.sheet || ""),
    "Включать",
    "",
    "https://docs.google.com/spreadsheets/d/" + id + "/edit"
  ]);
  return apiState();
}

function apiAddMap(req) {
  var header = String(req.header || "").replace(/^\s+|\s+$/g, "");
  var field = String(req.field || "").replace(/^\s+|\s+$/g, "");
  if (!header || !field) return apiState();
  registryBook().getSheetByName("Словарь полей").appendRow([header, field, req.type || "text"]);
  return apiState();
}

function apiSaveCorrection(req) {
  var book = registryBook();
  var sh = book.getSheetByName("Правки");
  var data = sh.getDataRange().getValues();
  var value = req.value;
  if (req.field === "money_in" || req.field === "money_out") {
    var parsed = asNumber(value);
    value = parsed === "" ? "" : parsed;
  }
  var found = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(req.key) && String(data[i][1]) === String(req.field)) found = i + 1;
  }
  if (found) sh.getRange(found, 3, 1, 2).setValues([[value, new Date()]]);
  else sh.appendRow([req.key, req.field, value, new Date()]);
  var col = req.field === "money_in" ? 6 : (req.field === "money_out" ? 7 : 0);
  if (col && req.vitrinaRow) book.getSheetByName("Витрина").getRange(Number(req.vitrinaRow), col).setValue(value);
  return { ok: true };
}

function registryBook() {
  var id = PropertiesService.getScriptProperties().getProperty("REGISTRY_ID");
  if (!id) throw new Error("Сначала запустите setupMasterRegistry");
  return SpreadsheetApp.openById(id);
}

function ensureShape(ss) {
  ensure(ss, "Управление Франшизами", ["ID_Франшизы", "Название", "Токен МойСклад", "Статус", "Организация МойСклад", "UUID", "Комментарий"]);
  ensure(ss, "Источники", ["ID_Франшизы", "Имя", "Идентификатор таблицы", "Имя листа", "Статус", "Тип бланка", "Ссылка"]);
  ensure(ss, "Словарь полей", ["Заголовок", "Поле хранилища", "Тип"]);
  ensure(ss, "Бланки", ["Тип бланка", "Обязательный заголовок"]);
  ensure(ss, "Журнал", ["Время", "Франшиза", "Источник", "Уровень", "Сообщение"]);
  ensure(ss, "Сырьё", RAW);
  ensure(ss, "Витрина", VITRINA);
  ensure(ss, "Правки", ["Ключ", "Поле", "Значение", "Когда"]);
  ensure(ss, "Вход 1С", ["Время", "JSON"]);
  ensure(ss, "Служебное", ["Параметр", "Значение"]);
  var first = ss.getSheetByName("Лист1");
  if (first && ss.getSheets().length > 1 && first.getLastRow() === 0) ss.deleteSheet(first);
  dropdown(ss.getSheetByName("Управление Франшизами"), 4, ["АКТИВЕН", "ЗАБЛОКИРОВАН", "ОТКЛЮЧЕН", "ОШИБКА ТОКЕНА"]);
  dropdown(ss.getSheetByName("Источники"), 5, ["Включать", "Исключить"]);
}

function ensure(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function dropdown(sh, col, list) {
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(true).build();
  sh.getRange(2, col, 200, 1).setDataValidation(rule);
}

function seed(ss) {
  var fr = ss.getSheetByName("Управление Франшизами");
  if (fr.getLastRow() < 2) {
    fr.appendRow(["FR-001", "ООО РемМонтаж", "", "АКТИВЕН", "", "", "Организация в МойСклад: РУСУГЛЕНЕФТЕГАЗ О."]);
  }
  var src = ss.getSheetByName("Источники");
  if (src.getLastRow() < 2) {
    var rows = [
      ["FR-001", "Контрагенты", BOOKS.counterparties, "Контрагенты", "Включать", "Контрагенты", bookUrl(BOOKS.counterparties)],
      ["FR-001", "Список контрагентов ДС", BOOKS.cash, "Список контрагентов", "Включать", "Список", bookUrl(BOOKS.cash)],
      ["FR-001", "Ген. Директор", BOOKS.hr, "Ген. Директор", "Включать", "Кадры", bookUrl(BOOKS.hr)],
      ["FR-001", "Бухгалтер", BOOKS.hr, "Бухгалтер", "Включать", "Кадры", bookUrl(BOOKS.hr)],
      ["FR-001", "Слесарь-ремонтник № 1", BOOKS.hr, "Слесарь-ремонтник № 1", "Включать", "Кадры", bookUrl(BOOKS.hr)],
      ["FR-001", "Товары", BOOKS.products, "Товары", "Включать", "Товары", bookUrl(BOOKS.products)]
    ];
    src.getRange(2, 1, rows.length, 7).setValues(rows);
  }
  var dict = ss.getSheetByName("Словарь полей");
  if (dict.getLastRow() < 2) {
    var maps = [
      ["Контрагенты", "name", "text"],
      ["Контрагент", "name", "text"],
      ["Наименование", "name", "text"],
      ["ФИО", "name", "text"],
      ["Сотрудник", "name", "text"],
      ["Дата", "report_date", "date"],
      ["Поступление", "money_in", "money"],
      ["Списание", "money_out", "money"],
      ["Цена", "money_in", "money"],
      ["Документы", "link", "link"],
      ["Ссылка", "link", "link"],
      ["ID кампании", "campaign_id", "text"],
      ["Кампания", "campaign_id", "text"]
    ];
    dict.getRange(2, 1, maps.length, 3).setValues(maps);
  }
  var blanks = ss.getSheetByName("Бланки");
  if (blanks.getLastRow() < 2) {
    blanks.getRange(2, 1, 4, 2).setValues([
      ["Контрагенты", "Контрагенты"],
      ["Товары", "Наименование"],
      ["Движение", "Дата"],
      ["Кадры", ""]
    ]);
  }
  var service = ss.getSheetByName("Служебное");
  if (service.getRange("A1").getValue() !== "Робот") {
    service.getRange(1, 1, 12, 2).setValues([
      ["Робот", "alexsklad083@gmail.com"],
      ["Доступ сотрудника", "Читатель на адрес робота"],
      ["Проект BigQuery", ""],
      ["Регион", "EU"],
      ["Customer ID Google Ads", ""],
      ["Реестр", ss.getUrl()],
      ["Ключ склейки", "дата и id кампании"],
      ["Прибыль на экране", "поступление минус расход рекламы, без себестоимости"],
      ["Кто видит пульт", "владелец проекта"],
      ["Обновление", "каждый день в 06:00 и по кнопке"],
      ["Себестоимость", "в прибыль не входит, пока нет отдельного решения"],
      ["Главный аккаунт", "alexsklad083@gmail.com"]
    ]);
  } else {
    service.getRange("B6").setValue(ss.getUrl());
  }
  if (!src.getRange("H1").getValue()) src.getRange("H1").setValue("Строка заголовка");
  var seen = {};
  var existingSources = src.getDataRange().getValues();
  for (var s = 1; s < existingSources.length; s++) seen[String(existingSources[s][2]) + "|" + String(existingSources[s][3])] = true;
  var clientRows = [];
  if (!seen[BOOKS.client + "|Контрагенты"]) {
    clientRows.push(["FR-001", "Контрагенты Русугленефтегаз", BOOKS.client, "Контрагенты", "Включать", "Контрагенты", bookUrl(BOOKS.client), ""]);
  }
  if (!seen[BOOKS.client + "|Накладные 2026г"]) {
    clientRows.push(["FR-001", "Накладные 2026", BOOKS.client, "Накладные 2026г", "Включать", "Накладные", bookUrl(BOOKS.client), 4]);
  }
  if (clientRows.length) src.getRange(src.getLastRow() + 1, 1, clientRows.length, 8).setValues(clientRows);
  var placed = src.getDataRange().getValues();
  for (var u = 1; u < placed.length; u++) {
    if (String(placed[u][2]) === BOOKS.client && String(placed[u][3]) === "Накладные 2026г" && Number(placed[u][7]) !== 4) {
      src.getRange(u + 1, 8).setValue(4);
    }
  }
  var haveMap = {};
  var dictValues = dict.getDataRange().getValues();
  for (var m = 1; m < dictValues.length; m++) haveMap[String(dictValues[m][0]).toLowerCase()] = true;
  var moreMaps = [
    ["Полное наименование поставщика", "name", "text"],
    ["Дата документа", "report_date", "date"],
    ["Сумма документа в рублях", "money_out", "money"]
  ].filter(function (row) { return !haveMap[row[0].toLowerCase()]; });
  if (moreMaps.length) dict.getRange(dict.getLastRow() + 1, 1, moreMaps.length, 3).setValues(moreMaps);
}

function bookUrl(id) {
  return "https://docs.google.com/spreadsheets/d/" + id + "/edit";
}

function readFranchiseList(book) {
  var values = book.getSheetByName("Управление Франшизами").getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    out.push({
      sheetRow: i + 1,
      id: String(values[i][0]),
      name: String(values[i][1] || ""),
      token: String(values[i][2] || ""),
      status: String(values[i][3] || "")
    });
  }
  return out;
}

function readSourceList(book) {
  var values = book.getSheetByName("Источники").getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i][2]) continue;
    out.push({
      sheetRow: i + 1,
      franchise: String(values[i][0] || ""),
      name: String(values[i][1] || ""),
      id: String(values[i][2] || ""),
      sheet: String(values[i][3] || ""),
      status: String(values[i][4] || ""),
      blank: String(values[i][5] || ""),
      url: String(values[i][6] || ""),
      headerRow: Number(values[i][7] || 0)
    });
  }
  return out;
}

function indexFranchises(list) {
  var map = {};
  for (var i = 0; i < list.length; i++) map[list[i].id] = list[i];
  return map;
}

function franchiseById(book, id) {
  return indexFranchises(readFranchiseList(book))[id] || null;
}

function sourceByRow(book, sheetRow) {
  if (!sheetRow) {
    var fr = readFranchiseList(book)[0];
    return { sheetRow: 0, franchise: fr ? fr.id : "", name: "МойСклад", sheet: "API", status: "Включать", id: "", url: "https://online.moysklad.ru/app/", virtual: true };
  }
  var list = readSourceList(book);
  for (var i = 0; i < list.length; i++) if (list[i].sheetRow === Number(sheetRow)) return list[i];
  throw new Error("Источник не найден");
}

function readDictionaryRows(book) {
  var values = book.getSheetByName("Словарь полей").getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    out.push({ header: String(values[i][0]), field: String(values[i][1] || ""), type: String(values[i][2] || "text") });
  }
  return out;
}

function loadDictionary(book) {
  var map = {};
  readDictionaryRows(book).forEach(function (row) {
    map[row.header.toLowerCase()] = row;
  });
  return map;
}

function requiredHeader(book, blank) {
  if (!blank) return "";
  var values = book.getSheetByName("Бланки").getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === blank) return String(values[i][1] || "").replace(/^\s+|\s+$/g, "");
  }
  return "";
}

function correctionMap(book) {
  var values = book.getSheetByName("Правки").getDataRange().getValues();
  var map = {};
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    var key = String(values[i][0]);
    if (!map[key]) map[key] = {};
    map[key][String(values[i][1])] = values[i][2];
  }
  return map;
}

function readJournal(book) {
  var values = book.getSheetByName("Журнал").getDataRange().getValues();
  var out = [];
  var start = Math.max(1, values.length - 40);
  for (var i = values.length - 1; i >= start; i--) {
    if (!values[i][0] && !values[i][4]) continue;
    var time = values[i][0];
    if (time instanceof Date) time = Utilities.formatDate(time, "Europe/Moscow", "dd.MM HH:mm");
    out.push({ time: String(time || ""), level: String(values[i][3] || ""), source: String(values[i][2] || ""), message: String(values[i][4] || "") });
  }
  return out;
}

function summarize(book, only) {
  var active = {};
  var msLive = {};
  var franchises = indexFranchises(readFranchiseList(book));
  readSourceList(book).forEach(function (src) {
    var fr = franchises[src.franchise];
    if (src.status === "Включать" && fr && fr.status === "АКТИВЕН") active[src.name + "|" + src.sheet] = true;
  });
  for (var id in franchises) {
    if (!Object.prototype.hasOwnProperty.call(franchises, id)) continue;
    var fr = franchises[id];
    if (fr.status === "АКТИВЕН" && String(fr.token || "").replace(/^\s+|\s+$/g, "")) msLive[fr.id] = true;
  }
  var values = book.getSheetByName("Витрина").getDataRange().getValues();
  var rows = 0;
  var moneyIn = 0;
  var moneyOut = 0;
  for (var i = 1; i < values.length; i++) {
    var key = values[i][1] + "|" + values[i][2];
    var fromApi = values[i][1] === "МойСклад" && msLive[String(values[i][0] || "")];
    if (only && only.virtual) {
      if (values[i][1] !== "МойСклад") continue;
    } else if (only) {
      if (values[i][1] !== only.name || values[i][2] !== only.sheet) continue;
    } else if (!active[key] && !fromApi && values[i][1] !== "ОБРАЗЕЦ") continue;
    rows++;
    moneyIn += Number(values[i][5] || 0);
    moneyOut += Number(values[i][6] || 0);
  }
  return { rows: rows, moneyIn: moneyIn, moneyOut: moneyOut };
}

function findHeader(values) {
  var best = 0;
  var bestScore = -1;
  var limit = Math.min(values.length, 8);
  for (var r = 0; r < limit; r++) {
    var score = 0;
    for (var c = 0; c < values[r].length; c++) if (String(values[r][c] || "").replace(/^\s+|\s+$/g, "")) score++;
    if (score > bestScore) {
      bestScore = score;
      best = r;
    }
  }
  return best;
}

function mapPayload(payload, dict, unknown, badNumber) {
  var mapped = {};
  for (var header in payload) {
    if (!Object.prototype.hasOwnProperty.call(payload, header)) continue;
    var rule = dict[String(header).toLowerCase()];
    if (!rule) {
      unknown[header] = true;
      continue;
    }
    var value = payload[header];
    if (rule.type === "money" || rule.type === "number") {
      var num = asNumber(value);
      if (num === "" && String(value || "").replace(/^\s+|\s+$/g, "")) badNumber[header] = true;
      mapped[rule.field] = num;
    } else if (rule.type === "date") {
      var date = asDate(value);
      if (!date && String(value || "").replace(/^\s+|\s+$/g, "")) badNumber[header] = true;
      mapped[rule.field] = date;
    } else if (rule.type === "link") {
      mapped[rule.field] = String(value || "");
    } else {
      mapped[rule.field] = value;
    }
  }
  return mapped;
}

function pickName(mapped, payload) {
  if (mapped.name) return String(mapped.name);
  var best = "";
  for (var key in payload) {
    if (!Object.prototype.hasOwnProperty.call(payload, key)) continue;
    var value = String(payload[key] || "");
    if (value.indexOf("http") === 0) continue;
    if (value.length > best.length) best = value;
  }
  return best;
}

function splitCell(formula, value) {
  var match = String(formula || "").match(/https?:\/\/[^"']+/);
  var text = value;
  if (value instanceof Date) text = Utilities.formatDate(value, "Europe/Moscow", "yyyy-MM-dd");
  if (text === null || text === undefined) text = "";
  return { text: text, url: match ? match[0] : "" };
}

function asNumber(value) {
  if (typeof value === "number") return value;
  var text = String(value || "").replace(/\u00a0/g, " ").replace(/\s/g, "").replace(",", ".");
  text = text.replace(/[^0-9.\-]/g, "").replace(/\.+$/g, "");
  if (!text) return "";
  if (!/^-?\d+(\.\d+)?$/.test(text)) return "";
  return Number(text);
}

function asDate(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, "Europe/Moscow", "yyyy-MM-dd");
  var text = String(value || "").replace(/^\s+|\s+$/g, "").replace(/\s*г\.?$/i, "").replace(/^\s+|\s+$/g, "");
  var match = text.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (match) return match[3] + "-" + ("0" + match[2]).slice(-2) + "-" + ("0" + match[1]).slice(-2);
  if (/^\d{4}-\d{2}-\d{2}/.test(text)) return text.substring(0, 10);
  return "";
}

function replacePairs(sheet, width, sourceName, sheetName, rows) {
  var data = sheet.getDataRange().getValues();
  var header = data[0] && data[0][0] ? data[0] : (sheet.getName() === "Витрина" ? VITRINA : RAW);
  var kept = [header.slice(0, width)];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === sourceName && data[i][2] === sheetName) continue;
    if (!data[i][0] && !data[i][1]) continue;
    kept.push(pad(data[i], width));
  }
  for (var r = 0; r < rows.length; r++) kept.push(pad(rows[r], width));
  sheet.clearContents();
  sheet.getRange(1, 1, kept.length, width).setValues(kept);
  sheet.setFrozenRows(1);
}

function appendRows(sheet, rows) {
  if (!rows.length) return;
  var start = Math.max(sheet.getLastRow(), 1) + 1;
  sheet.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
}

function pad(row, width) {
  var out = [];
  for (var i = 0; i < width; i++) out.push(row[i] === undefined || row[i] === null ? "" : row[i]);
  return out;
}

function logLine(book, franchise, source, level, message) {
  book.getSheetByName("Журнал").appendRow([new Date(), franchise, source, level, message]);
}

function openBook(id) {
  var last = null;
  for (var i = 0; i < 5; i++) {
    try { return SpreadsheetApp.openById(id); }
    catch (err) {
      last = err;
      Utilities.sleep(1000 * (i + 1));
    }
  }
  throw last;
}

function msFetch(token, path) {
  var response = UrlFetchApp.fetch("https://api.moysklad.ru/api/remap/1.2" + path, {
    method: "get",
    muteHttpExceptions: true,
    headers: {
      Authorization: "Bearer " + token,
      Accept: "application/json;charset=utf-8"
    }
  });
  Utilities.sleep(250);
  return { code: response.getResponseCode(), text: response.getContentText() };
}

function sheetIdFromUrl(url) {
  var match = String(url || "").match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match) return match[1];
  var plain = String(url || "").replace(/^\s+|\s+$/g, "");
  if (/^[a-zA-Z0-9-_]{20,}$/.test(plain)) return plain;
  throw new Error("Не вижу ссылку на Google Таблицу");
}

function journalAdvice(level, message) {
  var text = String(message || "");
  if (!text) return "";
  var rules = [
    ["Ключ не принят|Токен не похож", "Проблема в листе «Управление Франшизами», столбец токена. МойСклад ключ отверг, статус франшизы — ОШИБКА ТОКЕНА. Новый ключ берут в МойСклад: Настройки, Токен. Старый ключ повторно не показывается."],
    ["Токен не задан", "Столбец токена пуст. МойСклад не опрашивается. Таблицы из «Источники» при этом собираются."],
    ["не найден", "Имя листа в «Источники», столбец D, не найдено в таблице по ссылке столбца G. Исправьте имя листа или ссылку."],
    ["Нет столбца", "В исходном листе нет обязательного столбца из «Бланки». Верните столбец в источник или исправьте бланк. Остальные источники этот пропуск не останавливает."],
    ["Незнакомые столбцы", "Эти заголовки не занесены в «Словарь полей», в витрину они не легли. Сами строки источника собраны. Словарь — лист разработчика."],
    ["Текст вместо числа", "В источнике текст там, где ждали число или дату. В витрину по этой ячейке легла пустота. Правьте исходную ячейку, не витрину."],
    ["Франшиза не активна|Сбор пропущен из-за статуса", "Сбор выключен статусом: франшиза не АКТИВЕН или источник не Включать. Это не порча данных."],
    ["Исключить", "В «Источники» статус Исключить. Робот эту строку не читает."],
    ["Проект Google Cloud не задан", "Ячейка B3 листа «Служебное» пуста. Свод остаётся на «Витрине»."],
    ["Customer ID нет", "Ячейка B5 листа «Служебное» пуста. Расход рекламы брать неоткуда."],
    ["Ответ 429", "МойСклад ограничил частоту запросов. Эта сущность за проход не дописана. Следующий запуск повторит запрос."],
    ["Ответ ", "МойСклад ответил не 200. Эта сущность в витрину за проход не записана. 401 и 403 — ключ. Иной код — повтор с пульта."],
    ["Пауза на строке", "Робот упёрся в лимит времени и продолжит сам через минуту. Данные не испорчены."],
    ["Файл карточек ещё не передан", "В источнике товаров нет строк. Это пустой файл, не обрыв связи."]
  ];
  for (var i = 0; i < rules.length; i++) {
    if (new RegExp(rules[i][0]).test(text)) return rules[i][1];
  }
  if (level === "ошибка") return "Красная строка. Франшиза — столбец B, источник — столбец C, текст — столбец E. Цифры правят в источнике или в листе «Правки», не здесь.";
  if (level === "предупреждение") return "Сбор не остановлен. Часть листа или столбцов пропущена. Что именно — в столбце E.";
  if (level === "заметка") return "Это не сбой. Робот пишет, почему шаг пропущен или какой настройки не хватает.";
  return "";
}

function sampleTable() {
  return [
    ["ТЕСТ. Оплата от заказчика", "2026-10-01", 15000, ""],
    ["ТЕСТ. Оплата поставщику", "2026-10-02", "", 2500]
  ];
}

function backupTitlePrefix() {
  return "Бэкап Мастер-реестр РемМонтаж ";
}

function chooseBackupsToTrash(entries, keep) {
  var prefix = backupTitlePrefix();
  var mine = [];
  var limit = keep === undefined ? 14 : keep;
  for (var i = 0; i < entries.length; i++) {
    if (String(entries[i].name || "").indexOf(prefix) === 0) mine.push(entries[i]);
  }
  mine.sort(function (a, b) {
    var at = a.created instanceof Date ? a.created.getTime() : Number(a.created);
    var bt = b.created instanceof Date ? b.created.getTime() : Number(b.created);
    return bt - at;
  });
  var trash = [];
  for (var k = limit; k < mine.length; k++) trash.push(mine[k]);
  return trash;
}

function backupFolder() {
  var name = "Бэкапы РемМонтаж";
  var found = DriveApp.getFoldersByName(name);
  if (found.hasNext()) return found.next();
  return DriveApp.createFolder(name);
}

function backupRegistry() {
  var book = registryBook();
  var stamp = Utilities.formatDate(new Date(), "Europe/Moscow", "yyyy-MM-dd HH-mm");
  var title = backupTitlePrefix() + stamp;
  var folder = backupFolder();
  var copy = DriveApp.getFileById(book.getId()).makeCopy(title, folder);
  var files = [];
  var it = folder.getFiles();
  while (it.hasNext()) {
    var file = it.next();
    files.push({ name: file.getName(), created: file.getDateCreated(), file: file });
  }
  var trash = chooseBackupsToTrash(files, 14);
  for (var i = 0; i < trash.length; i++) trash[i].file.setTrashed(true);
  logLine(book, "", "бэкап", "заметка", "Копия сохранена: " + title);
  return copy.getUrl();
}

function placeCustomerSample() {
  var book = registryBook();
  var table = sampleTable();
  var sh = book.getSheetByName("Образец");
  if (!sh) sh = book.insertSheet("Образец");
  sh.getRange(1, 1, 1, 4).setValues([["Наименование", "Дата", "Поступление", "Списание"]]);
  sh.getRange(2, 1, table.length, 4).setValues(table);
  sh.getRange(2, 1, table.length, 4).setBackground("#FFF4CC");
  sh.getRange(5, 1).setValue("Образец для заказчика. Это не факт оплаты. В источнике списание 2500. На витрине 2600, потому что лист «Правки» держит замену. Перед каждым сбором книга копируется в папку «Бэкапы РемМонтаж», хранятся 14 копий.");
  var id = book.getId();
  var sources = book.getSheetByName("Источники");
  var have = sources.getDataRange().getValues();
  var seen = false;
  for (var i = 1; i < have.length; i++) {
    if (String(have[i][2]) === id && String(have[i][3]) === "Образец") seen = true;
  }
  if (!seen) {
    sources.appendRow([
      "FR-001", "ОБРАЗЕЦ", id, "Образец", "Включать", "",
      "https://docs.google.com/spreadsheets/d/" + id + "/edit#gid=0", ""
    ]);
  }
  var key = "FR-001|" + id + "|Образец|ТЕСТ. Оплата поставщику|2026-10-02";
  var fixes = book.getSheetByName("Правки");
  var fixRows = fixes.getDataRange().getValues();
  var found = 0;
  for (var f = 1; f < fixRows.length; f++) {
    if (String(fixRows[f][0]) === key && String(fixRows[f][1]) === "money_out") found = f + 1;
  }
  if (found) fixes.getRange(found, 3, 1, 2).setValues([[2600, new Date()]]);
  else fixes.appendRow([key, "money_out", 2600, new Date()]);
  var rows = [
    ["FR-001", "ОБРАЗЕЦ", "Образец", table[0][0], table[0][1], 15000, "", "", "", JSON.stringify({ sample: true, "Наименование": table[0][0], "Поступление": 15000 })],
    ["FR-001", "ОБРАЗЕЦ", "Образец", table[1][0], table[1][1], "", 2600, "", "", JSON.stringify({ sample: true, "Наименование": table[1][0], "Списание": 2500, "Правка": 2600 })]
  ];
  var loadedAt = Utilities.formatDate(new Date(), "Europe/Moscow", "yyyy-MM-dd HH:mm:ss");
  var raw = [
    ["FR-001", "ОБРАЗЕЦ", "Образец", loadedAt, "sheet", rows[0][9]],
    ["FR-001", "ОБРАЗЕЦ", "Образец", loadedAt, "sheet", rows[1][9]]
  ];
  replacePairs(book.getSheetByName("Витрина"), VITRINA.length, "ОБРАЗЕЦ", "Образец", rows);
  replacePairs(book.getSheetByName("Сырьё"), RAW.length, "ОБРАЗЕЦ", "Образец", raw);
  var memo = book.getSheetByName("Памятка");
  if (memo) {
    var memoRows = memo.getDataRange().getValues();
    var marked = false;
    for (var m = 0; m < memoRows.length; m++) if (String(memoRows[m][0]) === "ОБРАЗЕЦ") marked = true;
    if (!marked) {
      memo.appendRow(["ОБРАЗЕЦ", "Две строки с пометкой ТЕСТ. Это не факт оплаты.", "Источник «Образец»: поступление 15 000 и списание 2 500.", "На витрине списание 2 600. Разницу держит лист «Правки».", "Перед сбором в 06:00 робот кладёт копию книги в папку «Бэкапы РемМонтаж». Хранятся 14 копий.", "Не принимать эти суммы за кассу и не стирать пометку ТЕСТ."]);
    }
  }
  logLine(book, "FR-001", "ОБРАЗЕЦ", "заметка", "Образец для заказчика: две строки ТЕСТ. Это не факт оплаты.");
  return { rows: 2, receipt: 15000, sourceWriteOff: 2500, vitrinaWriteOff: 2600 };
}

function imitationProducts() {
  var note = "Образец для заказчика. Это не карточка склада.";
  return [
    ["ИМИТАЦИЯ. Кабель ВВГ 3х2.5", "IM-001", "ИМИТАЦИЯ. Кабель", 86, "м", note, "", ""],
    ["ИМИТАЦИЯ. Автомат 16А", "IM-002", "ИМИТАЦИЯ. Электрика", 420, "шт", note, "", ""],
    ["ИМИТАЦИЯ. Труба ПНД 32", "IM-003", "ИМИТАЦИЯ. Труба", 95, "м", note, "", ""],
    ["ИМИТАЦИЯ. Светильник LED 36Вт", "IM-004", "ИМИТАЦИЯ. Свет", 2100, "шт", note, "", ""],
    ["ИМИТАЦИЯ. Крепёж М8", "IM-005", "ИМИТАЦИЯ. Крепёж", 12, "шт", note, "", ""]
  ];
}

function imitationCashTargets(headerRows) {
  var top = headerRows[0] || [];
  var mid = headerRows[1] || [];
  var amountIn = 0;
  var amountOut = 0;
  var nameIn = 0;
  var nameOut = 0;
  var width = Math.max(top.length, mid.length, (headerRows[2] || []).length);
  for (var i = 0; i < width; i++) {
    var topCell = String(top[i] || "").toLowerCase();
    var midCell = String(mid[i] || "").toLowerCase();
    var subCell = String((headerRows[2] || [])[i] || "").toLowerCase();
    var blocked = topCell.indexOf("валют") >= 0 || midCell.indexOf("валют") >= 0 || topCell.indexOf("остаток") >= 0 || midCell.indexOf("остаток") >= 0 || subCell.indexOf("остаток") >= 0 || topCell.indexOf("гсм") >= 0;
    if (!blocked && !amountIn && topCell.indexOf("общая сумма поступления") >= 0) amountIn = i + 1;
    if (!blocked && !amountOut && topCell.indexOf("общая сумма списания") >= 0) amountOut = i + 1;
    if (!nameIn && midCell.indexOf("контрагенты") === -1 && midCell.indexOf("контрагент") >= 0 && midCell.indexOf("вход") >= 0) nameIn = i + 1;
    if (!nameOut && midCell.indexOf("контрагент") >= 0 && midCell.indexOf("исх") >= 0) nameOut = i + 1;
  }
  if (!amountIn && nameIn > 1) amountIn = nameIn - 1;
  return { amountIn: amountIn, amountOut: amountOut, nameIn: nameIn, nameOut: nameOut };
}

function placeImitationSample() {
  var table = imitationProducts();
  var goods = SpreadsheetApp.openById(BOOKS.products).getSheetByName("Товары");
  if (String(goods.getRange(2, 1).getValue() || "").indexOf("ИМИТАЦИЯ") !== 0) {
    goods.getRange(2, 1, table.length, 8).setValues(table);
    goods.getRange(2, 1, table.length, 8).setBackground("#FFF4CC");
  }
  var stamp = "Пять жёлтых строк с пометкой ИМИТАЦИЯ — образец. Это не склад и не файл заказчика. Перед настоящим файлом эти строки убрать.";
  if (String(goods.getRange(1, 1).getNote() || "").indexOf("ИМИТАЦИЯ") === -1) goods.getRange(1, 1).setNote(stamp);
  var registry = registryBook();
  var rows = [];
  for (var p = 0; p < table.length; p++) {
    rows.push(["FR-001", "ИМИТАЦИЯ", "Товары", table[p][0], "", "", "", "", "", JSON.stringify({ sample: true, imitation: true, "Цена": table[p][3] })]);
  }
  replacePairs(registry.getSheetByName("Витрина"), VITRINA.length, "ИМИТАЦИЯ", "Товары", rows);
  var memo = registry.getSheetByName("Памятка");
  if (memo) {
    var memoRows = memo.getDataRange().getValues();
    var marked = false;
    for (var m = 0; m < memoRows.length; m++) if (String(memoRows[m][0]) === "ИМИТАЦИЯ") marked = true;
    if (!marked) memo.appendRow(["ИМИТАЦИЯ", "Пять жёлтых товаров. Это не склад.", "Октябрь не заполнен: колонка поступления подписана как остаток кассы, колонка списания — как документы.", "Оплата образца — две строки ТЕСТ на листе «Образец».", "Перед файлом заказчика жёлтые строки убрать.", ""]);
  }
  var service = registry.getSheetByName("Служебное");
  if (service && String(service.getRange("A13").getValue() || "") === "Отчет Студии данных" && String(service.getRange("B13").getValue() || "") === "") {
    service.getRange("B13").setValue("https://datastudio.google.com/reporting/02d07e4d-d231-4bf9-9f8e-fc2c3c615e43");
  }
  logLine(registry, "FR-001", "ИМИТАЦИЯ", "заметка", "Имитация для образца: 5 товаров. Октябрь не тронут. Это не склад и не факт оплаты.");
  return { products: table.length, cash: "skipped" };
}

function reportTotals(book) {
  var sum = summarize(book);
  return { rows: sum.rows, moneyIn: sum.moneyIn, moneyOut: sum.moneyOut, ads: 0, profit: sum.moneyIn };
}

function reportFormulaNames(dialect, modern) {
  var ru = dialect.iff === "ЕСЛИ";
  return {
    sep: dialect.sep,
    sum: ru ? "СУММ" : "SUM",
    sumproduct: ru ? "СУММПРОИЗВ" : "SUMPRODUCT",
    n: "N",
    map: modern === "ru" ? "КАРТА" : "MAP",
    lambda: modern === "ru" ? "ЛЯМБДА" : "LAMBDA",
    if: dialect.iff,
    iferror: ru ? "ЕСЛИОШИБКА" : "IFERROR",
    vlookup: ru ? "ВПР" : "VLOOKUP",
    countifs: ru ? "СЧЁТЕСЛИМН" : "COUNTIFS",
    filter: ru ? "ФИЛЬТР" : "FILTER",
    no: ru ? "ЛОЖЬ" : "FALSE"
  };
}

function reportChartFormulas(dialect, modern) {
  var n = reportFormulaNames(dialect, modern);
  var s = n.sep;
  var flag = "=" + n.map + "(Витрина!A2:A5000" + s + "Витрина!B2:B5000" + s + "Витрина!C2:C5000" + s +
    n.lambda + "(code" + s + "src" + s + "sheet" + s +
    n.if + "(code=\"\"" + s + "0" + s +
    n.if + "(src=\"ОБРАЗЕЦ\"" + s + "1" + s +
    n.if + "(" + n.iferror + "(" + n.vlookup + "(code" + s + "'Управление Франшизами'!A:D" + s + "4" + s + n.no + ")" + s + "\"\")<>\"АКТИВЕН\"" + s + "0" + s +
    n.if + "(src=\"МойСклад\"" + s + "1" + s +
    n.if + "(" + n.countifs + "('Источники'!A:A" + s + "code" + s + "'Источники'!B:B" + s + "src" + s + "'Источники'!D:D" + s + "sheet" + s + "'Источники'!E:E" + s + "\"Включать\")>0" + s + "1" + s + "0)))))))";
  return {
    flag: flag,
    rows: "=" + n.sum + "(K2:K5000)",
    moneyIn: "=" + n.sumproduct + "(K2:K5000" + s + n.n + "(Витрина!F2:F5000))",
    moneyOut: "=" + n.sumproduct + "(K2:K5000" + s + n.n + "(Витрина!G2:G5000))",
    profit: "=B6-B8",
    inBar: "=B6",
    outBar: "=B7",
    profitBar: "=B9",
    detailName: "=" + n.iferror + "(" + n.filter + "(Витрина!D2:D5000" + s + "K2:K5000=1)" + s + "\"\")",
    detailIn: "=" + n.iferror + "(" + n.filter + "(Витрина!F2:F5000" + s + "K2:K5000=1)" + s + "\"\")",
    detailOut: "=" + n.iferror + "(" + n.filter + "(Витрина!G2:G5000" + s + "K2:K5000=1)" + s + "\"\")",
    detailDate: "=" + n.iferror + "(" + n.filter + "(Витрина!E2:E5000" + s + "K2:K5000=1)" + s + "\"\")"
  };
}

function reportDialectAttempts(sheet) {
  var detected = formulaDialect(sheet);
  var english = { sep: ";", iff: "IF" };
  var russian = { sep: ";", iff: "ЕСЛИ" };
  var attempts = [];
  function push(dialect, modern) {
    if (!dialect) return;
    attempts.push({ dialect: dialect, modern: modern || "" });
  }
  push(detected, "");
  if (detected && detected.iff === "ЕСЛИ") push(detected, "ru");
  push(english, "");
  push(russian, "");
  push(russian, "ru");
  return attempts;
}

function reportFormulaFailed(shown) {
  var text = String(shown || "");
  if (!text) return true;
  var bad = ["#", "NAME", "ИМЯ", "ERROR", "ОШИБКА", "N/A", "Н/Д", "VALUE", "ЗНАЧ", "Загрузка", "Loading"];
  for (var i = 0; i < bad.length; i++) if (text.indexOf(bad[i]) >= 0) return true;
  return false;
}

function placeReportCharts() {
  var book = registryBook();
  var sh = book.getSheetByName("Графики");
  if (!sh) sh = book.insertSheet("Графики");
  var old = sh.getCharts ? sh.getCharts() : [];
  for (var c = old.length - 1; c >= 0; c--) sh.removeChart(old[c]);
  sh.clearContents();
  if (sh.setTabColor) sh.setTabColor("#188038");
  var note = "Цифры как на пульте. Прибыль = поступление минус расход рекламы, без себестоимости. Списание показано отдельно и из прибыли не вычитается. ОБРАЗЕЦ входит всегда. Остальные строки входят, когда франшиза АКТИВЕН и источник Включать. Строки МойСклад входят при статусе АКТИВЕН. Расход рекламы 0: выгрузка не сохранена. Это не факт оплаты.";
  sh.getRange(1, 1).setValue("Графики");
  sh.getRange(2, 1).setValue(note);
  sh.getRange(4, 1, 1, 2).setValues([["Показатель", "Сумма"]]);
  sh.getRange(5, 1, 5, 1).setValues([["Строк в своде"], ["Поступления"], ["Списания"], ["Расход Google Ads"], ["Прибыль до себестоимости"]]);
  sh.getRange(8, 2).setValue(0);
  sh.getRange(11, 1).setValue("Столбцы графика");
  sh.getRange(12, 1, 3, 1).setValues([["Поступления"], ["Списания"], ["Прибыль до себестоимости"]]);
  sh.getRange(16, 1).setValue("Строки, которые входят в суммы");
  sh.getRange(17, 1, 1, 4).setValues([["Название", "Поступление", "Списание", "Дата"]]);
  sh.getRange(1, 11).setValue("Служебный признак. 1 — строка входит в суммы. Робот этот столбец не читает.");
  var attempts = reportDialectAttempts(sh);
  var used = null;
  for (var a = 0; a < attempts.length; a++) {
    var formulas = reportChartFormulas(attempts[a].dialect, attempts[a].modern);
    sh.getRange(2, 11).setFormula(formulas.flag);
    sh.getRange(5, 2).setFormula(formulas.rows);
    sh.getRange(6, 2).setFormula(formulas.moneyIn);
    sh.getRange(7, 2).setFormula(formulas.moneyOut);
    sh.getRange(9, 2).setFormula(formulas.profit);
    sh.getRange(12, 2).setFormula(formulas.inBar);
    sh.getRange(13, 2).setFormula(formulas.outBar);
    sh.getRange(14, 2).setFormula(formulas.profitBar);
    sh.getRange(18, 1).setFormula(formulas.detailName);
    sh.getRange(18, 2).setFormula(formulas.detailIn);
    sh.getRange(18, 3).setFormula(formulas.detailOut);
    sh.getRange(18, 4).setFormula(formulas.detailDate);
    if (typeof SpreadsheetApp !== "undefined" && SpreadsheetApp.flush) SpreadsheetApp.flush();
    var shown = "";
    var detailShown = "";
    try { shown = sh.getRange(6, 2).getDisplayValue(); } catch (err) { shown = ""; }
    try { detailShown = sh.getRange(18, 1).getDisplayValue(); } catch (err2) { detailShown = ""; }
    if (!reportFormulaFailed(shown) && !reportFormulaFailed(detailShown)) {
      used = formulas;
      break;
    }
  }
  if (!used) throw new Error("формула графика не посчиталась");
  if (sh.getRange(6, 2).setNumberFormat) {
    sh.getRange(6, 2, 4, 1).setNumberFormat("#,##0.00");
    sh.getRange(12, 2, 3, 1).setNumberFormat("#,##0.00");
    sh.getRange(18, 2, 20, 2).setNumberFormat("#,##0.00");
    sh.getRange(18, 4, 20, 1).setNumberFormat("yyyy-mm-dd");
  }
  if (sh.getRange(1, 1).setBackground) {
    sh.getRange(1, 1).setBackground("#188038").setFontColor("#FFFFFF").setFontWeight("bold");
    sh.getRange(2, 1).setBackground("#FFF4CC").setWrap(true);
    sh.getRange(4, 1, 1, 2).setBackground("#202124").setFontColor("#FFFFFF").setFontWeight("bold");
    sh.getRange(8, 2).setNote("Кабинет 868-944-6535. Валюта USD. Выгрузка в BigQuery не сохранена.");
  }
  if (sh.setColumnWidth) {
    sh.setColumnWidth(1, 280);
    sh.setColumnWidth(2, 180);
    sh.setColumnWidth(11, 360);
  }
  if (sh.setRowHeight) sh.setRowHeight(2, 72);
  if (sh.newChart && typeof Charts !== "undefined") {
    var totals = sh.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sh.getRange(12, 1, 3, 2))
      .setPosition(4, 6, 0, 0)
      .setOption("title", "Поступление, списание и прибыль")
      .setOption("legend", { position: "none" })
      .setOption("height", 300)
      .setOption("width", 520)
      .build();
    var lines = sh.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sh.getRange(17, 1, 14, 3))
      .setNumHeaders(1)
      .setPosition(16, 6, 0, 0)
      .setOption("title", "Строки свода")
      .setOption("height", 320)
      .setOption("width", 520)
      .build();
    sh.insertChart(totals);
    sh.insertChart(lines);
  }
  var vitrina = book.getSheetByName("Витрина");
  if (vitrina && vitrina.getIndex && book.moveActiveSheet) {
    book.setActiveSheet(sh);
    book.moveActiveSheet(vitrina.getIndex() + 1);
  }
  var memo = book.getSheetByName("Памятка");
  if (memo) {
    var memoRows = memo.getDataRange().getValues();
    var marked = false;
    for (var m = 0; m < memoRows.length; m++) if (String(memoRows[m][0]) === "Графики") marked = true;
    if (!marked) memo.appendRow(["Графики", "Столбцы как на пульте. Студия данных в этой стране не открывается.", "Прибыль = поступление минус расход рекламы. Списание из прибыли не вычитается.", "ОБРАЗЕЦ входит. Реклама 0. Это не факт оплаты.", "", ""]);
  }
  var service = book.getSheetByName("Служебное");
  if (service) {
    var labels = service.getDataRange().getValues();
    for (var r = 12; r < labels.length; r++) {
      if (String(labels[r][0]) === "Студия данных" && String(labels[r][1] || "").indexOf("datastudio.google.com") >= 0) {
        service.getRange(r + 1, 3).setValue("Ссылка на отчёт. В этой стране сервис не открывается. Рабочие графики — лист «Графики». Робот ссылку не читает.");
      }
    }
  }
  logLine(book, "FR-001", "Графики", "заметка", "Лист «Графики»: поступление, списание и прибыль до себестоимости. Это не факт оплаты.");
  var totalsNow = reportTotals(book);
  return { rows: totalsNow.rows, moneyIn: totalsNow.moneyIn, moneyOut: totalsNow.moneyOut, profit: totalsNow.profit, charts: sh.getCharts ? sh.getCharts().length : 0 };
}
