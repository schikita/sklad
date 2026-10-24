var fs = require("fs");
var path = require("path");
var vm = require("vm");

var failed = 0;
var passed = 0;

function assert(cond, message) {
  if (!cond) {
    failed++;
    console.error("FAIL " + message);
    return;
  }
  passed++;
}

function eq(actual, expected, message) {
  var ok = actual === expected;
  if (!ok && typeof actual === "object" && typeof expected === "object") ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (!ok) console.error("  actual=" + JSON.stringify(actual) + " expected=" + JSON.stringify(expected));
  assert(ok, message);
}

function Sheet(name, values) {
  this.name = name;
  this.values = values || [];
  this.frozen = 0;
  this.notes = {};
  this.formulas = {};
  this.charts = [];
}
Sheet.prototype.getName = function () { return this.name; };
Sheet.prototype.getLastRow = function () { return this.values.length; };
Sheet.prototype.getLastColumn = function () {
  var w = 1;
  this.values.forEach(function (row) { if (row.length > w) w = row.length; });
  return w;
};
Sheet.prototype.ensure = function (r, c) {
  while (this.values.length < r) this.values.push([]);
  while (this.values[r - 1].length < c) this.values[r - 1].push("");
};
Sheet.prototype.getDataRange = function () {
  var self = this;
  var rows = Math.max(self.values.length, 1);
  var cols = self.getLastColumn();
  return self.getRange(1, 1, rows, cols);
};
Sheet.prototype.clearContents = function () { this.values = []; this.formulas = {}; };
Sheet.prototype.setTabColor = function () {};
Sheet.prototype.setColumnWidth = function () {};
Sheet.prototype.getCharts = function () { return this.charts.slice(); };
Sheet.prototype.insertChart = function (chart) { this.charts.push(chart); };
Sheet.prototype.removeChart = function (chart) {
  this.charts = this.charts.filter(function (item) { return item !== chart; });
};
Sheet.prototype.getIndex = function () { return 4; };
Sheet.prototype.newChart = function () {
  var spec = { type: "", ranges: [], options: {}, position: null, headers: 0 };
  var chain = {
    setChartType: function (type) { spec.type = type; return chain; },
    addRange: function (range) { spec.ranges.push(range); return chain; },
    setPosition: function (row, col, x, y) { spec.position = [row, col, x, y]; return chain; },
    setOption: function (key, value) { spec.options[key] = value; return chain; },
    setNumHeaders: function (count) { spec.headers = count; return chain; },
    build: function () { return spec; }
  };
  return chain;
};
Sheet.prototype.appendRow = function (row) { this.values.push(row.slice()); };
Sheet.prototype.setFrozenRows = function (n) { this.frozen = n; };
Sheet.prototype.setRowHeight = function () {};
Sheet.prototype.insertRowBefore = function (at) {
  this.values.splice(at - 1, 0, []);
};
Sheet.prototype.deleteRow = function (at) { this.values.splice(at - 1, 1); };
Sheet.prototype.getRange = function (a, b, c, d) {
  var r1, c1, numR, numC;
  if (typeof a === "string") {
    var m = String(a).match(/^([A-Z]+)(\d+)$/);
    var col = 0;
    for (var i = 0; i < m[1].length; i++) col = col * 26 + (m[1].charCodeAt(i) - 64);
    r1 = Number(m[2]);
    c1 = col;
    numR = 1;
    numC = 1;
  } else if (c === undefined) {
    r1 = a; c1 = b; numR = 1; numC = 1;
  } else {
    r1 = a; c1 = b; numR = c; numC = d;
  }
  var self = this;
  function cell(rr, cc) {
    var row = self.values[rr - 1] || [];
    return row[cc - 1] === undefined ? "" : row[cc - 1];
  }
  var range = {
    getValue: function () { return cell(r1, c1); },
    setValue: function (v) { self.ensure(r1, c1); self.values[r1 - 1][c1 - 1] = v; return range; },
    getValues: function () {
      var out = [];
      for (var i = 0; i < numR; i++) {
        var line = [];
        for (var j = 0; j < numC; j++) line.push(cell(r1 + i, c1 + j));
        out.push(line);
      }
      return out;
    },
    setValues: function (grid) {
      for (var i = 0; i < grid.length; i++) {
        for (var j = 0; j < grid[i].length; j++) {
          self.ensure(r1 + i, c1 + j);
          self.values[r1 - 1 + i][c1 - 1 + j] = grid[i][j];
        }
      }
      return range;
    },
    getFormulas: function () {
      var out = [];
      for (var i = 0; i < numR; i++) {
        var line = [];
        for (var j = 0; j < numC; j++) line.push("");
        out.push(line);
      }
      return out;
    },
    setBackground: function () { return range; },
    setFontColor: function () { return range; },
    setFontWeight: function () { return range; },
    setVerticalAlignment: function () { return range; },
    setWrap: function () { return range; },
    getNote: function () { return self.notes[r1 + ":" + c1] || ""; },
    setNote: function (text) { self.notes[r1 + ":" + c1] = text; return range; },
    setFormula: function (formula) { self.formulas[r1 + ":" + c1] = formula; return range; },
    getFormula: function () { return self.formulas[r1 + ":" + c1] || ""; },
    getDisplayValue: function () {
      var formula = self.formulas[r1 + ":" + c1] || "";
      if (formula === "=IF(TRUE,1,0)" || formula === "=IF(TRUE;1;0)" || formula === "=ЕСЛИ(ИСТИНА;1;0)") return "1";
      if (formula.indexOf("REGEXMATCH") >= 0 || formula.indexOf("РЕГВЫРАЖ") >= 0) return "TRUE";
      if (!formula) return String(cell(r1, c1));
      return "15000";
    },
    setNumberFormat: function () { return range; },
    setFontSize: function () { return range; },
    clearContent: function () { return range; },
    clearDataValidations: function () { return range; },
    setDataValidation: function () { return range; }
  };
  return range;
};

function Book(id) {
  this.id = id;
  this.sheets = {};
  this.moved = [];
  this.active = "";
}
Book.prototype.getId = function () { return this.id; };
Book.prototype.getUrl = function () { return "https://docs.google.com/spreadsheets/d/" + this.id + "/edit"; };
Book.prototype.getSheetByName = function (name) { return this.sheets[name] || null; };
Book.prototype.insertSheet = function (name) {
  var sh = new Sheet(name, []);
  this.sheets[name] = sh;
  return sh;
};
Book.prototype.getSheets = function () {
  var out = [];
  for (var name in this.sheets) out.push(this.sheets[name]);
  return out;
};
Book.prototype.setActiveSheet = function (sh) { this.active = sh.getName(); };
Book.prototype.moveActiveSheet = function (pos) { this.moved.push(this.active + "@" + pos); };
Book.prototype.deleteSheet = function (sh) { delete this.sheets[sh.getName()]; };

function blankBook(id) {
  var book = new Book(id || "REGISTRYID1234567890abcd");
  book.sheets["Управление Франшизами"] = new Sheet("Управление Франшизами", [
    ["ID", "Название", "Токен", "Статус", "Орг", "UUID", "Комментарий", ""],
    ["", "справка", "", "", "", "", "", "СПРАВКА"],
    ["FR-001", "ООО РемМонтаж", "abc12345tokenvalue", "ОТКЛЮЧЕН", "ooorusuglneftgaz", "028234d9-cba0-11f1-0a80-0011004961d9", "", ""]
  ]);
  book.sheets["Источники"] = new Sheet("Источники", [
    ["ID", "Имя", "Таблица", "Лист", "Статус", "Бланк", "Ссылка", "Шапка", "Метка"],
    ["FR-001", "Франшиза FR-001", "", "", "", "", "", "", "ГРУППА"],
    ["FR-001", "Контрагенты", "sheet-counterparty-0001", "Контрагенты", "Включать", "Контрагенты", "", "", ""]
  ]);
  book.sheets["Словарь полей"] = new Sheet("Словарь полей", [
    ["Заголовок", "Поле", "Тип"],
    ["", "справка", "", ""],
    ["Наименование", "name", "text"],
    ["Дата", "report_date", "date"],
    ["Поступление", "money_in", "money"],
    ["Списание", "money_out", "money"],
    ["Ссылка", "link", "link"]
  ]);
  book.sheets["Бланки"] = new Sheet("Бланки", [
    ["Тип", "Столбец"],
    ["СПРАВКА", "не бланк"],
    ["Контрагенты", "Контрагенты"],
    ["Товары", "Наименование"],
    ["Кадры", ""]
  ]);
  book.sheets["Журнал"] = new Sheet("Журнал", [["Время", "Франшиза", "Источник", "Уровень", "Сообщение"]]);
  book.sheets["Витрина"] = new Sheet("Витрина", [["код", "источник", "лист", "название", "дата", "приход", "расход", "кампания", "ссылка", "след"]]);
  book.sheets["Сырьё"] = new Sheet("Сырьё", [["код", "источник", "лист", "когда", "тип", "след"]]);
  book.sheets["Правки"] = new Sheet("Правки", [["Ключ", "Поле", "Значение", "Когда"], ["", "справка", "", "", "СПРАВКА"]]);
  book.sheets["Вход 1С"] = new Sheet("Вход 1С", [["Время", "JSON"]]);
  book.sheets["Служебное"] = new Sheet("Служебное", [
    ["Робот", "alexsklad083@gmail.com"],
    ["Доступ", "читатель"],
    ["Проект BigQuery", "remmontazh-eu"],
    ["Регион", "EU"],
    ["Customer ID Google Ads", "868-944-6535"]
  ]);
  book.sheets["Памятка"] = new Sheet("Памятка", [["Памятка"]]);
  return book;
}

var logs = [];
var props = {};
var books = {};
var fetches = [];
var driveFiles = [];
var folders = {};
var nowQueue = [];

function resetWorld() {
  logs = [];
  props = { REGISTRY_ID: "REGISTRYID1234567890abcd" };
  books = {};
  fetches = [];
  driveFiles = [];
  folders = {};
  nowQueue = [];
  books[props.REGISTRY_ID] = blankBook(props.REGISTRY_ID);
}

var sandbox = {
  console: console,
  Date: Date,
  JSON: JSON,
  Math: Math,
  Number: Number,
  String: String,
  Object: Object,
  Array: Array,
  RegExp: RegExp,
  Error: Error,
  isNaN: isNaN,
  Logger: { log: function (m) { logs.push(String(m)); } },
  Utilities: {
    formatDate: function (date, tz, fmt) {
      var d = new Date(date.getTime() + 3 * 3600 * 1000);
      function p(n) { return (n < 10 ? "0" : "") + n; }
      var y = d.getUTCFullYear();
      var m = p(d.getUTCMonth() + 1);
      var day = p(d.getUTCDate());
      var hh = p(d.getUTCHours());
      var mm = p(d.getUTCMinutes());
      var ss = p(d.getUTCSeconds());
      if (fmt === "yyyy-MM-dd") return y + "-" + m + "-" + day;
      if (fmt === "yyyy-MM-dd HH:mm:ss") return y + "-" + m + "-" + day + " " + hh + ":" + mm + ":" + ss;
      if (fmt === "yyyy-MM-dd HH-mm") return y + "-" + m + "-" + day + " " + hh + "-" + mm;
      if (fmt === "dd.MM HH:mm") return day + "." + m + " " + hh + ":" + mm;
      return String(date);
    },
    sleep: function () {}
  },
  PropertiesService: {
    getScriptProperties: function () {
      return {
        getProperty: function (k) { return props[k] === undefined ? null : props[k]; },
        setProperty: function (k, v) { props[k] = String(v); },
        deleteProperty: function (k) { delete props[k]; }
      };
    }
  },
  SpreadsheetApp: {
    openById: function (id) {
      if (!books[id]) throw new Error("нет книги " + id);
      return books[id];
    },
    getActive: function () { return books[props.REGISTRY_ID]; },
    newDataValidation: function () {
      var chain = {
        requireValueInList: function () { return chain; },
        setAllowInvalid: function () { return chain; },
        build: function () { return { rule: true }; }
      };
      return chain;
    },
    flush: function () {}
  },
  Charts: { ChartType: { COLUMN: "COLUMN", BAR: "BAR" } },
  UrlFetchApp: {
    fetch: function (url) {
      var next = fetches.shift();
      if (!next) throw new Error("нет ответа на " + url);
      next.url = url;
      return {
        getResponseCode: function () { return next.code; },
        getContentText: function () { return next.text; }
      };
    }
  },
  HtmlService: {
    createTemplateFromFile: function () {
      return {
        franchise: "",
        evaluate: function () {
          return {
            setTitle: function () { return this; },
            setXFrameOptionsMode: function () { return { title: "Пульт РемМонтаж" }; }
          };
        }
      };
    },
    createHtmlOutputFromFile: function () { return { getContent: function () { return ""; } }; },
    XFrameOptionsMode: { ALLOWALL: "ALLOWALL" }
  },
  ContentService: {
    createTextOutput: function (text) {
      return {
        body: text,
        setMimeType: function () { return this; }
      };
    },
    MimeType: { JSON: "json" }
  },
  ScriptApp: {
    getProjectTriggers: function () { return []; },
    deleteTrigger: function () {},
    newTrigger: function () {
      var chain = {
        timeBased: function () { return chain; },
        everyDays: function () { return chain; },
        atHour: function () { return chain; },
        after: function () { return chain; },
        create: function () { return { created: true }; }
      };
      return chain;
    }
  },
  DriveApp: {
    getFoldersByName: function (name) {
      var list = folders[name] ? [folders[name]] : [];
      var i = 0;
      return { hasNext: function () { return i < list.length; }, next: function () { return list[i++]; } };
    },
    createFolder: function (name) {
      var folder = makeFolder(name);
      folders[name] = folder;
      return folder;
    },
    getFileById: function (id) {
      return {
        makeCopy: function (title, folder) {
          var file = { name: title, id: "copy-" + driveFiles.length, created: new Date(), trashed: false, url: "https://docs.google.com/spreadsheets/d/copy/edit" };
          file.getName = function () { return file.name; };
          file.getDateCreated = function () { return file.created; };
          file.setTrashed = function (v) { file.trashed = v; };
          file.getUrl = function () { return file.url; };
          driveFiles.push(file);
          folder.files.push(file);
          return file;
        }
      };
    }
  }
};

function makeFolder(name) {
  return {
    name: name,
    files: [],
    getFiles: function () {
      var i = 0;
      var list = this.files;
      return { hasNext: function () { return i < list.length; }, next: function () { return list[i++]; } };
    }
  };
}

sandbox.Date.now = function () {
  if (nowQueue.length) return nowQueue.shift();
  return Date.parse("2026-10-22T20:00:00Z");
};

vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(__dirname, "app_script_logic.js"), "utf8"), sandbox, { filename: "app_script_logic.js" });
vm.runInContext(fs.readFileSync(path.join(__dirname, "lay_operator_sheets.gs"), "utf8"), sandbox, { filename: "lay_operator_sheets.gs" });

function g(name) { return sandbox[name]; }

resetWorld();

eq(g("asNumber")(15000), 15000, "число остаётся числом");
eq(g("asNumber")("15 000,50"), 15000.5, "пробел и запятая");
eq(g("asNumber")("15\u00a0000"), 15000, "неразрывный пробел");
eq(g("asNumber")("2 500 руб."), 2500, "хвост валюты отбрасывается");
eq(g("asNumber")(""), "", "пустое не ноль");
eq(g("asNumber")("нет"), "", "текст не становится нулём");
eq(g("asNumber")("12.3.4"), "", "два разделителя не число");

eq(g("asDate")("01.10.2026"), "2026-10-01", "дата дд.мм.гггг");
eq(g("asDate")("1.2.2026 г."), "2026-02-01", "короткая дата и буква г");
eq(g("asDate")("2026-10-02T18:00:00"), "2026-10-02", "дата из ISO обрезается");
eq(g("asDate")("вчера"), "", "мусор не дата");
eq(g("asDate")(new Date(Date.UTC(2026, 9, 22, 0, 0, 0))), "2026-10-22", "объект даты в Москве");

var unknown = {};
var bad = {};
var mapped = g("mapPayload")({
  "Наименование": "ТЕСТ. Оплата",
  "Поступление": "15 000",
  "Дата": "02.10.2026",
  "Ссылка": "https://example.test/card",
  "Лишний": "мимо",
  "Списание": "текст"
}, {
  "наименование": { field: "name", type: "text" },
  "поступление": { field: "money_in", type: "money" },
  "списание": { field: "money_out", type: "money" },
  "дата": { field: "report_date", type: "date" },
  "ссылка": { field: "link", type: "link" }
}, unknown, bad);
eq(mapped.name, "ТЕСТ. Оплата", "имя из словаря");
eq(mapped.money_in, 15000, "деньги из текста");
eq(mapped.report_date, "2026-10-02", "дата из словаря");
eq(mapped.link, "https://example.test/card", "ссылка");
eq(mapped.money_out, "", "текст вместо суммы очищен");
assert(unknown["Лишний"] === true, "незнакомый столбец отмечен");
assert(bad["Списание"] === true, "плохая сумма отмечена");

eq(g("pickName")({ name: "Из словаря" }, { a: "длинное поле" }), "Из словаря", "имя берётся из словаря");
eq(g("pickName")({}, { url: "https://example.test/x", title: "ТЕСТ" }), "ТЕСТ", "ссылка не становится названием");
var split = g("splitCell")('=HYPERLINK("https://example.test/a";"карточка")', "карточка");
eq(split.url, "https://example.test/a", "адрес из формулы");
eq(split.text, "карточка", "текст ячейки сохранён");

eq(g("findHeader")([["", ""], ["Дата", "Сумма", "Имя"], ["1", "2"]]), 1, "шапка — самая заполненная из первых строк");
eq(g("pad")(["a"], 3).length, 3, "строка добивается до ширины листа");
eq(g("sheetIdFromUrl")("https://docs.google.com/spreadsheets/d/abc_DEF-123456789012345/edit"), "abc_DEF-123456789012345", "id из ссылки");
eq(g("sheetIdFromUrl")("abc_DEF-123456789012345678"), "abc_DEF-123456789012345678", "голый id");
var thrown = false;
try { g("sheetIdFromUrl")("не ссылка"); } catch (err) { thrown = true; }
assert(thrown, "мусор вместо ссылки отвергается");

var book = books[props.REGISTRY_ID];
var franchises = g("readFranchiseList")(book);
eq(franchises.length, 1, "строка-справка без кода не читается");
eq(franchises[0].id, "FR-001", "код франшизы");
eq(franchises[0].sheetRow, 3, "рабочая строка после справки");
eq(franchises[0].status, "ОТКЛЮЧЕН", "статус читается как есть");
assert(franchises[0].token === "abc12345tokenvalue", "ключ остаётся в строке франшизы");

var sources = g("readSourceList")(book);
eq(sources.length, 1, "группа без id таблицы не читается");
eq(sources[0].name, "Контрагенты", "имя источника");
eq(g("requiredHeader")(book, "Контрагенты"), "Контрагенты", "бланк возвращает обязательный столбец");
eq(g("requiredHeader")(book, "Кадры"), "", "пустой бланк ничего не требует");
eq(g("requiredHeader")(book, "СПРАВКА"), "не бланк", "метка справки не совпадает с типом источника случайно, если тип так назван");
eq(g("requiredHeader")(book, ""), "", "пустой тип бланка не проверяется");
var dict = g("readDictionaryRows")(book);
assert(dict[0].header === "Наименование", "пустая строка словаря пропускается");

book.sheets["Правки"].appendRow(["k|1", "money_in", 10, ""]);
book.sheets["Правки"].appendRow(["k|1", "money_in", 15, ""]);
book.sheets["Правки"].appendRow(["k|1", "money_out", 3, ""]);
var fixes = g("correctionMap")(book);
eq(fixes["k|1"].money_in, 15, "повтор ключа затирает старую правку");
eq(fixes["k|1"].money_out, 3, "второе поле того же ключа живёт рядом");

book.sheets["Витрина"].appendRow(["FR-001", "Контрагенты", "Контрагенты", "Живой", "2026-10-01", 100, 0, "", "", ""]);
book.sheets["Витрина"].appendRow(["FR-001", "МойСклад", "paymentin", "Платёж", "2026-10-01", 50, "", "", "", ""]);
book.sheets["Витрина"].appendRow(["FR-001", "ОБРАЗЕЦ", "Образец", "ТЕСТ", "2026-10-01", 15000, 2600, "", "", ""]);
var off = g("summarize")(book);
eq(off.rows, 1, "при выключенной франшизе на пульте остаётся только образец");
eq(off.moneyIn, 15000, "поступление образца видно при выключенной франшизе");
book.sheets["Управление Франшизами"].values[2][3] = "АКТИВЕН";
var on = g("summarize")(book);
eq(on.rows, 3, "активная франшиза считает источник, API и образец");
eq(on.moneyIn, 15150, "суммы складываются");
eq(on.moneyOut, 2600, "списание образца входит в свод");
var only = g("summarize")(book, { name: "ОБРАЗЕЦ", sheet: "Образец", virtual: false });
eq(only.rows, 1, "свод одного источника не тащит чужие строки");
var virt = g("summarize")(book, { virtual: true, name: "МойСклад", sheet: "API" });
eq(virt.rows, 1, "виртуальный источник считает только МойСклад");
eq(virt.moneyIn, 50, "платёж API");
book.sheets["Управление Франшизами"].values[2][3] = "ОТКЛЮЧЕН";

var state = g("apiState")();
assert(JSON.stringify(state).indexOf("abc12345tokenvalue") === -1, "пульт не отдаёт сам ключ");
eq(state.sources.length, 0, "выключенная франшиза не попадает в список источников пульта");

var names = g("franchiseNames")(book);
eq(names["FR-001"], "ООО РемМонтаж", "имя франшизы для подписи группы");
assert(!names[""], "справка не становится франшизой");

g("replacePairs")(book.sheets["Витрина"], 10, "ОБРАЗЕЦ", "Образец", [
  ["FR-001", "ОБРАЗЕЦ", "Образец", "ТЕСТ. Новая", "2026-10-03", 1, "", "", "", "{}"]
]);
var after = book.sheets["Витрина"].values;
eq(after[0][0], "код", "заголовок витрины не затирается");
assert(after.filter(function (row) { return row[1] === "ОБРАЗЕЦ"; }).length === 1, "старые строки этого источника заменены");
assert(after.some(function (row) { return row[1] === "Контрагенты"; }), "чужой источник остаётся");

var external = new Book("sheet-counterparty-0001");
external.sheets["Контрагенты"] = new Sheet("Контрагенты", [
  ["Контрагенты", "Пометка"],
  ["ТЕСТ. Контрагент", "лишнее"]
]);
books[external.id] = external;
g("syncSheet")(book, franchises[0], sources[0]);
var synced = book.sheets["Витрина"].values.filter(function (row) { return row[1] === "Контрагенты" && row[3] === "ТЕСТ. Контрагент"; });
eq(synced.length, 1, "строка источника легла в витрину");
eq(synced[0][3], "ТЕСТ. Контрагент", "название взято из столбца словаря");
assert(book.sheets["Журнал"].values.some(function (row) { return String(row[4]).indexOf("Незнакомые столбцы") >= 0; }), "лишний столбец записан предупреждением");
assert(book.sheets["Сырьё"].values.some(function (row) { return row[1] === "Контрагенты"; }), "сырьё хранит след той же строки");

external.sheets["Нет такого"] = null;
g("syncSheet")(book, { id: "FR-001", status: "АКТИВЕН" }, { id: external.id, name: "Пусто", sheet: "НетЛиста", franchise: "FR-001", blank: "", headerRow: 0 });
assert(book.sheets["Журнал"].values.some(function (row) { return String(row[4]).indexOf("не найден") >= 0; }), "нет листа — предупреждение");

g("syncSheet")(book, { id: "FR-001" }, { id: external.id, name: "Товары", sheet: "Контрагенты", blank: "Товары", headerRow: 1, franchise: "FR-001" });
assert(book.sheets["Журнал"].values.some(function (row) { return String(row[4]).indexOf("Нет столбца") >= 0; }), "нет обязательного столбца — лист пропущен");

var itemIn = { name: "", agent: { name: "ТЕСТ. Плательщик" }, moment: "2026-10-01 12:00:00", sum: 1500000, meta: { uuidHref: "https://online.moysklad.ru/app/#paymentin/edit?id=1" } };
var itemOut = { name: "ТЕСТ. Выплата", moment: "2026-10-02 12:00:00", sum: 250000, meta: {} };
var huge = { id: "1", name: "ТЕСТ. Длинный", sum: 100 };
huge.blob = new Array(46000).join("x");
g("writeMoySklad")(book, { id: "FR-001" }, "paymentin", [itemIn], true);
g("writeMoySklad")(book, { id: "FR-001" }, "paymentout", [itemOut], true);
g("writeMoySklad")(book, { id: "FR-001" }, "product", [{ name: "ТЕСТ. Товар", sum: 99900 }], true);
g("writeMoySklad")(book, { id: "FR-001" }, "counterparty", [huge], true);
var vin = book.sheets["Витрина"].values.filter(function (row) { return row[2] === "paymentin"; })[0];
var vout = book.sheets["Витрина"].values.filter(function (row) { return row[2] === "paymentout"; })[0];
var vprod = book.sheets["Витрина"].values.filter(function (row) { return row[2] === "product"; })[0];
var vhuge = book.sheets["Сырьё"].values.filter(function (row) { return row[2] === "counterparty"; })[0];
eq(vin[3], "ТЕСТ. Плательщик", "имя платежа берётся у контрагента");
eq(vin[5], 15000, "копейки API делятся на 100 и идут в поступление");
eq(vin[6], "", "входящий платёж не пишет списание");
eq(vout[6], 2500, "исходящий платёж — списание");
eq(vout[5], "", "исходящий платёж не пишет поступление");
eq(vprod[5], "", "товар не раздувает поступление");
assert(String(vhuge[5]).indexOf("truncated") >= 0, "длинный JSON обрезается и помечается");

g("pullMoySklad")(book, { id: "FR-001", token: "short", sheetRow: 3 });
assert(book.sheets["Журнал"].values.some(function (row) { return String(row[4]).indexOf("Токен не похож") >= 0; }), "короткий ключ не уходит в API");
g("pullMoySklad")(book, { id: "FR-001", token: "", sheetRow: 3 });
assert(book.sheets["Журнал"].values.some(function (row) { return String(row[4]).indexOf("Токен не задан") >= 0; }), "пустой ключ пропускает API");
fetches.push({ code: 401, text: "" });
g("pullMoySklad")(book, { id: "FR-001", token: "abcdefghij", sheetRow: 3 });
eq(book.sheets["Управление Франшизами"].values[2][3], "ОШИБКА ТОКЕНА", "401 ставит статус ошибки ключа");
fetches.push({ code: 200, text: JSON.stringify({ rows: [{ name: "Фирма", id: "org-1" }] }) });
fetches.push({ code: 429, text: "" });
fetches.push({ code: 500, text: "" });
for (var left = 0; left < 6; left++) fetches.push({ code: 200, text: JSON.stringify({ rows: [] }) });
g("pullMoySklad")(book, { id: "FR-009", token: "abcdefghij", sheetRow: 3 });
assert(book.sheets["Журнал"].values.some(function (row) { return String(row[4]).indexOf("Ответ 500") >= 0; }), "повтор после 429 и чужой код пишутся ошибкой");
fetches.push({ code: 200, text: JSON.stringify({ rows: [{ name: "Фирма", id: "org-1" }] }) });
["counterparty", "product", "paymentin", "paymentout", "customerorder", "demand", "employee"].forEach(function () {
  fetches.push({ code: 200, text: JSON.stringify({ rows: [] }) });
});
book.sheets["Управление Франшизами"].values[2][3] = "АКТИВЕН";
g("pullMoySklad")(book, { id: "FR-001", token: "abcdefghij", sheetRow: 3 });
eq(book.sheets["Управление Франшизами"].values[2][4], "Фирма", "принятый ключ записывает имя фирмы");
eq(book.sheets["Управление Франшизами"].values[2][5], "org-1", "принятый ключ записывает UUID");

book.sheets["Служебное"].values[2][1] = "";
book.sheets["Служебное"].values[4][1] = "";
g("noteWarehouse")(book);
assert(book.sheets["Журнал"].values.some(function (row) { return String(row[4]).indexOf("Проект Google Cloud не задан") >= 0; }), "пустой B3 — заметка");
assert(book.sheets["Журнал"].values.some(function (row) { return String(row[4]).indexOf("Customer ID нет") >= 0; }), "пустой B5 — заметка");
book.sheets["Служебное"].values[2][1] = "remmontazh-eu";
book.sheets["Служебное"].values[4][1] = "868-944-6535";

var saved = g("apiSaveCorrection")({ key: "FR-001|x|Лист|ТЕСТ|2026-10-01", field: "money_in", value: "1 500", vitrinaRow: 2 });
eq(saved.ok, true, "правка сохраняется");
var again = g("apiSaveCorrection")({ key: "FR-001|x|Лист|ТЕСТ|2026-10-01", field: "money_in", value: "1600", vitrinaRow: 2 });
eq(again.ok, true, "повторная правка не плодит строку");
var moneyRows = book.sheets["Правки"].values.filter(function (row) { return row[0] === "FR-001|x|Лист|ТЕСТ|2026-10-01"; });
eq(moneyRows.length, 1, "тот же ключ и поле остаются одной строкой");
eq(moneyRows[0][2], 1600, "в правке лежит число");
eq(book.sheets["Витрина"].values[1][5], 1600, "цифра сразу пишется в витрину");

var headerStatus = book.sheets["Источники"].values[0][4];
g("apiSetSourceStatus")({ sheetRow: 1, status: "Исключить" });
eq(book.sheets["Источники"].values[0][4], headerStatus, "строка шапки не меняется");
g("apiSetSourceStatus")({ sheetRow: 3, status: "Исключить" });
eq(book.sheets["Источники"].values[2][4], "Исключить", "статус источника меняется");
g("apiSetSourceStatus")({ sheetRow: 3, status: "ЧУЖОЙ" });
eq(book.sheets["Источники"].values[2][4], "Исключить", "чужой статус не записывается");
g("apiSetFranchiseStatus")({ franchise: "FR-001", status: "ЗАБЛОКИРОВАН" });
eq(book.sheets["Управление Франшизами"].values[2][3], "ЗАБЛОКИРОВАН", "статус франшизы меняется");
g("apiSetFranchiseStatus")({ franchise: "FR-001", status: "АКТИВЕН" });
var added = g("apiAddSource")({ url: "https://docs.google.com/spreadsheets/d/newSheetId_000000000001/edit", name: "ТЕСТ лист", sheet: "Лист" });
assert(added.sources.some(function (src) { return src.name === "ТЕСТ лист"; }), "новый источник появляется");
var beforeMaps = g("readDictionaryRows")(book).length;
g("apiAddMap")({ header: "", field: "name" });
eq(g("readDictionaryRows")(book).length, beforeMaps, "пустой заголовок в словарь не пишется");
g("apiAddMap")({ header: "ТЕСТ столбец", field: "name", type: "text" });
eq(g("readDictionaryRows")(book).length, beforeMaps + 1, "новый заголовок пишется");

var rows = g("apiRows")({ sheetRow: 3 });
assert(rows.every(function (row) { return row.name !== undefined; }), "строки источника имеют название");
var post = g("doPost")({ postData: { contents: new Array(50001).join("a") } });
eq(JSON.parse(post.body).ok, true, "посылка 1С принята");
eq(String(book.sheets["Вход 1С"].values[1][1]).length, 45000, "тело длиннее 45000 обрезается");
var page = g("doGet")({ parameter: { franchise: "FR-001" } });
eq(page.title, "Пульт РемМонтаж", "пульт открывается");
var badPage = g("doGet")({ parameter: { franchise: "<script>" } });
eq(badPage.title, "Пульт РемМонтаж", "чужой код в адресе не ломает страницу");

eq(g("journalAdvice")("ошибка", "Ключ не принят"), g("journalAdvice")("ошибка", "Ключ не принят"), "совет по ключу стабилен");
assert(g("journalAdvice")("ошибка", "Ключ не принят").indexOf("ОШИБКА ТОКЕНА") >= 0, "отвергнутый ключ ведёт в столбец токена");
assert(g("journalAdvice")("ошибка", "Токен не похож на ключ API").indexOf("Токен") >= 0, "кривой ключ объяснён");
assert(g("journalAdvice")("заметка", "Токен не задан, API пропущен").indexOf("пуст") >= 0, "пустой ключ не называется обрывом");
assert(g("journalAdvice")("предупреждение", "Лист «Касса» не найден").indexOf("Имя листа") >= 0, "нет листа — править источники");
assert(g("journalAdvice")("предупреждение", "Нет столбца «Дата»").indexOf("Бланки") >= 0, "нет столбца — бланк");
assert(g("journalAdvice")("предупреждение", "Незнакомые столбцы пропущены").indexOf("Словарь") >= 0, "чужие столбцы — словарь");
assert(g("journalAdvice")("предупреждение", "Текст вместо числа очищен").indexOf("исходн") >= 0, "текст вместо числа правят в источнике");
assert(g("journalAdvice")("заметка", "Франшиза не активна").indexOf("не порча") >= 0, "выключенный статус не порча данных");
assert(g("journalAdvice")("ошибка", "Ответ 429").indexOf("частот") >= 0, "429 отдельно от прочих ответов");
assert(g("journalAdvice")("ошибка", "Ответ 500").indexOf("не 200") >= 0, "прочий ответ API");
assert(g("journalAdvice")("заметка", "Пауза на строке реестра 3").indexOf("минуту") >= 0, "пауза времени не порча");
assert(g("journalAdvice")("ок", "2 строк. Файл карточек ещё не передан").indexOf("пустой файл") >= 0, "пустые товары — не обрыв");
assert(g("journalAdvice")("ошибка", "совсем новое сообщение").indexOf("Красная строка") >= 0, "незнакомая ошибка всё равно красная инструкция");
eq(g("journalAdvice")("ок", ""), "", "пустой текст без совета");
eq(g("journalAdvice")("ок", "7 строк"), "", "обычный успех без инструкции");

eq(g("sampleTable")().length, 2, "образец — две строки");
assert(String(g("sampleTable")()[0][0]).indexOf("ТЕСТ") === 0, "первая строка помечена ТЕСТ");
assert(String(g("sampleTable")()[1][0]).indexOf("ТЕСТ") === 0, "вторая строка помечена ТЕСТ");
eq(g("sampleTable")()[0][2], 15000, "образец поступления");
eq(g("sampleTable")()[1][3], 2500, "образец списания в источнике");

var placed = g("placeCustomerSample")();
eq(placed.rows, 2, "в книгу легли две строки");
eq(placed.vitrinaWriteOff, 2600, "на витрине списание уже с правкой");
var sampleRows = book.sheets["Витрина"].values.filter(function (row) { return row[1] === "ОБРАЗЕЦ"; });
eq(sampleRows.length, 2, "на витрине ровно две строки образца");
eq(sampleRows[0][5], 15000, "поступление образца");
eq(sampleRows[1][6], 2600, "списание образца после правки");
eq(book.sheets["Образец"].values[2][3], 2500, "в источнике списание без правки");
g("placeCustomerSample")();
eq(book.sheets["Витрина"].values.filter(function (row) { return row[1] === "ОБРАЗЕЦ"; }).length, 2, "повтор не плодит строки витрины");
eq(book.sheets["Правки"].values.filter(function (row) { return row[1] === "money_out" && row[2] === 2600; }).length, 1, "правка образца одна");
assert(book.sheets["Памятка"].values.some(function (row) { return row[0] === "ОБРАЗЕЦ"; }), "памятка говорит, что это образец");
var sampleSummary = g("summarize")(book, { name: "ОБРАЗЕЦ", sheet: "Образец" });
eq(sampleSummary.moneyIn, 15000, "свод видит поступление образца");
eq(sampleSummary.moneyOut, 2600, "свод видит списание после правки, не 2500");

var old = [];
for (var n = 0; n < 16; n++) old.push({ name: g("backupTitlePrefix")() + "2026-10-" + (n < 10 ? "0" : "") + n, created: n, file: { trashed: false, setTrashed: function (v) { this.trashed = v; } } });
old.push({ name: "Чужая книга", created: 100, file: { trashed: false } });
var trash = g("chooseBackupsToTrash")(old, 14);
eq(trash.length, 2, "сверх 14 копий уходят в корзину");
eq(trash[trash.length - 1].created, 0, "самая старая копия уходит в корзину");
assert(trash.every(function (item) { return item.name.indexOf("Чужая") === -1; }), "чужие файлы не трогаются");

folders["Бэкапы РемМонтаж"] = makeFolder("Бэкапы РемМонтаж");
for (var b = 0; b < 14; b++) {
  folders["Бэкапы РемМонтаж"].files.push({
    name: g("backupTitlePrefix")() + "старая " + b,
    getName: function () { return this.name; },
    getDateCreated: function () { return new Date(2026, 0, this.n); },
    n: b,
    trashed: false,
    setTrashed: function (v) { this.trashed = v; }
  });
}
var backupUrl = g("backupRegistry")();
assert(backupUrl.indexOf("https://docs.google.com") === 0, "бэкап возвращает ссылку на копию");
assert(book.sheets["Журнал"].values.some(function (row) {
  return String(row[4]).indexOf("Копия сохранена") >= 0 && String(row[4]).indexOf("abc12345tokenvalue") === -1;
}), "журнал бэкапа без ключа");
var trashedCount = folders["Бэкапы РемМонтаж"].files.filter(function (file) { return file.trashed; }).length;
eq(trashedCount, 1, "пятнадцатая старая копия уходит в корзину после новой");

nowQueue = [1000, 1000 + 5 * 60 * 1000];
props.ETL_CURSOR = "0";
book.sheets["Управление Франшизами"].values[2][3] = "АКТИВЕН";
g("continueETL")();
assert(book.sheets["Журнал"].values.some(function (row) { return String(row[4]).indexOf("Пауза на строке") >= 0; }), "лимит времени пишет паузу");
eq(props.ETL_CURSOR, "0", "курсор остаётся на непрочитанной строке");

var sql = fs.readFileSync(path.join(__dirname, "create_views.sql"), "utf8");
["active_franchises", "v_marketing_roas", "global_financial_summary", "view_inventory_abc_xyz", "view_staff"].forEach(function (name) {
  assert(sql.indexOf("analytics." + name) >= 0, "представление " + name + " есть");
});
assert(sql.indexOf('location = "EU"') >= 0, "SQL собран под EU");
assert(sql.indexOf("profit_before_cogs") >= 0, "прибыль в SQL без себестоимости");
assert(sql.indexOf('location = "US"') === -1, "регион US в запросах не задан");
assert(sql.indexOf("IFNULL(s.revenue, 0) - IFNULL(a.spend, 0) AS profit_before_cogs") >= 0, "прибыль в SQL — поступления минус реклама");

var client = fs.readFileSync(path.join(__dirname, "Client_Logic.html"), "utf8");
assert(client.indexOf("Прибыль до себестоимости") >= 0, "пульт подписывает прибыль до себестоимости");
var lay = fs.readFileSync(path.join(__dirname, "lay_operator_sheets.gs"), "utf8");
var order = ["Памятка", "Управление Франшизами", "Источники", "Витрина", "Правки", "Журнал", "Бланки", "Сырьё", "Вход 1С", "Словарь полей", "Служебное"];
var block = lay.slice(lay.indexOf("orderSheets(ss"));
var pos = 0;
order.forEach(function (name) {
  var at = block.indexOf('"' + name + '"', pos);
  assert(at > pos, "лист " + name + " стоит в порядке памятки");
  pos = at;
});
assert(String(g("runETL")).indexOf("backupRegistry") >= 0, "утренний сбор сначала копирует книгу");
var filler = fs.readFileSync(path.join(__dirname, "..", "moysklad", "fill_sheets.gs"), "utf8").replace(/\r\n/g, "\n");
assert(filler.indexOf("function myFunction()") === 0, "кнопка редактора по-прежнему открывает заполнение листов");
assert(filler.indexOf("fillCounterparties();\n  fillCashList();\n  fillHr();") >= 0, "кнопка редактора вызывает три заполнения");
assert(filler.indexOf("placeCustomerSample") === -1, "заполнение листов не запускает образец");
assert(filler.indexOf("backupRegistry") === -1, "заполнение листов не запускает бэкап");
var indexHtml = fs.readFileSync(path.join(__dirname, "Index.html"), "utf8");
assert(indexHtml.indexOf("Пульт РемМонтаж") >= 0, "страница пульта подписана");
assert(lay.indexOf("function layOperatorSheets") >= 0, "раскладка листов на месте");
assert(lay.indexOf("ОБРАЗЕЦ") >= 0, "памятка описывает образец");

var marker = new Sheet("Маркер", [["A"], ["данные"]]);
g("upsertMarkerRow")(marker, 2, "СПРАВКА", ["", "СПРАВКА"], "#fff");
eq(marker.values[1][1], "СПРАВКА", "справка вставляется второй строкой");
eq(marker.values[2][0], "данные", "прежняя строка сдвигается вниз");
g("upsertMarkerRow")(marker, 2, "СПРАВКА", ["", "СПРАВКА"], "#fff");
eq(marker.values.length, 3, "повтор справки не вставляет ещё строку");

g("orderSheets")(book, order);
eq(book.moved[0], "Памятка@1", "памятка становится первым листом");
eq(book.moved[book.moved.length - 1].indexOf("Служебное") === 0, true, "служебное остаётся последним");

var goods = g("imitationProducts")();
eq(goods.length, 5, "имитация товаров — пять строк");
assert(goods.every(function (row) { return String(row[0]).indexOf("ИМИТАЦИЯ") === 0 && String(row[6]) === ""; }), "у имитации есть пометка и нет ссылки");
var cashCols = g("imitationCashTargets")([
  ["Гм остаток", "", "2026-10-01", "2026-10-31", "Общая сумма поступления", "", "Контрагенты входящие ДС", "", "Общая сумма списания", "", "ГСМ", "", "Общая сумма поступления валюты", "", "Беларусские рубли", "Общая сумма списания валюты"],
  ["Операционная прибыль", "", "", "", "Остаток касса за последний период", "Сумма остатка на валютном счете", "", "Контрагент входящие ДС", "Документы", "", "Контрагент исходящие ДС", "Документы", "", "Контрагент входящие ДС", "Документы", "", "Контрагент исходящие ДС"],
  ["Дата", "Сумма остатка на р/сч", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]
]);
eq(cashCols.nameIn, 8, "имя поступления — столбец контрагента");
eq(cashCols.nameOut, 11, "имя списания — первый исходящий контрагент");
eq(cashCols.amountIn, 7, "детектор видит соседний с именем столбец");
eq(cashCols.amountOut, 9, "детектор видит столбец списания");
var octTop = ["Гм остаток", "", "2026-10-01", "2026-10-31", "Общая сумма поступления", "", "Контрагенты входящие ДС", "", "Общая сумма списания", "", "ГСМ"];
var octMid = ["Операционная прибыль", "", "", "", "Остаток касса за последний период", "Сумма остатка на валютном счете", "", "Контрагент входящие ДС", "Документы", "", "Контрагент исходящие ДС"];
assert(String(octTop[4]).indexOf("Общая сумма поступления") >= 0 && String(octMid[4]).indexOf("Остаток") >= 0, "колонка поступления одновременно подписана как остаток кассы");
assert(String(octMid[cashCols.amountOut - 1]).indexOf("Документы") >= 0, "колонка списания на живой форме подписана как документы");
assert(String(g("placeImitationSample")).indexOf("Октябрь не тронут") >= 0, "образец не пишет суммы в октябрь");
var productId = "1mMZlLOu9XP2BTzs_frWq6uAKzoqC_TC3bBJ1JmWzlyY";
books[productId] = new Book(productId);
books[productId].sheets["Товары"] = new Sheet("Товары", [["Наименование", "Артикул", "Группа", "Цена", "Единица", "Описание", "Ссылка", "Штрихкод"]]);
var reg = books[props.REGISTRY_ID];
reg.getSheetByName("Служебное").getRange(13, 1).setValue("Отчет Студии данных");
var placed = g("placeImitationSample")();
eq(placed.products, 5, "в книгу легло пять товаров");
eq(placed.cash, "skipped", "касса октября не заполняется");
var goodsSheet = books[productId].sheets["Товары"];
eq(goodsSheet.values.length, 6, "нет шестой строки-пояснения в товарах");
eq(goodsSheet.values[1][0], "ИМИТАЦИЯ. Кабель ВВГ 3х2.5", "первая строка — кабель");
eq(String(goodsSheet.values[1][6]), "", "у образца нет ссылки");
assert(String(goodsSheet.notes["1:1"]).indexOf("ИМИТАЦИЯ") >= 0, "пояснение висит заметкой у шапки");
eq(reg.getSheetByName("Витрина").values.filter(function (row) { return row[1] === "ИМИТАЦИЯ"; }).length, 5, "витрина показывает пять строк имитации");
eq(reg.getSheetByName("Служебное").getRange("B3").getValue(), "remmontazh-eu", "проект BigQuery на месте");
eq(reg.getSheetByName("Служебное").getRange("B13").getValue(), "https://datastudio.google.com/reporting/02d07e4d-d231-4bf9-9f8e-fc2c3c615e43", "ссылка отчёта записана в B13");
assert(JSON.stringify(reg.getSheetByName("Журнал").values).indexOf("abc12345tokenvalue") === -1, "журнал имитации без ключа");
g("placeImitationSample")();
eq(goodsSheet.values.length, 6, "повтор не добавляет товары");
eq(reg.getSheetByName("Памятка").values.filter(function (row) { return row[0] === "ИМИТАЦИЯ"; }).length, 1, "памятка об имитации одна");

reg.getSheetByName("Управление Франшизами").values[2][3] = "ОТКЛЮЧЕН";
reg.getSheetByName("Витрина").values = [
  ["код", "источник", "лист", "название", "дата", "приход", "расход", "кампания", "ссылка", "след"],
  ["FR-001", "ОБРАЗЕЦ", "Образец", "ТЕСТ. Оплата от заказчика", "2026-10-01", 15000, "", "", "", ""],
  ["FR-001", "ОБРАЗЕЦ", "Образец", "ТЕСТ. Оплата поставщику", "2026-10-02", "", 2600, "", "", ""],
  ["FR-001", "ИМИТАЦИЯ", "Товары", "ИМИТАЦИЯ. Кабель", "", "", 50, "", "", ""],
  ["FR-001", "Контрагенты", "Контрагенты", "Чужие деньги", "", 999, "", "", "", ""]
];
var offChart = g("reportTotals")(reg);
eq(offChart.rows, 2, "в график при выключенной франшизе входят две строки образца");
eq(offChart.moneyIn, 15000, "поступление графика — только образец");
eq(offChart.moneyOut, 2600, "списание графика — правка образца");
eq(offChart.ads, 0, "расход рекламы в графике ноль");
eq(offChart.profit, 15000, "прибыль графика не вычитает списание");
reg.getSheetByName("Служебное").appendRow(["Студия данных", "https://datastudio.google.com/reporting/02d07e4d-d231-4bf9-9f8e-fc2c3c615e43", "старый текст"]);
var drawn = g("placeReportCharts")();
eq(drawn.moneyIn, 15000, "построение графика не меняет сумму");
eq(drawn.charts, 2, "на листе два графика");
var chartSheet = reg.getSheetByName("Графики");
eq(chartSheet.getRange(9, 2).getFormula(), "=B6-B8", "прибыль = поступление минус реклама");
var flag = chartSheet.getRange(2, 11).getFormula();
assert(flag.indexOf("ОБРАЗЕЦ") >= 0 && flag.indexOf("АКТИВЕН") >= 0 && flag.indexOf("Включать") >= 0, "признак строки повторяет правила пульта");
assert((flag.match(/\(/g) || []).length === (flag.match(/\)/g) || []).length, "скобки формулы сходятся");
assert(flag.indexOf("abc12345tokenvalue") === -1, "формула графика без ключа");
eq(chartSheet.getRange(1, 1).getValue(), "Графики", "лист графиков подписан");
assert(String(chartSheet.getRange(2, 1).getValue()).indexOf("не факт оплаты") >= 0, "на графике видно, что это не факт оплаты");
eq(reg.getSheetByName("Управление Франшизами").getRange(3, 4).getValue(), "ОТКЛЮЧЕН", "франшиза остаётся выключенной");
eq(reg.getSheetByName("Служебное").getRange("B3").getValue(), "remmontazh-eu", "проект после графиков на месте");
eq(reg.getSheetByName("Служебное").getRange("B13").getValue(), "https://datastudio.google.com/reporting/02d07e4d-d231-4bf9-9f8e-fc2c3c615e43", "ссылка студии в B13 не стёрта");
var studio = reg.getSheetByName("Служебное").values.filter(function (row) { return row[0] === "Студия данных"; })[0];
eq(studio[1], "https://datastudio.google.com/reporting/02d07e4d-d231-4bf9-9f8e-fc2c3c615e43", "ссылка студии в инструкции на месте");
assert(String(studio[2]).indexOf("лист «Графики»") >= 0, "инструкция указывает на лист графиков");
assert(JSON.stringify(chartSheet.values).indexOf("abc12345tokenvalue") === -1, "лист графиков без ключа");
assert(JSON.stringify(reg.getSheetByName("Журнал").values).indexOf("abc12345tokenvalue") === -1, "журнал графиков без ключа");
eq(reg.getSheetByName("Памятка").values.filter(function (row) { return row[0] === "Графики"; }).length, 1, "памятка о графиках одна");
g("placeReportCharts")();
eq(reg.getSheetByName("Графики").getCharts().length, 2, "повтор не плодит графики");
eq(reg.getSheetByName("Памятка").values.filter(function (row) { return row[0] === "Графики"; }).length, 1, "повтор не плодит строку памятки");
assert(lay.indexOf("Рабочие графики — лист «Графики»") >= 0, "раскладка знает, что графики в таблице");
assert(filler.indexOf("placeReportCharts") === -1, "заполнение листов не строит графики");

console.log("passed " + passed + ", failed " + failed);
if (failed) process.exit(1);
