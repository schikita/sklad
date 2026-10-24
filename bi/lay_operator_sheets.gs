function layOperatorSheets() {
  var ss = registryBook();
  var green = "#188038";
  var amber = "#E37400";
  var blue = "#1A73E8";
  var gray = "#5F6368";
  var helpBg = "#E6F4EA";
  var money = "#F9AB00";

  layFranchise(ss, green, helpBg);
  laySources(ss, green);
  layDictionary(ss, gray, helpBg);
  layBlanks(ss, amber, helpBg);
  layJournal(ss, amber);
  layRaw(ss, blue);
  layVitrina(ss, green, money);
  layCorrections(ss, green, helpBg);
  layInbox(ss, blue);
  layService(ss, gray);
  layMemo(ss, green, amber, blue, gray);
  orderSheets(ss, [
    "Памятка",
    "Управление Франшизами",
    "Источники",
    "Витрина",
    "Правки",
    "Журнал",
    "Бланки",
    "Сырьё",
    "Вход 1С",
    "Словарь полей",
    "Служебное"
  ]);
  ss.setActiveSheet(ss.getSheetByName("Памятка"));
  Logger.log("готово " + ss.getUrl());
  return ss.getUrl();
}

function layFranchise(ss, green, helpBg) {
  var sh = ss.getSheetByName("Управление Франшизами");
  var headers = [
    "ID франшизы. Код вида FR-001. Тот же код стоит в листе «Источники». Это не порядковый номер 1, 2, 3.",
    "Название франшизы в этом реестре. Не логин МойСклад и не почта.",
    "Токен МойСклад. Секретный ключ API. Пустая ячейка — опрос МойСклад не идёт. Таблицы при этом читаются.",
    "Статус. АКТИВЕН — сбор идёт. ОШИБКА ТОКЕНА — ключ отвергнут. ЗАБЛОКИРОВАН и ОТКЛЮЧЕН — сбор выключен.",
    "Организация МойСклад. Имя фирмы из API. Не логин, не почта и не контрагент.",
    "UUID. Внутренний номер этой фирмы в МойСклад. Не пароль и не логин.",
    "Комментарий. На сбор не влияет."
  ];
  sh.getRange(1, 1, 1, 7).setValues([headers]).setNotes([[
    "Пишет человек при заведении франшизы. Робот ищет франшизу по этому коду.",
    "Подпись для людей и для пульта. В МойСклад таким именем организация может не называться.",
    "Робот читает ключ отсюда. В памятки и в журнал ключ не копируется.",
    "ОШИБКА ТОКЕНА робот ставит сам, если МойСклад ответил 401 или 403. Остальные статусы ставит человек.",
    "Откуда: ответ МойСклад, карточка организации, поле name. Робот пишет сам после принятого ключа. Контрагенты лежат отдельно, в «Витрине», источник МойСклад, сущность counterparty.",
    "Откуда: та же карточка организации, поле id. Карточка фирмы открывается так: online.moysklad.ru/app/#company/edit?id= и этот номер.",
    "Свободная пометка человека."
  ]]);
  styleHeader(sh, 7, green);
  sh.getRange(1, 5, 1, 2).setBackground("#174EA6");
  var help = [
    "",
    "Строка-справка. Робот её не читает. Ниже — рабочие франшизы.",
    "Секрет. Не пересылать и не диктовать.",
    "Красный статус — сначала «Журнал», потом этот столбец.",
    "Это наименование организации внутри МойСклад. Его вернул API по ключу слева. Логин, почта и контрагент здесь не пишутся.",
    "Номер нужен, чтобы открыть карточку фирмы в МойСклад. Паролем он не является.",
    "Можно писать пояснение для следующей смены.",
    "СПРАВКА"
  ];
  var row = upsertMarkerRow(sh, 8, "СПРАВКА", help, helpBg);
  sh.getRange(row, 5, 1, 2).setBackground("#D2E3FC");
  dropdown(sh, 4, ["АКТИВЕН", "ЗАБЛОКИРОВАН", "ОТКЛЮЧЕН", "ОШИБКА ТОКЕНА"]);
  sh.getRange(row, 4).clearDataValidations();
  paintRules(sh, [
    rule('=$D2="АКТИВЕН"', sh.getRange("D2:D300"), "#CEFAD0", "#0D652D"),
    rule('=$D2="ОШИБКА ТОКЕНА"', sh.getRange("D2:D300"), "#F4C7C3", "#7F1D1D"),
    rule('=$D2="ЗАБЛОКИРОВАН"', sh.getRange("D2:D300"), "#E8EAED", "#3C4043"),
    rule('=$D2="ОТКЛЮЧЕН"', sh.getRange("D2:D300"), "#E8EAED", "#3C4043")
  ]);
  sh.setFrozenRows(2);
  sh.setColumnWidths(1, 1, 260);
  sh.setColumnWidth(2, 280);
  sh.setColumnWidth(3, 280);
  sh.setColumnWidth(4, 280);
  sh.setColumnWidth(5, 340);
  sh.setColumnWidth(6, 320);
  sh.setColumnWidth(7, 260);
  sh.showColumns(1, 7);
  sh.hideColumns(8);
  sh.setTabColor(green);
}

function laySources(ss, green) {
  var sh = ss.getSheetByName("Источники");
  removeRowGroups(sh);
  var last = Math.max(sh.getLastRow(), 2);
  for (var r = last; r >= 2; r--) {
    if (String(sh.getRange(r, 9).getValue()) === "ГРУППА") sh.deleteRow(r);
  }
  var names = franchiseNames(ss);
  last = Math.max(sh.getLastRow(), 1);
  var width = Math.max(sh.getLastColumn(), 8);
  var values = last >= 2 ? sh.getRange(2, 1, last - 1, width).getValues() : [];
  var blocks = [];
  var current = null;
  for (var i = 0; i < values.length; i++) {
    if (!values[i][2]) continue;
    var fr = String(values[i][0] || "");
    var row = i + 2;
    if (!current || current.fr !== fr) {
      current = { fr: fr, start: row, end: row };
      blocks.push(current);
    } else current.end = row;
  }
  for (var b = blocks.length - 1; b >= 0; b--) {
    var block = blocks[b];
    sh.insertRowBefore(block.start);
    var title = names[block.fr] || "";
    var label = "Франшиза " + block.fr + (title ? " · " + title : "") + ". Минус слева сворачивает только её строки.";
    sh.getRange(block.start, 1, 1, 9).setValues([[block.fr, label, "", "", "", "", "", "", "ГРУППА"]]);
    sh.getRange(block.start, 1, 1, 8).setBackground(green).setFontColor("#FFFFFF").setFontWeight("bold").setWrap(true);
    sh.setRowHeight(block.start, 36);
  }
  sh.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
  last = sh.getLastRow();
  if (last >= 2) {
    var marks = sh.getRange(2, 9, last - 1, 1).getValues();
    var ids = sh.getRange(2, 3, last - 1, 1).getValues();
    for (var m = 0; m < marks.length; m++) {
      if (String(marks[m][0]) !== "ГРУППА") continue;
      var from = -1;
      var to = -1;
      for (var j = m + 1; j < marks.length; j++) {
        if (String(marks[j][0]) === "ГРУППА") break;
        if (ids[j][0]) {
          if (from < 0) from = j;
          to = j;
        }
      }
      if (from >= 0) sh.getRange(from + 2, 1, to - from + 1, 1).shiftRowGroupDepth(1);
    }
  }
  var headers = [
    "ID франшизы. Код FR-001 из листа «Управление Франшизами». Не номер строки 1, 2, 3.",
    "Имя источника на пульте. По нему же строка ищется в «Витрине».",
    "Идентификатор Google-таблицы, откуда робот читает. Менять только вместе со ссылкой.",
    "Имя листа внутри той таблицы. Текст должен совпасть символ в символ.",
    "Статус. Включать — строка идёт в свод. Исключить — робот её не читает.",
    "Тип бланка. Метка листа «Бланки»: какой столбец обязан быть в источнике.",
    "Ссылка на таблицу-источник.",
    "Номер строки, где лежат заголовки. Пусто — робот ищет сам. У накладных стоит 4.",
    "Метка группы. Робот не читает. У строки франшизы должно остаться ГРУППА."
  ];
  sh.getRange(1, 1, 1, 9).setValues([headers]);
  styleHeader(sh, 9, green);
  sh.setFrozenRows(1);
  dropdown(sh, 5, ["Включать", "Исключить"]);
  var groupMarks = sh.getLastRow() >= 2 ? sh.getRange(2, 9, sh.getLastRow() - 1, 1).getValues() : [];
  for (var g = 0; g < groupMarks.length; g++) {
    if (String(groupMarks[g][0]) === "ГРУППА") sh.getRange(g + 2, 5).clearDataValidations();
  }
  paintRules(sh, [
    rule('=$E2="Включать"', sh.getRange("E2:E500"), "#CEFAD0", "#0D652D"),
    rule('=$E2="Исключить"', sh.getRange("E2:E500"), "#E8EAED", "#3C4043")
  ]);
  sh.setColumnWidth(1, 260);
  sh.setColumnWidth(2, 360);
  sh.setColumnWidth(3, 280);
  sh.setColumnWidth(4, 240);
  sh.setColumnWidth(5, 220);
  sh.setColumnWidth(6, 260);
  sh.setColumnWidth(7, 280);
  sh.setColumnWidth(8, 240);
  sh.showColumns(1, 8);
  sh.hideColumns(9);
  sh.setTabColor(green);
}

function layDictionary(ss, gray, helpBg) {
  var sh = ss.getSheetByName("Словарь полей");
  var headers = [
    "НЕ ДЛЯ ОПЕРАТОРА. Заголовок столбца в исходной таблице. Текст должен совпасть целиком.",
    "Поле витрины: name, report_date, money_in, money_out, link, campaign_id.",
    "Тип значения: text, date, money, link."
  ];
  sh.getRange(1, 1, 1, 3).setValues([headers]).setNotes([[
    "Лист разработчика. Оператор цифры здесь не меняет.",
    "name — название, report_date — дата, money_in — поступление, money_out — списание, link — ссылка, campaign_id — кампания.",
    "money и date: если в источнике текст, ячейка витрины остаётся пустой, в журнале будет предупреждение."
  ]]);
  styleHeader(sh, 3, gray);
  var row = upsertMarkerRow(sh, 5, "СПРАВКА", [
    "",
    "Этот лист к работе оператора не относится. Его читает только робот, чтобы понять, какой столбец источника класть в витрину.",
    "Новую строку добавляет разработчик, когда в источнике появляется новый заголовок.",
    "",
    "СПРАВКА"
  ], "#E8EAED");
  var dialect = formulaDialect(sh);
  if (dialect) {
    var s = dialect.sep;
    var meaning = dialect.iff + '(A2:A=""' + s + '""' + s +
      dialect.iff + '(B2:B="name"' + s + '"Витрина, столбец Название. Тип: "&C2:C' + s +
      dialect.iff + '(B2:B="report_date"' + s + '"Витрина, столбец Дата. Тип: "&C2:C' + s +
      dialect.iff + '(B2:B="money_in"' + s + '"Витрина, столбец Поступление. Тип: "&C2:C' + s +
      dialect.iff + '(B2:B="money_out"' + s + '"Витрина, столбец Списание. Тип: "&C2:C' + s +
      dialect.iff + '(B2:B="link"' + s + '"Витрина, столбец Ссылка. Тип: "&C2:C' + s +
      dialect.iff + '(B2:B="campaign_id"' + s + '"Витрина, столбец Кампания. Тип: "&C2:C' + s +
      '"Поле вне стандартного набора. Без разработчика не менять."))))))';
    sh.getRange(row, 4).setFormula("=" + dialect.arr + "(" + meaning + ")");
    sh.getRange(1, 4).setValue("Куда ляжет в витрине. Формула, руками не писать.");
    sh.getRange(1, 4).setBackground(gray).setFontColor("#FFFFFF").setFontWeight("bold").setWrap(true);
  }
  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 420);
  sh.setColumnWidth(2, 360);
  sh.setColumnWidth(3, 220);
  sh.setColumnWidth(4, 360);
  sh.showColumns(1, 4);
  sh.hideColumns(5);
  sh.setTabColor(gray);
}

function layBlanks(ss, amber, helpBg) {
  var sh = ss.getSheetByName("Бланки");
  var headers = [
    "Тип бланка. Метка из листа «Источники», столбец «Тип бланка». Это не печатная форма и не договор.",
    "Обязательный столбец исходного листа. Нет столбца с таким именем — робот пропускает источник и пишет предупреждение в «Журнал».",
    "Что делает эта строка."
  ];
  sh.getRange(1, 1, 1, 3).setValues([headers]).setNote("Бланки проверки исходных таблиц. Сюда не кладут счета, договоры и кадровые сканы.");
  styleHeader(sh, 3, amber);
  var row = upsertMarkerRow(sh, 4, "СПРАВКА", [
    "СПРАВКА",
    "Бланки проверки чужих таблиц. Робот сверяет имя столбца. Печати, договора и платёжки здесь не хранятся.",
    "Тип «Накладные» и тип «Список» в этом листе не заданы, поэтому для них обязательный столбец не проверяется.",
    "СПРАВКА"
  ], "#FEF7E0");
  var help = {
    "Контрагенты": "Источник с этим типом читается только если в его листе есть столбец «Контрагенты».",
    "Товары": "Источник с этим типом читается только если в его листе есть столбец «Наименование».",
    "Движение": "Источник с этим типом читается только если в его листе есть столбец «Дата».",
    "Кадры": "Обязательный столбец пуст. Робот имя столбца не проверяет и читает кадровый лист как есть."
  };
  var last = sh.getLastRow();
  if (last >= 2) {
    var types = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < types.length; i++) {
      var key = String(types[i][0] || "");
      if (help[key]) sh.getRange(i + 2, 3).setValue(help[key]);
    }
  }
  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 360);
  sh.setColumnWidth(2, 420);
  sh.setColumnWidth(3, 420);
  sh.showColumns(1, 3);
  sh.hideColumns(4);
  sh.setTabColor(amber);
}

function layJournal(ss, amber) {
  var sh = ss.getSheetByName("Журнал");
  var headers = [
    "Время записи. Пояс Москва. Новые строки снизу.",
    "Франшиза. Код FR-00x. Пусто — запись про весь проект.",
    "Источник. Имя из «Источники», сущность МойСклад, BigQuery или Google Ads. Отсюда видно, откуда пришла проблема.",
    "Уровень. ок — шаг прошёл. заметка — не сбой. предупреждение — сбор жив, часть пропущена. ошибка — шаг не записан.",
    "Текст робота: что сделано или что сломалось.",
    "Что делать. Формула по столбцу E. Руками не писать, иначе формула погаснет."
  ];
  sh.getRange(1, 1, 1, 6).setValues([headers]);
  styleHeader(sh, 6, amber);
  var dialect = formulaDialect(sh);
  if (!dialect) {
    Logger.log("Формула журнала не принята, колонка «Что делать» пропущена");
    dialect = null;
  }
  if (dialect) {
  var s = dialect.sep;
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
  var expr = dialect.iff + '(D2:D="ошибка"' + s + '"Красная строка. Франшиза — столбец B, источник — столбец C, текст — столбец E. Цифры правят в источнике или в листе «Правки», не здесь."' + s +
    dialect.iff + '(D2:D="предупреждение"' + s + '"Сбор не остановлен. Часть листа или столбцов пропущена. Что именно — в столбце E."' + s +
    dialect.iff + '(D2:D="заметка"' + s + '"Это не сбой. Робот пишет, почему шаг пропущен или какой настройки не хватает."' + s + '"")))';
  for (var i = rules.length - 1; i >= 0; i--) {
    expr = dialect.iff + "(" + dialect.re + "(E2:E" + s + '"' + rules[i][0] + '")' + s + '"' + rules[i][1] + '"' + s + expr + ")";
  }
  expr = dialect.iff + "(" + dialect.len + "(E2:E)=0" + s + '""' + s + expr + ")";
  sh.getRange("F2").setFormula("=" + dialect.arr + "(" + expr + ")");
  }
  paintRules(sh, [
    rule('=$D2="ошибка"', sh.getRange("A2:F5000"), "#F4C7C3", "#7F1D1D"),
    rule('=$D2="предупреждение"', sh.getRange("A2:F5000"), "#FDE293", "#7A4F01")
  ]);
  if (!sh.getFilter()) sh.getRange(1, 1, Math.max(sh.getLastRow(), 2), 6).createFilter();
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 180);
  sh.setColumnWidth(3, 280);
  sh.setColumnWidth(4, 280);
  sh.setColumnWidth(5, 420);
  sh.setColumnWidth(6, 460);
  sh.setTabColor(amber);
}

function layRaw(ss, blue) {
  var sh = ss.getSheetByName("Сырьё");
  var headers = [
    "Код франшизы. Тот же FR-00x, что в «Управление Франшизами» и в «Витрине».",
    "Источник. Имя из «Источники» или МойСклад, если строка пришла по API.",
    "Лист исходной таблицы или сущность МойСклад.",
    "Когда робот забрал строку. Пояс Москва.",
    "Тип следа: sheet — строка таблицы. Иначе имя сущности МойСклад.",
    "JSON как пришло, до раскладки в витрину. Сбой сбора сюда не пишется, он в «Журнале». Красным здесь только обрезанная строка."
  ];
  sh.getRange(1, 1, 1, 6).setValues([headers]).setNotes([[
    "Робот перезаписывает лист при сборе. Ручная правка исчезнет.",
    "Сверяйте с именем источника, если в витрине чужая строка.",
    "counterparty — контрагенты, product — товары, paymentin — входящие платежи, paymentout — исходящие, customerorder — заказы, demand — отгрузки, employee — сотрудники.",
    "Время записи следа, не дата документа.",
    "По этой метке видно, таблица это или API.",
    "Если внутри JSON есть truncated, МойСклад отдал строку длиннее 45000 знаков и она обрезана. Остальные сбои смотрите в «Журнале»."
  ]]);
  styleHeader(sh, 6, blue);
  var dialect = formulaDialect(sh);
  var sep = dialect ? dialect.sep : ";";
  var fn = dialect && dialect.isn === "ЕЧИСЛО"
    ? '=ЕЧИСЛО(ПОИСК("truncated"' + sep + "$F2))"
    : '=ISNUMBER(SEARCH("truncated"' + sep + "$F2))";
  paintRules(sh, [rule(fn, sh.getRange("A2:F5000"), "#F4C7C3", "#7F1D1D")]);
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, 6, 280);
  sh.setColumnWidth(6, 420);
  sh.setTabColor(blue);
}

function layVitrina(ss, green, money) {
  var sh = ss.getSheetByName("Витрина");
  var headers = [
    "Код франшизы. Тот же FR-00x, что в «Управление Франшизами», столбец A.",
    "Источник. Имя из «Источники» или МойСклад для строк API.",
    "Лист источника или сущность МойСклад. Отсюда видно, откуда строка.",
    "Название. Контрагент, товар, сотрудник или документ. Берётся из столбца, который словарь пометил как name.",
    "Дата документа или операции. У справочников часто пусто — это нормально.",
    "Поступление. Деньги в. Смотреть в первую очередь. Платёж МойСклад paymentin робот делит на 100, потому что API отдаёт копейки.",
    "Списание. Деньги из. Смотреть в первую очередь. paymentout и столбцы money_out. Сумма накладной по словарю ложится сюда, это не поступление.",
    "Id рекламной кампании из источника. Расход рекламы в этот столбец сам не пишется.",
    "Ссылка на карточку или на ячейку источника.",
    "Строка как пришла. Спорную цифру сверяйте здесь и в источнике. truncated — строка МойСклад обрезана."
  ];
  sh.getRange(1, 1, 1, 10).setValues([headers]).setNotes([[
    "Пульт группирует строки по этому коду.",
    "Следующий сбор перезаписывает строки этого источника. Постоянную замену суммы кладите в «Правки».",
    "Не путать с листом этой книги. Это имя листа в чужой таблице либо сущность API.",
    "Пустое название значит, что в источнике не нашлось текстового столбца.",
    "Формат даты в витрине — гггг-мм-дд.",
    "Контрагенты, товары и сотрудники денег здесь не имеют. Пусто — нормально.",
    "Не прибавляйте списание к поступлению. На пульте прибыль считается из поступлений минус расход рекламы, без себестоимости.",
    "Пока выгрузка Google Ads не сохранена, расхода в системе нет.",
    "Открывайте ссылку, если нужно увидеть карточку, а не свод.",
    "Технический след. Оператору нужен, только когда цифра спорит с источником."
  ]]);
  styleHeader(sh, 10, green);
  sh.getRange(1, 6, 1, 2).setBackground(money).setFontColor("#202124");
  var dialect = formulaDialect(sh);
  var sep = dialect ? dialect.sep : ";";
  var fn = dialect && dialect.isn === "ЕЧИСЛО"
    ? '=ЕЧИСЛО(ПОИСК("truncated"' + sep + "$J2))"
    : '=ISNUMBER(SEARCH("truncated"' + sep + "$J2))";
  paintRules(sh, [rule(fn, sh.getRange("A2:J5000"), "#F4C7C3", "#7F1D1D")]);
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 240);
  sh.setColumnWidth(2, 220);
  sh.setColumnWidth(3, 240);
  sh.setColumnWidth(4, 280);
  sh.setColumnWidth(5, 240);
  sh.setColumnWidth(6, 300);
  sh.setColumnWidth(7, 320);
  sh.setColumnWidth(8, 240);
  sh.setColumnWidth(9, 240);
  sh.setColumnWidth(10, 360);
  sh.setTabColor(green);
}

function layCorrections(ss, green, helpBg) {
  var sh = ss.getSheetByName("Правки");
  var headers = [
    "Ключ строки: франшиза, id таблицы, лист, название и дата через |. По нему робот находит строку витрины.",
    "Поле. money_in — поступление. money_out — списание. Другие поля робот с этого листа в витрину не переносит.",
    "Число, которое держится вместо посчитанного.",
    "Когда человек сохранил правку. Новая запись с тем же ключом и полем затирает старую."
  ];
  sh.getRange(1, 1, 1, 4).setValues([headers]).setNotes([[
    "Ключ собирает пульт. Руками его лучше не переписывать: робот перестанет находить строку.",
    "Это не список всех правок бизнеса. Только замена поступления или списания.",
    "Пустое число очищает сумму.",
    "Автоматического поиска ошибок нет. Строку создаёт человек."
  ]]);
  styleHeader(sh, 4, green);
  upsertMarkerRow(sh, 5, "СПРАВКА", [
    "",
    "Строку создаёт человек на пульте. Робот её сам не придумывает. При каждом сборе он заново подставляет это число в витрину, чтобы новый опрос не стёр замену.",
    "Смотреть сюда, если в витрине цифра не такая, как в источнике: её могли подменить здесь.",
    "Время последнего сохранения. История прежних значений не хранится.",
    "СПРАВКА"
  ], helpBg);
  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 360);
  sh.setColumnWidth(2, 360);
  sh.setColumnWidth(3, 280);
  sh.setColumnWidth(4, 280);
  sh.showColumns(1, 4);
  sh.hideColumns(5);
  sh.setTabColor(green);
}

function layInbox(ss, blue) {
  var sh = ss.getSheetByName("Вход 1С");
  var headers = [
    "Время прихода посылки. Пояс Москва.",
    "Тело посылки, как пришло, до 45000 знаков.",
    "Куда это идёт."
  ];
  sh.getRange(1, 1, 1, 3).setValues([headers]).setNotes([[
    "Строку добавляет адрес робота, когда на него приходит POST.",
    "Это сырой JSON. Разбора проводок здесь нет.",
    "В витрину, в сырьё и в прибыль эта посылка сама не попадает."
  ]]);
  styleHeader(sh, 3, blue);
  var dialect = formulaDialect(sh);
  if (dialect) {
    var s = dialect.sep;
    var text = "В витрину само не попадает. Это архив POST на адрес робота. Проводки отсюда не создаются.";
    sh.getRange("C2").setFormula("=" + dialect.arr + "(" + dialect.iff + "(" + dialect.len + "(B2:B)=0" + s + '""' + s + '"' + text + '"))');
  }
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 240);
  sh.setColumnWidth(2, 420);
  sh.setColumnWidth(3, 420);
  sh.setTabColor(blue);
}

function layService(ss, gray) {
  var sh = ss.getSheetByName("Служебное");
  var notes = {
    "Робот": "Почта, под которой работает скрипт. Доступ к книгам источников должен быть у неё. Робот ячейку не перечитывает как команду.",
    "Доступ сотрудника": "Сотрудника пускают читателем на почту робота, не на личную. Робот ячейку не исполняет.",
    "Доступ сотрудников": "Сотрудника пускают читателем на почту робота, не на личную. Робот ячейку не исполняет.",
    "Проект BigQuery": "Робот проверяет только, что ячейка не пуста. Пусто — в журнале заметка, свод остаётся на витрине.",
    "Регион": "Регион таблиц BigQuery. Держать EU. US сломает запросы: они собраны под EU. Робот регион из ячейки не подставляет.",
    "Customer ID Google Ads": "Робот проверяет только заполненность. Пусто — в журнале заметка, расход рекламы не подтягивается.",
    "Реестр": "Адрес этой книги. Робот записывает его при установке.",
    "Ключ склейки": "Правило стыковки выручки и рекламы: дата плюс id кампании. Робот текст правила не исполняет, оно для людей и для SQL.",
    "Прибыль на экране": "Правило пульта. Себестоимость в эту формулу не входит.",
    "Кто видит пульт": "Кто открывает пульт. На сбор таблиц не влияет.",
    "Обновление": "Расписание. Исполняет триггер runETL: каждый день в 06:00 и кнопка на пульте.",
    "Себестоимость": "На экране не вычитается, пока нет отдельного решения.",
    "Главный аккаунт": "Учётная запись Google, на которой собран контур."
  };
  var last = Math.max(sh.getLastRow(), 1);
  var body = sh.getRange(1, 1, last, 3).getValues();
  for (var i = 0; i < body.length; i++) {
    var label = String(body[i][0] || "");
    if (!label || label === "ИНСТРУКЦИЯ") continue;
    if (notes[label] && !String(body[i][2] || "")) sh.getRange(i + 1, 3).setValue(notes[label]).setWrap(true);
    else if (!notes[label] && !String(body[i][2] || "") && i >= 12) {
      sh.getRange(i + 1, 3).setValue("Робот эту строку не читает. Это запись человека.").setWrap(true);
    }
  }
  sh.setColumnWidth(1, 260);
  sh.setColumnWidth(2, 420);
  sh.setColumnWidth(3, 520);
  var cut = 0;
  last = sh.getLastRow();
  var colA = sh.getRange(1, 1, last, 1).getValues();
  for (var a = 13; a < colA.length; a++) {
    if (String(colA[a][0]) === "ИНСТРУКЦИЯ") cut = a + 1;
  }
  if (cut && sh.getLastRow() >= cut) sh.getRange(cut, 1, sh.getLastRow() - cut + 1, 4).clearContent();
  var start = cut || (sh.getLastRow() + 2);
  if (start < 14) start = 14;
  var project = String(sh.getRange("B3").getValue() || "");
  var region = String(sh.getRange("B4").getValue() || "");
  var ads = String(sh.getRange("B5").getValue() || "");
  var looker = "";
  var scan = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  for (var s = 0; s < scan.length; s++) {
    var blob = String(scan[s][0] || "") + " " + String(scan[s][1] || "");
    if (blob.indexOf("datastudio.google.com") >= 0 || blob.indexOf("lookerstudio.google.com") >= 0) looker = String(scan[s][1] || "");
  }
  if (!looker) looker = "https://datastudio.google.com/reporting/02d07e4d-d231-4bf9-9f8e-fc2c3c615e43";
  var rows = [
    ["ИНСТРУКЦИЯ", "Лист не для оператора. Рабочий порядок листов — на «Памятке».", "Ниже факты контура. Робот исполняет только B3, B5 и B6."],
    ["Что читает робот", "B3 проект. B5 кабинет Ads. B6 адрес этой книги.", "Остальные ячейки — память настроек. Пустая B3 или B5 даёт заметку в «Журнале», сбор таблиц не останавливает."],
    ["Проект Google Cloud", project, "Имя remmontazh, номер 908335923939. Оплаты нет. Песочница. Карту не привязывать. Start free и Upgrade не нажимать."],
    ["Регион BigQuery", region, "Датасеты raw_data и analytics. Запросы собраны под EU."],
    ["Консоль BigQuery", "https://console.cloud.google.com/bigquery?project=" + (project || "remmontazh-eu"), "Человек открывает сам. Робот эту ссылку не читает."],
    ["Кабинет Google Ads", ads, "Валюта USD, пояс Europe/Minsk. Кампаний нет. Номер в кабинете пишется с дефисами, в выгрузке BigQuery — без дефисов."],
    ["Выгрузка расхода", "Не сохранена.", "Форма ads-spend в датасет raw_data заполнена. Google остановил вход на создании ключа доступа. Ключ не создавался. Расход на пульте 0. Таблица ads_campaign_stats пустая."],
    ["Пульт", "https://script.google.com/macros/s/AKfycbyrI9lD_y0La2LAlkWhrR0cra0V4yZSG6zeRDK72NFLlQYRrtsirsUhARjVqVBQTDDscw/exec", "Прибыль на экране: поступление минус расход рекламы, без себестоимости. Видит владелец проекта."],
    ["Скрипт", "https://script.google.com/u/0/home/projects/1filQcJ57U7dOq-bmQO0712WdMDUz4YtAomRncw9XprNGi-4PY3Cqb1mf/edit", "Сбор — runETL, каждый день в 06:00 и кнопка пульта. Подписи листов заново ставит layOperatorSheets. Ключи и B3:B12 она не трогает."],
    ["Студия данных", looker, "Ссылка на отчёт. В этой стране сервис не открывается. Рабочие графики — лист «Графики». Робот ссылку не читает."],
    ["МойСклад", "Имя фирмы и UUID пишутся в «Управление Франшизами», столбцы E и F.", "Логин в этот лист не кладётся. Ключ лежит в столбце C того листа. Контрагент — это не организация."],
    ["1С", "Входящий POST падает в лист «Вход 1С».", "В витрину посылка сама не идёт. Адреса базы 1С в контуре пока нет."],
    ["Куда добавлять новое", "Новая франшиза — строка в «Управление Франшизами», затем её источники.", "После новой франшизы запустите layOperatorSheets: она заново свернёт группы «Источники»."],
    ["Куда добавлять новое", "Новый файл — строка в «Источники», статус Включать.", "Обязательный столбец — лист «Бланки». Новый заголовок — лист «Словарь полей». Замена суммы — лист «Правки»."],
    ["Что не включать", "Оплату Google, пробные 300 долларов, кампанию Ads.", "Регион EU не менять на US."]
  ];
  sh.getRange(start, 1, rows.length, 3).setValues(rows).setWrap(true).setVerticalAlignment("top");
  sh.getRange(start, 1, 1, 3).setBackground(gray).setFontColor("#FFFFFF").setFontWeight("bold");
  sh.getRange(start + 1, 1, rows.length - 1, 3).setBackground("#F8F9FA");
  sh.setTabColor(gray);
}

function layMemo(ss, green, amber, blue, gray) {
  var sh = ss.getSheetByName("Памятка");
  if (!sh) sh = ss.insertSheet("Памятка");
  sh.clear();
  var rows = [
    ["Памятка. Книга «Мастер-реестр РемМонтаж»", "", "", "", "", ""],
    ["Цвет вкладки", "Кому лист", "Лист", "За что отвечает", "Что делать", "Что не делать"],
    ["Зелёный", "Оператор, каждый день", "Управление Франшизами", "Одна строка — одна франшиза. Код, статус, имя фирмы МойСклад и её UUID.", "Следить, чтобы статус был АКТИВЕН. При красном статусе открыть «Журнал».", "Не стирать ключ. Не писать в UUID руками. Не путать организацию с контрагентом."],
    ["Зелёный", "Оператор, каждый день", "Источники", "Какие таблицы робот читает. Строки свернуты по франшизе: минус слева прячет одну франшизу.", "Ставить Включать или Исключить.", "Не менять id таблицы и имя листа наугад. Код франшизы — это FR-001, не номер 1."],
    ["Зелёный", "Оператор, каждый день", "Витрина", "Свод для пульта. Жёлтые столбцы — поступление и списание.", "Смотреть название, дату, две денежные колонки и ссылку.", "Не править цифру здесь насовсем: сбор перезапишет источник. Постоянная замена — лист «Правки»."],
    ["Зелёный", "Оператор, проверять", "Правки", "Ручная замена поступления или списания. Человек сохраняет на пульте. Робот при каждом сборе подставляет её снова.", "Если цифра в витрине не как в источнике — искать её здесь, смотреть время в столбце «Когда».", "Не ждать, что робот сам найдёт ошибку и заведёт строку. Сам он строки не создаёт."],
    ["Янтарный", "Оператор, когда сбор странный", "Журнал", "Ход скрипта. Красная строка — шаг не записан. Жёлтая — сбор жив, часть пропущена.", "Читать столбцы Франшиза, Источник и «Что делать».", "Не удалять красную строку, пока причина не закрыта. Цифры здесь не правят."],
    ["Янтарный", "Оператор, когда журнал пишет «Нет столбца»", "Бланки", "Проверка чужой таблицы: какой столбец в ней обязан быть. Не договоры и не печатные бланки.", "Сверить имя столбца с исходным листом.", "Не складывать сюда сканы и счета."],
    ["Синий", "Технический след", "Сырьё", "Копия строк до раскладки в витрину. Показывает, откуда взялся свод.", "Открывать, только если цифра спорит: столбец JSON — что реально пришло.", "Не править. Робот перезаписывает лист. Сбой сбора лежит в «Журнале», не здесь."],
    ["Синий", "Технический след", "Вход 1С", "Архив посылок, которые пришли на адрес робота.", "Смотреть, была ли посылка и когда.", "Не ждать проводок. В витрину эти строки сами не идут."],
    ["Серый", "Не для оператора", "Словарь полей", "Перевод заголовка исходной таблицы в поле витрины. Лист разработчика.", "Не открывать для смены цифр.", "Не переименовывать поля name, money_in, money_out."],
    ["Серый", "Не для оператора", "Служебное", "Настройки проекта, BigQuery, рекламы, пульта и расписания. Полная выкладка — внизу того листа, блок ИНСТРУКЦИЯ.", "Не открывать в смене.", "Не включать оплату Google. Не менять регион EU. Не стирать B3 и B5."],
    ["", "", "", "", "", ""],
    ["Организация МойСклад", "Имя фирмы из карточки организации МойСклад.", "Пишет робот в лист «Управление Франшизами», столбец E, после того как ключ принят.", "Это не логин, не почта и не контрагент. Контрагенты лежат в «Витрине», источник МойСклад, сущность counterparty.", "", ""],
    ["UUID", "Внутренний номер той же карточки организации.", "Столбец F того же листа. По нему открывается фирма в МойСклад.", "Это не пароль, не логин и не токен.", "", ""],
    ["Как идёт сбор", "1. Каждый день в 06:00 и по кнопке пульта.", "2. Берутся франшизы со статусом АКТИВЕН и источники со статусом Включать.", "3. «Бланки» проверяют обязательный столбец. «Словарь» раскладывает ячейки. «Правки» заново подставляют сохранённые суммы.", "4. МойСклад опрашивается ключом. Имя фирмы и UUID обновляются.", "5. Итог — «Витрина». След — «Сырьё». Запись о шаге — «Журнал»."],
    ["Каждый день смотреть", "Журнал: есть ли красная строка.", "Витрина: поступление и списание.", "Правки: не появилась ли подмена.", "Управление франшизами: статус не стал ли ОШИБКА ТОКЕНА.", ""],
    ["Можно не открывать", "Словарь полей.", "Служебное.", "Сырьё, пока цифра не спорит с источником.", "Вход 1С, пока посылок нет.", ""],
    ["ОБРАЗЕЦ", "Две строки с пометкой ТЕСТ. Это не факт оплаты.", "Источник «Образец»: поступление 15 000 и списание 2 500.", "На витрине списание 2 600. Разницу держит лист «Правки».", "Перед сбором в 06:00 робот кладёт копию книги в папку «Бэкапы РемМонтаж». Хранятся 14 копий.", "Не принимать эти суммы за кассу и не стирать пометку ТЕСТ."]
  ];
  sh.getRange(1, 1, rows.length, 6).setValues(rows).setWrap(true).setVerticalAlignment("top");
  sh.getRange(1, 1, 1, 6).merge().setBackground(green).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(14);
  sh.setRowHeight(1, 36);
  sh.getRange(2, 1, 1, 6).setBackground("#202124").setFontColor("#FFFFFF").setFontWeight("bold");
  var colors = ["#E6F4EA", "#E6F4EA", "#E6F4EA", "#E6F4EA", "#FEF7E0", "#FEF7E0", "#E8F0FE", "#E8F0FE", "#F1F3F4", "#F1F3F4"];
  for (var i = 0; i < colors.length; i++) sh.getRange(i + 3, 1, 1, 6).setBackground(colors[i]);
  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 180);
  sh.setColumnWidth(2, 220);
  sh.setColumnWidth(3, 240);
  sh.setColumnWidth(4, 380);
  sh.setColumnWidth(5, 340);
  sh.setColumnWidth(6, 340);
  sh.setTabColor(green);
  for (var r = 3; r <= rows.length; r++) sh.setRowHeight(r, 72);
  sh.setRowHeight(14, 48);
  sh.setRowHeight(15, 48);
}

function orderSheets(ss, names) {
  var pos = 1;
  for (var i = 0; i < names.length; i++) {
    var sh = ss.getSheetByName(names[i]);
    if (!sh) continue;
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(pos);
    pos++;
  }
}

function franchiseNames(ss) {
  var sh = ss.getSheetByName("Управление Франшизами");
  var values = sh.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0] || String(values[i][7] || "") === "СПРАВКА") continue;
    map[String(values[i][0])] = String(values[i][1] || "");
  }
  return map;
}

function upsertMarkerRow(sheet, markerCol, marker, values, bg) {
  var last = Math.max(sheet.getLastRow(), 1);
  var col = sheet.getRange(1, markerCol, last, 1).getValues();
  var row = 0;
  for (var i = 0; i < col.length; i++) if (String(col[i][0]) === marker) row = i + 1;
  if (!row) {
    sheet.insertRowBefore(2);
    row = 2;
  }
  sheet.getRange(row, 1, 1, values.length).setValues([values]);
  sheet.getRange(row, 1, 1, values.length).setWrap(true).setBackground(bg).setFontColor("#202124").setVerticalAlignment("middle");
  sheet.setRowHeight(row, 84);
  return row;
}

function styleHeader(sheet, cols, bg) {
  var range = sheet.getRange(1, 1, 1, cols);
  range.setBackground(bg).setFontColor("#FFFFFF").setFontWeight("bold").setWrap(true).setVerticalAlignment("middle");
  sheet.setRowHeight(1, 96);
  sheet.setFrozenRows(1);
}

function paintRules(sheet, built) {
  var rules = sheet.getConditionalFormatRules();
  var kept = [];
  for (var i = 0; i < rules.length; i++) {
    var cond = rules[i].getBooleanCondition();
    var formula = cond ? String(cond.getCriteriaValues()[0] || "") : "";
    if (formula.indexOf("ошибка") !== -1 || formula.indexOf("ОШИБКА") !== -1 || formula.indexOf("предупреждение") !== -1 || formula.indexOf("truncated") !== -1 || formula.indexOf("АКТИВЕН") !== -1 || formula.indexOf("Включать") !== -1 || formula.indexOf("Исключить") !== -1 || formula.indexOf("ЗАБЛОКИРОВАН") !== -1 || formula.indexOf("ОТКЛЮЧЕН") !== -1) continue;
    kept.push(rules[i]);
  }
  sheet.setConditionalFormatRules(kept.concat(built));
}

function rule(formula, range, bg, color) {
  return SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setBackground(bg)
    .setFontColor(color)
    .setRanges([range])
    .build();
}

function formulaDialect(sheet) {
  var tests = [
    { sep: ",", iff: "IF", len: "LEN", re: "REGEXMATCH", arr: "ARRAYFORMULA", isn: "ISNUMBER", search: "SEARCH", probe: "=IF(TRUE,1,0)", reProbe: '=REGEXMATCH("не найден","не найден")' },
    { sep: ";", iff: "IF", len: "LEN", re: "REGEXMATCH", arr: "ARRAYFORMULA", isn: "ISNUMBER", search: "SEARCH", probe: "=IF(TRUE;1;0)", reProbe: '=REGEXMATCH("не найден";"не найден")' },
    { sep: ";", iff: "ЕСЛИ", len: "ДЛСТР", re: "РЕГВЫРАЖ", arr: "ARRAYFORMULA", isn: "ЕЧИСЛО", search: "ПОИСК", probe: "=ЕСЛИ(ИСТИНА;1;0)", reProbe: '=РЕГВЫРАЖ("не найден";"не найден")' },
    { sep: ";", iff: "ЕСЛИ", len: "ДЛСТР", re: "REGEXMATCH", arr: "МАССИВ", isn: "ЕЧИСЛО", search: "ПОИСК", probe: "=ЕСЛИ(ИСТИНА;1;0)", reProbe: '=REGEXMATCH("не найден";"не найден")' }
  ];
  for (var i = 0; i < tests.length; i++) {
    if (probeFormula(sheet, tests[i].probe) !== "1") continue;
    var re = probeFormula(sheet, tests[i].reProbe);
    if (re === "TRUE" || re === "ИСТИНА" || re === "true") return tests[i];
  }
  return null;
}

function probeFormula(sheet, formula) {
  var cell = sheet.getRange("T1");
  var oldFormula = cell.getFormula();
  var oldValue = cell.getValue();
  cell.clearContent();
  cell.setFormula(formula);
  SpreadsheetApp.flush();
  var shown = String(cell.getDisplayValue() || "");
  cell.clearContent();
  if (oldFormula) cell.setFormula(oldFormula);
  else if (oldValue !== "") cell.setValue(oldValue);
  return shown;
}

function removeRowGroups(sheet) {
  var max = Math.max(Math.min(sheet.getMaxRows(), sheet.getLastRow() + 15), 2);
  for (var n = 0; n < 40; n++) {
    var removed = false;
    for (var r = 1; r <= max; r++) {
      var depth = 0;
      try { depth = sheet.getRowGroupDepth(r); } catch (err) { depth = 0; }
      if (depth > 0) {
        try {
          sheet.getRowGroup(r, depth).remove();
          removed = true;
          break;
        } catch (err2) {}
      }
    }
    if (!removed) break;
    max = Math.max(sheet.getMaxRows(), 1);
  }
}
