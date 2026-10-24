function myFunction() {
  fillCounterparties();
  fillCashList();
  fillHr();
}

function msCompany(id) {
  return "https://online.moysklad.ru/app/#company/edit?id=" + id;
}

function msEmployee(id) {
  return "https://online.moysklad.ru/app/#employee/edit?id=" + id;
}

function card(url) {
  return '=HYPERLINK("' + url + '";"карточка МойСклад")';
}

function person(url, title) {
  return '=HYPERLINK("' + url + '";"' + title + '")';
}

function fillCounterparties() {
  var sh = SpreadsheetApp.openById("1PBYZcU1TNnkeYfueQLoNDOY0Owou6crqEEIsSVhSKZg").getSheetByName("Контрагенты");
  var rows = [
    ["Спец NET", "", "+375 298 34 75 47", "krasavin.group@gmail.com", "", card(msCompany("e16a27ab-cbdc-11f1-0a80-16b600001d8f")), "", ""],
    ["ЛНК Квантет ООО", "", "+7 961 713 8422", "ink.kvantet@rambler.ru", "", card(msCompany("18f28cfe-cbdd-11f1-0a80-0edc00002e1d")), "", ""],
    ["Все для сварки Ц ООО", "", "3842 57 03 03, svarkamira@yandex.ru", "Демаков В.И.", "", card(msCompany("3d418d2b-cbdd-11f1-0a80-0edc0000304f")), "", ""],
    ["Игнатенко Оксана Александровна", "", "+7 926 065 7296, oksana.nkr@mail.ru", "Демаков В.И.", "", card(msCompany("6aa17d92-cbdd-11f1-0a80-158c00002b07")), "", ""],
    ["ИП Борисов Д.К.", "", "troyans5@rambler.ru", "Демаков В.И.", "", card(msCompany("8e3c50be-cbdd-11f1-0a80-158c00002bd9")), "", ""],
    ['Частное предприятие "Амикгрупп"', "", "+375 29 834 73 47", "Демаков В.И.", "", card(msCompany("ade7ddff-cbdd-11f1-0a80-03bd00002cd2")), "", ""],
    ['Государственное предприятие "Национальный центр поддержки экспорта"', "", "+375 212 02 11", "Демаков В.И.", "", card(msCompany("d19be611-cbdd-11f1-0a80-117200002cbf")), "", ""]
  ];
  sh.getRange(4, 1, rows.length, 8).setValues(rows);
}

function fillCashList() {
  var sh = SpreadsheetApp.openById("1SlIUkncDyrhpvs7FPmtWY6zvX684nbiO-s0bVrzFOGQ").getSheetByName("Список контрагентов");
  var org = "РУСУГЛЕНЕФТЕГАЗ О.";
  var names = [
    ["Спец NET", org, "", card(msCompany("e16a27ab-cbdc-11f1-0a80-16b600001d8f"))],
    ["ЛНК Квантет ООО", org, "", card(msCompany("18f28cfe-cbdd-11f1-0a80-0edc00002e1d"))],
    ["Все для сварки Ц ООО", org, "", card(msCompany("3d418d2b-cbdd-11f1-0a80-0edc0000304f"))],
    ["Игнатенко Оксана Александровна", org, "", card(msCompany("6aa17d92-cbdd-11f1-0a80-158c00002b07"))],
    ["ИП Борисов Д.К.", org, "", card(msCompany("8e3c50be-cbdd-11f1-0a80-158c00002bd9"))],
    ['Частное предприятие "Амикгрупп"', org, "", card(msCompany("ade7ddff-cbdd-11f1-0a80-03bd00002cd2"))],
    ['Государственное предприятие "Национальный центр поддержки экспорта"', org, "", card(msCompany("d19be611-cbdd-11f1-0a80-117200002cbf"))]
  ];
  sh.getRange(2, 1, names.length, 4).setValues(names);
}

function fillHr() {
  var ss = SpreadsheetApp.openById("1Dy844ebEUk-H-MYeKTQpWN6Wsh-UIT05HaoxtfABSh8");
  ss.getSheetByName("Ген. Директор").getRange(4, 2).setFormula(person(msEmployee("19439d6b-cbde-11f1-0a80-1e3700002f46"), "Демаков В.И."));
  ss.getSheetByName("Бухгалтер").getRange(4, 2).setFormula(person(msEmployee("463debae-cbde-11f1-0a80-11720000333b"), "Хлебникова Ю.А."));
  ss.getSheetByName("Слесарь-ремонтник № 1").getRange(4, 2).setFormula(person(msEmployee("7ff6dfcb-cbde-11f1-0a80-182b00002f54"), "Дмитриев Михаил Андреевич"));
}
