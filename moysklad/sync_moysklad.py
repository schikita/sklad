# -*- coding: utf-8 -*-
"""Выгрузка пробного МойСклад в CSV под текущие таблицы РемМонтаж.

Доступ: переменная MOYSKLAD_TOKEN либо пара MOYSKLAD_LOGIN и MOYSKLAD_PASSWORD.
Запуск из этой папки:
  python sync_moysklad.py
  python sync_moysklad.py --month 2026-09
  python sync_moysklad.py --import-products ..\\товары.csv
"""

import argparse
import base64
import csv
import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

API = "https://api.moysklad.ru/api/remap/1.2"
ROOT = Path(__file__).resolve().parent
OUT = ROOT / "out"
MAPPING = json.loads((ROOT / "mapping.json").read_text(encoding="utf-8"))


def headers():
    token = os.environ.get("MOYSKLAD_TOKEN", "").strip()
    if token:
        auth = "Bearer " + token
    else:
        login = os.environ.get("MOYSKLAD_LOGIN", "").strip()
        password = os.environ.get("MOYSKLAD_PASSWORD", "").strip()
        if not login or not password:
            sys.exit(
                "Нет доступа. Задайте MOYSKLAD_TOKEN "
                "или MOYSKLAD_LOGIN и MOYSKLAD_PASSWORD."
            )
        raw = base64.b64encode(f"{login}:{password}".encode("utf-8")).decode("ascii")
        auth = "Basic " + raw
    return {
        "Authorization": auth,
        "Accept": "application/json;charset=utf-8",
        "Content-Type": "application/json",
    }


def request(method, path, body=None):
    data = None if body is None else json.dumps(body).encode("utf-8")
    req = urllib.request.Request(API + path, data=data, headers=headers(), method=method)
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            payload = resp.read().decode("utf-8")
            return json.loads(payload) if payload else {}
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        sys.exit(f"МойСклад ответил {exc.code}: {detail[:500]}")


def get_all(path, extra=""):
    rows = []
    offset = 0
    while True:
        sep = "&" if "?" in path or extra else "?"
        query = f"{path}{extra}{sep}limit=1000&offset={offset}"
        page = request("GET", query)
        batch = page.get("rows", [])
        rows.extend(batch)
        if len(batch) < 1000:
            return rows
        offset += 1000


def money(value):
    if value is None:
        return ""
    return f"{value / 100:.2f}".replace(".", ",")


def name_of(ref):
    if not isinstance(ref, dict):
        return ""
    return ref.get("name") or ""


def attr_map(row):
    found = {}
    for item in row.get("attributes") or []:
        value = item.get("value")
        if isinstance(value, dict):
            value = value.get("name") or value.get("href") or ""
        found[item.get("name") or ""] = "" if value is None else str(value)
    return found


def write_csv(path, fieldnames, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    print(f"{path.name}: {len(rows)}")


def export_counterparties():
    columns = MAPPING["spreadsheets"]["counterparties"]["columns"]
    rows = []
    for item in get_all("/entity/counterparty"):
        attrs = attr_map(item)
        extra = []
        if item.get("inn"):
            extra.append("ИНН " + item["inn"])
        if item.get("legalTitle"):
            extra.append(item["legalTitle"])
        rows.append({
            "Контрагенты": item.get("name") or "",
            "Дополнительная информация": "; ".join(extra),
            "Контакты бухгалтерии": item.get("phone") or item.get("email") or "",
            "Контакты с пометкой": item.get("email") or "",
            "Договор": attrs.get("Договор") or attrs.get("Ссылка на договор") or "",
            "Документы": attrs.get("Документы") or attrs.get("Ссылка") or "",
            "Инструкции по контрагентам": attrs.get("Инструкции") or "",
            "Исходящие письма": attrs.get("Исходящие письма") or "",
        })
    write_csv(OUT / "контрагенты.csv", columns, rows)


def export_products():
    columns = MAPPING["products_sheet_columns"]
    rows = []
    for item in get_all("/entity/product", "?expand=productFolder,uom"):
        folder = item.get("productFolder") or {}
        uom = item.get("uom") or {}
        sale = ((item.get("salePrices") or [{}])[0]).get("value")
        rows.append({
            "Наименование": item.get("name") or "",
            "Артикул": item.get("article") or item.get("code") or "",
            "Группа": folder.get("name") or "",
            "Цена": money(sale),
            "Единица": uom.get("name") or "",
            "Описание": item.get("description") or "",
            "Ссылка": (item.get("meta") or {}).get("uuidHref") or "",
            "Штрихкод": barcode(item),
        })
    write_csv(OUT / "товары.csv", columns, rows)


def barcode(item):
    codes = item.get("barcodes") or []
    if not codes:
        return ""
    first = codes[0]
    if isinstance(first, dict):
        return next(iter(first.values()), "")
    return str(first)


def export_cash(month):
    columns = MAPPING["spreadsheets"]["cash"]["flat_columns"]
    rows = []
    for kind, sign in (("paymentin", "in"), ("paymentout", "out")):
        for item in get_all(f"/entity/{kind}", "?expand=agent,organization"):
            moment = (item.get("moment") or "")[:10]
            if month and not moment.startswith(month):
                continue
            amount = money(item.get("sum"))
            rate = item.get("rate") if isinstance(item.get("rate"), dict) else {}
            rows.append({
                "Дата": moment,
                "Организация": name_of(item.get("organization")),
                "Контрагент": name_of(item.get("agent")),
                "Поступление": amount if sign == "in" else "",
                "Списание": amount if sign == "out" else "",
                "Валюта": name_of(rate.get("currency")),
                "Документ": item.get("name") or item.get("description") or "",
                "Ссылка": (item.get("meta") or {}).get("uuidHref") or "",
            })
    rows.sort(key=lambda row: (row["Дата"], row["Организация"], row["Контрагент"]))
    suffix = month or "все"
    write_csv(OUT / f"движение-{suffix}.csv", columns, rows)


def rub_currency():
    for item in get_all("/entity/currency"):
        if item.get("isoCode") == "RUB":
            return item["meta"]
    sys.exit("В МойСклад нет валюты RUB.")


def import_products(path):
    currency = rub_currency()
    with Path(path).open(encoding="utf-8-sig", newline="") as handle:
        source = list(csv.DictReader(handle))
    created = 0
    for row in source:
        name = (row.get("Наименование") or row.get("name") or "").strip()
        if not name:
            continue
        price = (row.get("Цена") or "0").replace(" ", "").replace(",", ".")
        try:
            kopecks = int(round(float(price) * 100))
        except ValueError:
            kopecks = 0
        body = {
            "name": name,
            "article": (row.get("Артикул") or "").strip(),
            "description": (row.get("Описание") or "").strip(),
            "salePrices": [{"value": kopecks, "currency": {"meta": currency}}],
        }
        code = (row.get("Штрихкод") or "").strip()
        if code:
            body["barcodes"] = [{"ean13": code}]
        request("POST", "/entity/product", body)
        created += 1
    print(f"создано товаров: {created}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--month", default="", help="Месяц движения денег, например 2026-09")
    parser.add_argument("--import-products", dest="products", default="")
    args = parser.parse_args()
    if args.products:
        import_products(args.products)
        return
    export_counterparties()
    export_products()
    export_cash(args.month)
    print("Готово. Файлы лежат в папке out.")


if __name__ == "__main__":
    main()
