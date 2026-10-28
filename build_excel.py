#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_excel.py — автоматическая сборка отчёта FAO Fishstat в один Excel-файл.

Два режима:
  1) derby  (боевой): подключается к Apache Derby через JDBC (jaydebeapi),
     выполняет table1.sql ... table10.sql и пишет результат на листы Excel.
  2) csv    (для проверки без БД): считает то же самое прямо по дампу CSV
     из папки CSV/ — логика идентична SQL-запросам.

Требования:
  pip install jaydebeapi openpyxl JPype1        # для режима derby
  pip install openpyxl                          # для обоих режимов

Запуск (Derby в embedded-режиме, путь к базе -Dderby.system.home или абсолютный):
  python build_excel.py --mode derby ^
      --db-path /path/to/db/FISHSTAT ^
      --jvm-lib /path/to/jdk/lib/server/jvm.dll ^
      --derby-jars "C:/derby/lib/derby.jar;C:/derby/lib/derbyshared.jar;C:/derby/lib/derbytools.jar" ^
      --out report.xlsx

Запуск (проверка без БД):
  python build_excel.py --mode csv --out report.xlsx

Важно: Derby embedded нельзя открывать одновременно из DBeaver/DataGrip —
перед запуском закройте соединения с базой в IDE.
"""

import argparse
import csv
import os
import re
import sys
from decimal import Decimal, InvalidOperation

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------------ config ---

SCHEMA = "FISHSTAT"
CSV_DIR = "CSV"  # переопределяется из аргументов
YEARS = [2020, 2021, 2022, 2023, 2024]

# коды 10 основных групп FAO: группы 1..9 -> '1'..'9', группа 10 -> 'A'
GROUP_CODES = ['1', '2', '3', '4', '5', '6', '7', '8', '9', 'A']
GROUP_NUMS = list(range(1, 11))

# таблица(N) -> номер группы; пустые группы (нет данных) пропускаются
GROUP_TABLES = {4: 1, 5: 2, 6: 3, 7: 4, 8: 5, 9: 6, 10: 7}

SHEET_NAMES = {
    1: "1_Топ стран",
    2: "2_Группы и товары",
    3: "3_Страны по годам",
    4: "4_Группа1 Fresh",
    5: "5_Группа2 Dried",
    6: "6_Группа3 Pelagic",
    7: "7_Группа4 Demersal",
    8: "8_Группа5 Crust",
    9: "9_Группа6 Mollusc",
    10: "10_Группа7 Other",
}

# русские заголовки колонок для оформления (по позициям)
RU_HEADERS = {
    1: ["Страна", "2020", "2021", "2022", "2023", "2024"],
    2: ["Группа", "Товар", "2020", "2021", "2022", "2023", "2024"],
    3: ["Страна", "Год", "ВСЕГО"] + [f"Группа {i}" for i in GROUP_NUMS],
}
RU_HEADERS_T4PLUS = ["Позиция (товар / страна)", "2020", "2021", "2022", "2023", "2024", "Итого 5 лет"]

# служебные колонки, которые НЕ пишутся в Excel (если вдруг вернулись из БД)
DROP_COLS = {"LVL", "GRPKEY", "ORDKEY"}

# ------------------------------------------------------------- styling ------

FILL_HEADER = PatternFill("solid", fgColor="1F4E79")
FILL_GROUP = PatternFill("solid", fgColor="DDEBF7")   # строки товаров (группирующие)
FONT_HEADER = Font(bold=True, color="FFFFFF", size=11)
FONT_ITEM = Font(bold=True, size=11)
FONT_COUNTRY = Font(size=10, color="404040")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NUM_FMT = "#,##0"


def style_sheet(ws, n_cols, item_col_idx=0, has_item_rows=False):
    """Оформление листа: жирная шапка, границы, числа, ширины колонок."""
    ws.freeze_panes = "A2"
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    for row in ws.iter_rows(min_row=2, max_col=n_cols):
        first = str(row[item_col_idx].value or "")
        is_item = has_item_rows and not first.startswith("     - ")
        for cell in row:
            cell.border = BORDER
            if isinstance(cell.value, (int, float, Decimal)):
                cell.number_format = NUM_FMT
                cell.alignment = Alignment(horizontal="right")
            if is_item:
                cell.font = FONT_ITEM
                cell.fill = FILL_GROUP
            elif has_item_rows:
                cell.font = FONT_COUNTRY
    # ширины
    for c in range(1, n_cols + 1):
        width = 12
        if c == 1:
            width = 55
        elif c == 2 and n_cols > 7:
            width = 10
        best = width
        for r in range(1, min(ws.max_row, 300) + 1):
            v = ws.cell(row=r, column=c).value
            if v is not None:
                best = max(best, min(len(str(v)) + 2, 60))
        ws.column_dimensions[get_column_letter(c)].width = best


# ------------------------------------------------------------ excel write ---

def add_sheet(wb, name, headers, rows, has_item_rows=False):
    ws = wb.create_sheet(title=name[:31])
    ws.append(headers)
    for r in rows:
        ws.append(list(r))
    style_sheet(ws, len(headers), item_col_idx=0, has_item_rows=has_item_rows)
    return ws


# --------------------------------------------------------------- derby mode --

def connect_derby(args):
    import jaydebeapi
    drv = args.derby_driver
    jars = ";".join(args.derby_jars.split(os.pathsep if ":" in args.derby_jars else ";"))
    props = {}
    if args.derby_home:
        props["derby.system.home"] = args.derby_home
    url = f"jdbc:derby:{args.db_path};create=false"
    return jaydebeapi.connect(drv, url, props, jars)


SQL_COMMENT_RE = re.compile(r"--[^\n]*")


def load_sql(path):
    with open(path, encoding="utf-8") as f:
        sql = f.read().strip().rstrip(";")
    return sql


def run_derby_mode(args, wb):
    conn = connect_derby(args)
    try:
        cur = conn.cursor()
        for tno in [1, 2, 3] + sorted(GROUP_TABLES):
            path = os.path.join(args.sql_dir, f"table{tno}.sql")
            if not os.path.exists(path):
                print(f"  пропущен: {path}")
                continue
            sql = load_sql(path)
            print(f"  выполнение table{tno}.sql ...")
            cur.execute(sql)
            cols = [d[0].strip('"') for d in cur.description]
            keep = [i for i, c in enumerate(cols) if c.upper() not in DROP_COLS]
            headers = [cols[i] for i in keep]
            # русские заголовки, если число колонок совпадает
            ru = RU_HEADERS.get(tno)
            if ru and len(ru) == len(headers):
                headers = ru
            rows = [[row[i] for i in keep] for row in cur.fetchall()]
            has_items = tno >= 4
            add_sheet(wb, SHEET_NAMES[tno], headers, rows, has_item_rows=has_items)
            print(f"    лист '{SHEET_NAMES[tno]}': {len(rows)} строк")
    finally:
        try:
            cur.close()
        except Exception:
            pass
        conn.close()


# ----------------------------------------------------------------- csv mode --

def to_dec(x):
    x = (x or "").strip().strip('"')
    if x == "":
        return Decimal(0)
    try:
        return Decimal(x)
    except InvalidOperation:
        return Decimal(0)


def read_csv(name):
    with open(os.path.join(CSV_DIR, name), encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def run_csv_mode(args, wb):
    commodities = read_csv("REF_COMMODITY.csv")
    majors = read_csv("REF_COMMODITY_MAJOR2.csv")
    rel = read_csv("REL_COMMODITY_MAJOR.csv")
    countries = read_csv("REF_COUNTRY.csv")
    trade = read_csv("TSD_TRADE_PP_QUANTITY.csv")

    cm_by_id = {c["ID"]: c for c in commodities}
    mj_by_id = {m["ID"]: m for m in majors}
    country_by_code = {c["UN_CODE"].zfill(3): c for c in countries}
    child_to_major = {}
    for r in rel:
        g = mj_by_id.get(r["PARENT_ID"])
        if g:
            child_to_major[r["CHILD_ID"]] = g["CODE"]  # '1'..'9','A'
    fao_group = {c["FAO_CODE"]: child_to_major.get(c["ID"], "") for c in commodities}

    YCOL = {y: f"VALUE_Y{y}" for y in YEARS}
    # подготовим факты: (country_code, commodity, {year: dec})
    facts = []
    for row in trade:
        vals = {y: to_dec(row.get(YCOL[y])) for y in YEARS}
        facts.append((row["COUNTRY"].strip(), row["COMMODITY"].strip(), vals))

    def tot(vals):
        return sum(vals.values(), Decimal(0))

    # ---- Таблица 1: топ-10 стран -------------------------------------------
    by_country = {}
    for cc, cm, vals in facts:
        d = by_country.setdefault(cc, {y: Decimal(0) for y in YEARS})
        for y in YEARS:
            d[y] += vals[y]
    ranked = sorted(by_country.items(), key=lambda kv: -sum(kv[1].values(), Decimal(0)))[:10]
    rows1 = [[country_by_code.get(cc, {}).get("NAME_EN", cc)] +
             [float(d[y]) for y in YEARS] for cc, d in ranked]
    add_sheet(wb, SHEET_NAMES[1], RU_HEADERS[1], rows1)
    print(f"    лист '{SHEET_NAMES[1]}': {len(rows1)} строк")

    # ---- Таблица 2: все товары внутри 10 групп ----------------------------
    by_com = {}
    for cc, cm, vals in facts:
        d = by_com.setdefault(cm, {y: Decimal(0) for y in YEARS})
        for y in YEARS:
            d[y] += vals[y]

    def gsort(code):
        return int(code) if code.isdigit() else 99

    rows2 = []
    for gi, gcode in enumerate(GROUP_CODES, start=1):
        items = []
        for c in commodities:
            if fao_group.get(c["FAO_CODE"]) != gcode:
                continue
            vals = by_com.get(c["FAO_CODE"], {y: Decimal(0) for y in YEARS})
            items.append((c["NAME_EN"], vals))
        items.sort(key=lambda it: -it[1][2024])
        for name, vals in items:
            rows2.append([f"{gi}. {majors_name(gcode)}", name] + [float(vals[y]) for y in YEARS])
    add_sheet(wb, SHEET_NAMES[2], RU_HEADERS[2], rows2)
    print(f"    лист '{SHEET_NAMES[2]}': {len(rows2)} строк")

    # ---- Таблица 3: страны x годы x группы --------------------------------
    gidx = {g: i for i, g in enumerate(GROUP_CODES)}
    by_cy = {}
    for cc, cm, vals in facts:
        g = fao_group.get(cm, "")
        for y in YEARS:
            key = (cc, y)
            d = by_cy.setdefault(key, [Decimal(0)] * 10)
            if g in gidx:
                d[gidx[g]] += vals[y]
    rows3 = []
    names_sorted = sorted({cc for cc, _ in by_cy},
                          key=lambda k: country_by_code.get(k, {}).get("NAME_EN", k))
    for cc in names_sorted:
        for y in YEARS:
            arr = by_cy.get((cc, y), [Decimal(0)] * 10)
            total = sum(arr, Decimal(0))
            rows3.append([country_by_code.get(cc, {}).get("NAME_EN", cc), y, float(total)] +
                         [float(v) for v in arr])
    add_sheet(wb, SHEET_NAMES[3], RU_HEADERS[3], rows3)
    print(f"    лист '{SHEET_NAMES[3]}': {len(rows3)} строк")

    # ---- Таблицы 4..10: товар -> страны ------------------------------------
    by_cm_country = {}
    for cc, cm, vals in facts:
        d = by_cm_country.setdefault((cm, cc), {y: Decimal(0) for y in YEARS})
        for y in YEARS:
            d[y] += vals[y]

    for tno, gnum in sorted(GROUP_TABLES.items()):
        gcode = GROUP_CODES[gnum - 1]
        rows = []
        comm_in_group = [c for c in commodities if fao_group.get(c["FAO_CODE"]) == gcode]
        prod = []
        for c in comm_in_group:
            vals = by_com.get(c["FAO_CODE"], {y: Decimal(0) for y in YEARS})
            if tot(vals) > 0:
                prod.append((c, vals))
        prod.sort(key=lambda it: -tot(it[1]))
        for c, vals in prod:
            rows.append([c["NAME_EN"]] + [float(vals[y]) for y in YEARS] + [float(tot(vals))])
            cs = [(cc, v) for (cm, cc), v in by_cm_country.items()
                  if cm == c["FAO_CODE"] and tot(v) > 0]
            cs.sort(key=lambda it: -tot(it[1]))
            for cc, v in cs:
                nm = country_by_code.get(cc.zfill(3), {}).get("NAME_EN", cc)
                rows.append([f"     - {nm}"] + [float(v[y]) for y in YEARS] + [float(tot(v))])
        add_sheet(wb, SHEET_NAMES[tno], RU_HEADERS_T4PLUS, rows, has_item_rows=True)
        print(f"    лист '{SHEET_NAMES[tno]}': {len(rows)} строк")


MAJOR_NAME_BY_CODE = {}


def majors_name(gcode):
    return MAJOR_NAME_BY_CODE.get(gcode, gcode)


# ------------------------------------------------------------------- main ---

def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--mode", choices=["derby", "csv"], default="csv")
    ap.add_argument("--out", default="report.xlsx")
    ap.add_argument("--sql-dir", default=os.path.dirname(os.path.abspath(__file__)))
    ap.add_argument("--csv-dir", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "CSV"))
    # параметры Derby/JDBC
    ap.add_argument("--db-path", help="путь к базе Derby (например .../FISHSTAT)")
    ap.add_argument("--derby-home", help="derby.system.home (родительский каталог базы)")
    ap.add_argument("--derby-jars", help="пути к derby.jar;derbyshared.jar;derbytools.jar")
    ap.add_argument("--derby-driver", default="org.apache.derby.jdbc.EmbeddedDriver")
    ap.add_argument("--jvm-lib", help="путь к jvm.dll/jlibvm.so для JPype")
    args = ap.parse_args()

    if args.mode == "derby" and args.jvm_lib:
        os.environ.setdefault("JPype_JVM_PATH", args.jvm_lib)

    wb = Workbook()
    wb.remove(wb.active)

    if args.mode == "csv":
        global MAJOR_NAME_BY_CODE, CSV_DIR
        CSV_DIR = args.csv_dir
        for m in read_csv("REF_COMMODITY_MAJOR2.csv"):
            MAJOR_NAME_BY_CODE[m["CODE"]] = m["NAME_EN"]
        print("[csv-режим] расчёт по дампу CSV ...")
        run_csv_mode(args, wb)
    else:
        print("[derby-режим] подключение к Apache Derby ...")
        run_derby_mode(args, wb)

    wb.save(args.out)
    print(f"Готово: {os.path.abspath(args.out)}")


if __name__ == "__main__":
    main()
