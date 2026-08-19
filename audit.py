#!/usr/bin/env python3
"""
audit.py — универсальный аудит файла расходов (CSV или Excel).

Использование:
    python audit.py <путь_к_файлу> [--report <путь_к_отчёту>]

Скрипт сам определяет, какие столбцы содержат дату, категорию,
описание и сумму — сначала по названию заголовка, а если это не
удалось, по содержимому столбцов. Поэтому подходит для любого файла
расходов, а не только для конкретного набора колонок.
"""

import argparse
import csv
import datetime
import io
import statistics
import sys
from collections import defaultdict
from decimal import Decimal, InvalidOperation

try:
    from dateutil import parser as dateutil_parser
    HAVE_DATEUTIL = True
except ImportError:
    HAVE_DATEUTIL = False


DATE_KEYWORDS = [
    "дата", "date", "день", "число", "период", "period", "месяц", "month", "timestamp", "created",
]
AMOUNT_KEYWORDS = [
    "сумма", "amount", "итого", "total", "стоимость", "price", "cost", "value", "sum", "руб",
]
CATEGORY_KEYWORDS = [
    "категория", "category", "тип", "type", "статья", "group", "группа", "класс", "раздел",
]
DESCRIPTION_KEYWORDS = [
    "описание", "description", "комментар", "назначение", "note", "деталь", "purpose", "наименован",
]


# ---------------------------------------------------------------- чтение файла

def read_rows(filepath):
    """Возвращает (headers, rows), rows — список списков строковых значений."""
    lower = filepath.lower()
    if lower.endswith((".xlsx", ".xlsm")):
        return read_excel(filepath)
    if lower.endswith(".xls"):
        raise SystemExit(
            "Формат .xls (старый Excel) не поддерживается напрямую.\n"
            "Пересохраните файл в .xlsx или .csv и попробуйте снова."
        )
    return read_csv(filepath)


def read_csv(filepath):
    raw = None
    for enc in ("utf-8-sig", "utf-8", "cp1251", "latin-1"):
        try:
            with open(filepath, encoding=enc) as f:
                raw = f.read()
            break
        except (UnicodeDecodeError, LookupError):
            continue
    if raw is None:
        raise SystemExit(f"Не удалось прочитать файл {filepath} — неизвестная кодировка.")

    sample = raw[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
    except csv.Error:
        dialect = csv.excel

    reader = csv.reader(io.StringIO(raw), dialect)
    all_rows = [row for row in reader if any(cell.strip() for cell in row)]
    if not all_rows:
        raise SystemExit(f"Файл {filepath} пуст.")

    return all_rows[0], all_rows[1:]


def read_excel(filepath):
    try:
        import openpyxl
    except ImportError:
        raise SystemExit(
            "Для чтения Excel-файлов нужна библиотека openpyxl.\n"
            "Установите её: pip install openpyxl"
        )
    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    ws = wb.active
    all_rows = []
    for row in ws.iter_rows(values_only=True):
        if any(c is not None and str(c).strip() for c in row):
            all_rows.append(["" if c is None else str(c) for c in row])
    if not all_rows:
        raise SystemExit(f"Файл {filepath} пуст.")

    return all_rows[0], all_rows[1:]


# ------------------------------------------------------- разбор значений

def try_parse_amount(value):
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    s = s.replace("\xa0", "").replace(" ", "")
    s = s.replace("₽", "").replace("руб.", "").replace("руб", "").replace("RUB", "").replace("$", "")
    if "," in s and "." in s:
        s = s.replace(",", "")
    elif "," in s and "." not in s:
        s = s.replace(",", ".")
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def try_parse_date(value):
    s = str(value).strip()
    if not s:
        return None
    if HAVE_DATEUTIL:
        try:
            return dateutil_parser.parse(s, dayfirst=True, fuzzy=False)
        except (ValueError, OverflowError, TypeError):
            return None
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d"):
        try:
            return datetime.datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def normalize(s):
    return (s or "").strip().lower()


def header_matches(header, keywords):
    h = normalize(header)
    return any(k in h for k in keywords)


# ------------------------------------------------------- автоопределение столбцов

def detect_columns(headers, rows, sample_size=300):
    n_cols = len(headers)
    col_values = [[] for _ in range(n_cols)]
    for row in rows:
        for i in range(n_cols):
            col_values[i].append(row[i] if i < len(row) else "")

    date_frac, amount_frac, distinct_ratio = [], [], []
    for i in range(n_cols):
        vals = [v for v in col_values[i][:sample_size] if str(v).strip()]
        total = len(vals) or 1
        d_ok = sum(1 for v in vals if try_parse_date(v) is not None)
        a_ok = sum(1 for v in vals if try_parse_amount(v) is not None)
        distinct = len(set(str(v).strip() for v in vals))
        date_frac.append(d_ok / total)
        amount_frac.append(a_ok / total)
        distinct_ratio.append(distinct / total)

    used = set()

    def pick_by_header(keywords):
        for i, h in enumerate(headers):
            if i not in used and header_matches(h, keywords):
                return i
        return None

    def pick_by_content(frac, min_frac):
        candidates = [i for i in range(n_cols) if i not in used and frac[i] >= min_frac]
        if not candidates:
            return None
        return max(candidates, key=lambda i: frac[i])

    date_col = pick_by_header(DATE_KEYWORDS)
    if date_col is None:
        date_col = pick_by_content(date_frac, 0.7)
    if date_col is not None:
        used.add(date_col)

    amount_col = pick_by_header(AMOUNT_KEYWORDS)
    if amount_col is None:
        amount_col = pick_by_content(amount_frac, 0.7)
    if amount_col is not None:
        used.add(amount_col)

    if amount_col is None:
        raise SystemExit(
            "Не удалось автоматически определить столбец с суммой.\n"
            f"Заголовки файла: {headers}\n"
            "Переименуйте нужный столбец в 'Сумма'/'Amount' или проверьте формат чисел в нём."
        )

    category_col = pick_by_header(CATEGORY_KEYWORDS)
    if category_col is None:
        cat_candidates = [
            i for i in range(n_cols)
            if i not in used and amount_frac[i] < 0.5 and date_frac[i] < 0.5
            and 0 < distinct_ratio[i] < 0.9
        ]
        if cat_candidates:
            category_col = min(cat_candidates, key=lambda i: distinct_ratio[i])
    if category_col is not None:
        used.add(category_col)

    description_col = pick_by_header(DESCRIPTION_KEYWORDS)
    if description_col is None:
        desc_candidates = [
            i for i in range(n_cols)
            if i not in used and amount_frac[i] < 0.5 and date_frac[i] < 0.5
        ]
        if desc_candidates:
            description_col = max(desc_candidates, key=lambda i: distinct_ratio[i])
    if description_col is not None:
        used.add(description_col)

    return {"date": date_col, "amount": amount_col, "category": category_col, "description": description_col}


def get(row, idx):
    if idx is None:
        return ""
    return row[idx] if idx < len(row) else ""


# ------------------------------------------------------------------- отчёт

def main():
    parser = argparse.ArgumentParser(description="Аудит файла расходов (CSV или Excel).")
    parser.add_argument("filepath", help="Путь к CSV или Excel файлу с расходами")
    parser.add_argument("--report", default="report.txt", help="Куда сохранить отчёт (по умолчанию report.txt)")
    args = parser.parse_args()

    headers, rows = read_rows(args.filepath)
    cols = detect_columns(headers, rows)

    lines = []

    def out(s=""):
        lines.append(s)

    out("=== АУДИТ РАСХОДОВ ===")
    out(f"Файл: {args.filepath}")
    out(f"Всего записей: {len(rows)}")
    out()
    out("Определённые столбцы:")
    for key in ("date", "category", "description", "amount"):
        idx = cols[key]
        out(f"   {key}: {headers[idx] if idx is not None else '(не найден)'}")
    out()

    parsed = []
    invalid = []
    for i, row in enumerate(rows):
        amt = try_parse_amount(get(row, cols["amount"]))
        if amt is None:
            invalid.append((i, row))
            continue
        parsed.append({
            "idx": i,
            "date": get(row, cols["date"]),
            "category": (get(row, cols["category"]) or "").strip() or "(без категории)",
            "description": get(row, cols["description"]),
            "amount": amt,
        })

    if invalid:
        out(f"Записей с некорректной/пустой суммой (исключены из расчётов): {len(invalid)}")
        for i, row in invalid[:20]:
            out(f"   строка {i + 2}: {row}")
        if len(invalid) > 20:
            out(f"   ... и ещё {len(invalid) - 20}")
        out()

    if not parsed:
        raise SystemExit("Не удалось разобрать ни одной строки с суммой — проверьте файл.")

    # 1. Общая сумма
    total = sum((p["amount"] for p in parsed), Decimal(0))
    out("1. ОБЩАЯ СУММА РАСХОДОВ")
    out(f"   {total}")
    out()

    # 2. По категориям
    cat_sums = defaultdict(lambda: Decimal(0))
    cat_counts = defaultdict(int)
    for p in parsed:
        cat_sums[p["category"]] += p["amount"]
        cat_counts[p["category"]] += 1

    out("2. СУММА ПО КАТЕГОРИЯМ (по убыванию)")
    for cat, s in sorted(cat_sums.items(), key=lambda x: -x[1]):
        out(f"   {cat}: {s} ({cat_counts[cat]} записей)")
    out()

    # 3. Топ-5
    out("3. ТОП-5 САМЫХ КРУПНЫХ ТРАТ")
    for p in sorted(parsed, key=lambda p: -p["amount"])[:5]:
        out(f"   строка {p['idx'] + 2} | {p['date']} | {p['category']} | {p['description']} | {p['amount']}")
    out()

    # 4. Дубликаты
    out("4. ДУБЛИКАТЫ (одинаковые дата, описание и сумма)")
    seen = defaultdict(list)
    for p in parsed:
        key = (p["date"], p["description"], str(p["amount"]))
        seen[key].append(p["idx"])
    dupes = {k: v for k, v in seen.items() if len(v) > 1}
    if not dupes:
        out("   Дубликатов не найдено")
    else:
        total_dupe_rows = sum(len(v) for v in dupes.values())
        out(f"   Найдено групп дубликатов: {len(dupes)}, всего строк: {total_dupe_rows}")
        for (date, desc, amt), idxs in dupes.items():
            row_numbers = ", ".join(str(i + 2) for i in idxs)
            out(f"   {date} | {desc} | {amt} -> строки: {row_numbers}")
    out()

    # 5. Аномалии
    out("5. АНОМАЛИИ (отклонение > 3 стандартных отклонений от среднего по категории)")
    cat_amounts = defaultdict(list)
    for p in parsed:
        cat_amounts[p["category"]].append(float(p["amount"]))

    anomalies = []
    for p in parsed:
        amts = cat_amounts[p["category"]]
        if len(amts) < 2:
            continue
        mean = statistics.mean(amts)
        stdev = statistics.stdev(amts)
        if stdev == 0:
            continue
        z = abs(float(p["amount"]) - mean) / stdev
        if z > 3:
            anomalies.append((z, p, mean, stdev))

    if not anomalies:
        out("   Аномалий не найдено")
    else:
        anomalies.sort(key=lambda x: -x[0])
        for z, p, mean, stdev in anomalies:
            out(
                f"   строка {p['idx'] + 2} | {p['date']} | {p['category']} | {p['description']} | {p['amount']} "
                f"(среднее по категории: {mean:.2f}, σ={stdev:.2f}, z={z:.2f})"
            )
    out()

    # 6. Отрицательные суммы
    out("6. СТРОКИ С ОТРИЦАТЕЛЬНЫМИ СУММАМИ")
    negatives = [p for p in parsed if p["amount"] < 0]
    if not negatives:
        out("   Не найдено")
    else:
        for p in negatives:
            out(f"   строка {p['idx'] + 2} | {p['date']} | {p['category']} | {p['description']} | {p['amount']}")

    report = "\n".join(lines)
    with open(args.report, "w", encoding="utf-8") as f:
        f.write(report + "\n")

    print(report)
    print(f"\nОтчёт сохранён в {args.report}")


if __name__ == "__main__":
    main()
