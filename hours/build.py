#!/usr/bin/env python3
"""Сборка дашборда загрузки команды из выгрузки Google-таблицы «Учет времени по проектам».

Вход:  hours.xlsx — экспорт таблицы в формате .xlsx
Выход: data.json      — агрегированный куб (сотрудник × проект × месяц × категория × макрогруппа)
       dashboard.html — самодостаточный дашборд (template.html + вшитый data.json)

Ни один из выходных файлов не коммитится: репозиторий публичный, а данные внутренние.
См. README.md.

Запуск:  python3 build.py [путь_к_xlsx]
"""

import csv
import datetime
import json
import os
import re
import sys
from collections import defaultdict

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "hours.xlsx")

# Вкладки таблицы: по одной на сотрудника + служебные
SERVICE_SHEETS = {"справочники", "Новый сравочник", "план_по_периодам"}

# Порядок категорий фиксирован — от него зависит порядок цветовых слотов в дашборде
CATEGORIES = ["Коммерческие", "Внутренние (NLB)", "Пресейл", "Отпуск"]

# Фактические типы работ (55 значений свободного списка) → макрогруппы.
# Правила упорядочены: побеждает первое совпадение.
MACRO_RULES = [
    (r"^Отпуск", "Отпуск"),
    (r"^Пресейл", "Пресейл"),
    (r"Развитие НЛБ|Анонимизация дашбордов для НЛБ|Описание кейса", "Внутреннее развитие"),
    (r"Синк ->|Встречи по проекту с командой НЛБ", "Внутренние синки"),
    (r"Администрирование проекта|Управление и администрирование|Трекинг задач|Подготовка презентаций",
     "Управление проектом"),
    (r"Системные встречи|Рабочие встречи|Встречи с командой заказчиков", "Встречи с заказчиком"),
    (r"Разработка мод|^Разработка$|Бэк дашборда|Разработка интерфейс|Разработка визуализ|"
     r"Тестирование|Поддержка и обновление|Публикация|Создание подключения|Разработка отчета|"
     r"Разработка прототипов|Разработка и адаптация интерфейсов|Итерационные доработки",
     "Разработка"),
    (r"Организация и настройка процесса получения данных|Настройка ролевой модели|"
     r"Выгрузка необходимых данных|Настройка инфраструктуры|Управление доступами",
     "ИТ-инфраструктура"),
    (r"Обучение пользователей|Передача проекта", "Обучение и передача"),
    (r"Анализ|Формирование ТЗ|Разработка ТЗ|Проектирование ТЗ|Разработка макетов|"
     r"Разработка целевого состояния|Разработка правил расчетов|Изучение исходных данных|"
     r"интервью|Снятие запрос|Формирование требований|Проверка качества|Выработка",
     "Аналитика и проектирование"),
]
MACRO_FALLBACK = "Прочее"

MACRO_ORDER = [
    "Аналитика и проектирование", "Разработка", "Встречи с заказчиком", "Управление проектом",
    "ИТ-инфраструктура", "Обучение и передача", "Пресейл", "Внутреннее развитие",
    "Внутренние синки", "Отпуск", "Прочее",
]


# Кириллические буквы, неотличимые на вид от латинских. В названиях проектов
# они смешаны: одно и то же название набрано то так, то эдак, и справочник
# перестаёт находиться. Ключ сводит оба написания к одному виду.
CONFUSABLES = str.maketrans("АВЕКМНОРСТУХасеорхуѕі", "ABEKMHOPCTYXaceopxysi")


def norm(s):
    return re.sub(r"\s+", " ", str(s)).strip() if s is not None else ""


def key(name):
    """Ключ сопоставления названий: регистр, пробелы и кириллица-двойники не мешают."""
    return norm(name).translate(CONFUSABLES).lower()


def macro_of(work_type):
    for pattern, group in MACRO_RULES:
        if re.search(pattern, work_type, re.IGNORECASE):
            return group
    return MACRO_FALLBACK


def parse_date(value):
    """Даты в таблице лежат тремя способами: datetime, серийный номер Excel и текст dd.mm.yyyy."""
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    if isinstance(value, (int, float)):
        return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(value))).date()
    if isinstance(value, str):
        text = value.strip()
        for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y"):
            try:
                return datetime.datetime.strptime(text, fmt).date()
            except ValueError:
                pass
        m = re.match(r"^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})", text)
        if m:
            return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    return None


def find_header(ws):
    """Шапка стоит на 1-й или 2-й строке, а у трёх человек слева добавлены колонки период."""
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=4, values_only=True), start=1):
        cells = [norm(c).lower() for c in row]
        if "дата" in cells and "к-во часов" in cells:
            return row_idx, cells
    return None, None


def read_projects(wb):
    """Справочник проектов: заказчик → статус / тип / РП. Ключ — нормализованное название."""
    ws = wb["справочники"]
    projects = {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        if not row or not row[1]:
            continue
        name = key(row[1])
        if name in projects:
            continue
        projects[name] = {
            "display": norm(row[1]),
            "type": norm(row[0]),
            "status": norm(row[2]),
            "budget_client": norm(row[3]),
            "rp": norm(row[4]),
        }
    return projects


def read_facts(wb, projects):
    facts, issues = [], []
    for ws in wb.worksheets:
        person = ws.title
        if person in SERVICE_SHEETS:
            continue
        header_row, cells = find_header(ws)
        if header_row is None:
            issues.append(f"{person}: не найдена шапка таблицы")
            continue
        idx = {key: cells.index(key) for key in ("дата", "к-во часов", "проект", "тип работ")
               if key in cells}
        for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
            def cell(key):
                i = idx.get(key)
                return row[i] if i is not None and i < len(row) else None

            date = parse_date(cell("дата"))
            if date is None:
                continue
            try:
                hours = float(cell("к-во часов"))
            except (TypeError, ValueError):
                issues.append(f"{person} {date}: не число в часах — {cell('к-во часов')!r}")
                continue
            project = norm(cell("проект"))
            work_type = norm(cell("тип работ"))
            if not project:
                issues.append(f"{person} {date}: пустой проект ({hours} ч)")
                continue
            facts.append({
                "person": person,
                "date": date,
                "hours": hours,
                "project": project,
                "work_type": work_type,
            })
    return facts, issues


def categorize(fact, projects):
    """Категория часа для расчёта загрузки."""
    meta = projects.get(key(fact["project"]), {})
    if macro_of(fact["work_type"]) == "Отпуск":
        return "Отпуск"
    if meta.get("status") == "подтвержденные":
        return "Коммерческие"
    if meta.get("type", "").lower() == "пресейл" or fact["work_type"].startswith("Пресейл"):
        return "Пресейл"
    return "Внутренние (NLB)"


def main():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    projects = read_projects(wb)
    facts, issues = read_facts(wb, projects)
    if not facts:
        sys.exit("Не удалось прочитать ни одной строки факта")

    # Одно и то же название, набранное вперемешку кириллицей и латиницей, должно
    # стать одним проектом. Каноническим считаем написание из справочника.
    for fact in facts:
        meta = projects.get(key(fact["project"]))
        fact["project"] = meta["display"] if meta else norm(fact["project"])

    people = sorted({f["person"] for f in facts})
    project_names = sorted({f["project"] for f in facts})
    months = sorted({f["date"].strftime("%Y-%m") for f in facts})
    macros = [m for m in MACRO_ORDER if any(macro_of(f["work_type"]) == m for f in facts)]

    p_idx = {name: i for i, name in enumerate(people)}
    proj_idx = {name: i for i, name in enumerate(project_names)}
    m_idx = {name: i for i, name in enumerate(months)}
    c_idx = {name: i for i, name in enumerate(CATEGORIES)}
    mac_idx = {name: i for i, name in enumerate(macros)}

    cube = defaultdict(float)
    for fact in facts:
        cell = (
            p_idx[fact["person"]],
            proj_idx[fact["project"]],
            m_idx[fact["date"].strftime("%Y-%m")],
            c_idx[categorize(fact, projects)],
            mac_idx[macro_of(fact["work_type"])],
        )
        cube[cell] += fact["hours"]

    unknown = [p for p in project_names if key(p) not in projects]

    data = {
        "meta": {
            "source": "Учет времени по проектам (Google Sheets)",
            "period": [min(f["date"] for f in facts).isoformat(),
                       max(f["date"] for f in facts).isoformat()],
            "rows": len(facts),
            "hours": round(sum(f["hours"] for f in facts), 1),
            "issues": issues,
            "unknown_projects": unknown,
        },
        "people": people,
        "months": months,
        "categories": CATEGORIES,
        "macros": macros,
        "projects": [
            {
                "name": name,
                "status": projects.get(key(name), {}).get("status", ""),
                "type": projects.get(key(name), {}).get("type", ""),
                "rp": projects.get(key(name), {}).get("rp", ""),
            }
            for name in project_names
        ],
        "cube": [[k[0], k[1], k[2], k[3], k[4], round(v, 2)] for k, v in sorted(cube.items())],
    }

    with open(os.path.join(HERE, "data.json"), "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, separators=(",", ":"))

    with open(os.path.join(HERE, "facts.csv"), "w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(["сотрудник", "дата", "часы", "проект", "тип работ", "категория", "макрогруппа"])
        for fact in sorted(facts, key=lambda f: (f["date"], f["person"])):
            writer.writerow([fact["person"], fact["date"].isoformat(), fact["hours"],
                             fact["project"], fact["work_type"],
                             categorize(fact, projects), macro_of(fact["work_type"])])

    with open(os.path.join(HERE, "template.html"), encoding="utf-8") as fh:
        template = fh.read()
    html = template.replace("/*__DATA__*/null", json.dumps(data, ensure_ascii=False))
    with open(os.path.join(HERE, "dashboard.html"), "w", encoding="utf-8") as fh:
        fh.write(html)

    print(f"строк факта: {len(facts)}  часов: {data['meta']['hours']}")
    print(f"период: {data['meta']['period'][0]} — {data['meta']['period'][1]}")
    print(f"сотрудников: {len(people)}  проектов: {len(project_names)}  месяцев: {len(months)}")
    print(f"ячеек куба: {len(data['cube'])}")
    if unknown:
        print(f"проектов нет в справочнике: {unknown}")
    if issues:
        print(f"проблемных строк: {len(issues)}")
        for issue in issues[:10]:
            print("   ", issue)


if __name__ == "__main__":
    main()
