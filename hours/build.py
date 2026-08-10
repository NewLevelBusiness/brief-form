#!/usr/bin/env python3
"""Сборка дашборда загрузки команды из выгрузок Google-таблиц.

Вход:  hours.xlsx — «Учет времени по проектам», факт по вкладкам сотрудников
       plan.xlsx  — «НЛБ План часов для фин модели», план часов (необязателен)
Выход: data.json      — агрегированный куб (сотрудник × проект × месяц × категория × макрогруппа)
       dashboard.html — самодостаточный дашборд (template.html + вшитый data.json)

Ни один из выходных файлов не коммитится: репозиторий публичный, а данные внутренние.
См. README.md.

Запуск:  python3 build.py [путь_к_факту] [путь_к_плану]
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
PLAN_SRC = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "plan.xlsx")
# Ручные поправки к сопоставлению «ФИО из плана → вкладка в факте».
# Значение null означает «в факте этого человека нет, сравнивать не с чем».
MAPPING_OVERRIDE = os.path.join(HERE, "mapping.json")

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


def read_facts(wb, projects, only_person=None):
    """Разбирает вкладки сотрудников. only_person ограничивает разбор одной
    вкладкой — нужно для персональных файлов, где рядом лежат чужие копии."""
    facts, issues = [], []
    for ws in wb.worksheets:
        person = ws.title
        if person in SERVICE_SHEETS:
            continue
        if only_person is not None and person != only_person:
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


# План ведётся по «Фамилия Имя», факт — по вкладкам с уменьшительными именами.
# Таблица общая, не про конкретных людей: сводит полное имя к обиходному.
DIMINUTIVES = {
    "александр": "саша", "александра": "саша", "алексей": "лёша", "анастасия": "настя",
    "андрей": "андрей", "анна": "аня", "антон": "антон", "василий": "вася",
    "виктория": "вика", "владимир": "вова", "георгий": "гоша", "дарья": "даша",
    "денис": "денис", "дмитрий": "дима", "евгений": "женя", "евгения": "женя",
    "екатерина": "катя", "елена": "лена", "иван": "ваня", "илья": "илья",
    "ирина": "ира", "константин": "костя", "ксения": "ксюша", "мария": "маша",
    "михаил": "миша", "наталия": "наташа", "наталья": "наташа", "николай": "коля",
    "ольга": "оля", "павел": "паша", "пётр": "петя", "петр": "петя",
    "роман": "рома", "светлана": "света", "сергей": "сергей", "татьяна": "таня",
    "юлия": "юля", "юрий": "юра", "яна": "яна",
}

# Строки плана, которые описывают не человека, а итог или календарь
PLAN_SERVICE_ROWS = {"по произв календарю", "часы штатных", "часы внештатных"}
PLAN_SUBTOTALS = {"часы на проекты", "часы на нлб", "часы на НЛБ"}
MONTHS_RU = ["январь", "февраль", "март", "апрель", "май", "июнь",
             "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"]


def match_people(plan_names, fact_names):
    """Сопоставляет ФИО из плана со вкладками факта по обиходному имени.

    Неоднозначности не разрешает: если на одну вкладку претендуют двое (два
    Александра — это два разных Саши), оба помечаются как несопоставленные,
    чтобы сравнение план/факт не приписало часы не тому человеку.
    """
    fact_by_key = {norm(f).lower(): f for f in fact_names}
    # Персональные файлы дают человека полным именем, вкладки основного — коротким.
    # Порядок слов в ФИО в двух источниках разный, поэтому сравниваем множествами.
    fact_by_tokens = {frozenset(norm(f).lower().split()): f for f in fact_names}
    claims = defaultdict(list)
    for full in plan_names:
        parts = norm(full).split()
        exact = fact_by_tokens.get(frozenset(p.lower() for p in parts))
        if exact:
            claims[exact].append(full)
            continue
        for token in parts[1:] or parts:  # обычно «Фамилия Имя», но не всегда
            low = token.lower()
            hit = fact_by_key.get(DIMINUTIVES.get(low, low)) or fact_by_key.get(low)
            if hit:
                claims[hit].append(full)
                break
    mapping, ambiguous = {}, {}
    for fact_name, contenders in claims.items():
        if len(contenders) == 1:
            mapping[contenders[0]] = fact_name
        else:
            ambiguous[fact_name] = sorted(contenders)
    return mapping, ambiguous


def pick_plan_sheet(wb):
    """Файл плана хранит историю версий отдельными листами. Берём самую полную
    из тех, что названы планом: черновики лежат на листах вида «Лист6»."""
    best, best_key = None, None
    for ws in wb.worksheets:
        if "план" not in ws.title.lower():
            continue
        filled, hours = set(), 0.0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or norm(row[0]).lower() != "часы":
                continue
            if norm(row[1]).lower() in PLAN_SERVICE_ROWS or norm(row[2]).lower() in PLAN_SUBTOTALS:
                continue
            for j in range(3, 15):
                try:
                    v = float(row[j])
                except (TypeError, ValueError, IndexError):
                    continue
                if v:
                    filled.add(j)
                    hours += v
        key = (len(filled), hours)
        if filled and (best_key is None or key > best_key):
            best, best_key = ws, key
    return best


def read_plan(path, year_hint):
    """План часов: сотрудник × проект × месяц, плюс норма производственного календаря."""
    if not os.path.exists(path):
        return None
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = pick_plan_sheet(wb)
    if ws is None:
        return None
    # Год берём из названия листа: колонки подписаны только месяцами
    found = re.search(r"\b(20\d{2})\b", ws.title)
    year_hint = found.group(1) if found else year_hint
    cells, calendar, names = defaultdict(float), {}, set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or norm(row[0]).lower() != "часы":
            continue
        who, project = norm(row[1]), norm(row[2])
        for j in range(3, 15):
            try:
                value = float(row[j])
            except (TypeError, ValueError, IndexError):
                continue
            month = "%s-%02d" % (year_hint, j - 2)
            if who.lower() == "по произв календарю":
                calendar[month] = value
            elif who.lower() in PLAN_SERVICE_ROWS or project.lower() in PLAN_SUBTOTALS:
                continue
            elif value:
                cells[(who, project, month)] += value
                names.add(who)
    return {"sheet": ws.title, "cells": cells, "calendar": calendar, "names": names}


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


def row_key(fact):
    """Ключ строки для поиска дублей между файлами. Тип работ входит в ключ: без
    него двое, отметившие в один день по часу на одном проекте, выглядят как дубль."""
    return (fact["date"], round(fact["hours"], 2), fact["project"], fact["work_type"])


# Если персональный файл содержит почти все строки одноимённой вкладки основного
# файла, значит вкладка — устаревший огрызок того же учёта, и её надо выбросить.
SUPERSEDE_THRESHOLD = 0.9
# Ниже этой доли пересечение — совпадение независимых записей, а не дубль.
COINCIDENCE_RATIO = 0.2
COINCIDENCE_ROWS = 20


def read_extra_facts(projects, main_facts):
    """Часть команды ведёт учёт в собственных файлах — ссылки на них лежат в листе
    «справочники» в колонке «Ссылки на учет рабочих часов».

    Имя человека берётся из имени файла, а не из названия вкладки: у двух разных
    сотрудников вкладки могут называться одинаково («Саша» у обоих Александров),
    и по вкладке их не различить.

    Возвращает (строки, замечания, что_добавлено, каких_людей_убрать_из_основного).
    """
    folder = os.path.join(HERE, "extra")
    if not os.path.isdir(folder):
        return [], [], [], set()

    main_by_person = defaultdict(set)
    for fact in main_facts:
        main_by_person[fact["person"]].add(row_key(fact))

    facts, issues, added, superseded = [], [], [], set()
    for filename in sorted(os.listdir(folder)):
        if not filename.endswith(".xlsx") or filename.startswith("~"):
            continue
        person = os.path.splitext(filename)[0]
        wb = openpyxl.load_workbook(os.path.join(folder, filename), data_only=True)
        rows, problems = read_facts(wb, projects)
        # в персональных файлах рядом лежат рабочие листы без учёта часов — это норма
        issues += [f"{filename}: {p}" for p in problems if "не найдена шапка" not in p]
        for fact in rows:
            fact["person"] = person
        keys = {row_key(f) for f in rows}

        # с кем из основного файла этот человек пересекается
        for main_person, main_keys in main_by_person.items():
            if main_person in superseded or not main_keys:
                continue
            shared = len(main_keys & keys)
            if not shared:
                continue
            ratio = shared / len(main_keys)
            if ratio >= SUPERSEDE_THRESHOLD:
                superseded.add(main_person)
                added.append((filename, person, None,
                              f"заменяет вкладку «{main_person}» основного файла "
                              f"({len(main_keys)} строк, совпадение {ratio * 100:.0f}%)"))
            elif ratio >= COINCIDENCE_RATIO and shared >= COINCIDENCE_ROWS:
                issues.append(
                    f"{filename}: {shared} строк совпадают с вкладкой «{main_person}» "
                    f"основного файла ({ratio * 100:.0f}% её объёма) — часы задвоятся, "
                    f"проверьте, один ли это человек")

        facts += rows
        added.append((filename, person, len(rows), sum(f["hours"] for f in rows)))
    return facts, issues, added, superseded


def main():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    projects = read_projects(wb)
    facts, issues = read_facts(wb, projects)
    if not facts:
        sys.exit("Не удалось прочитать ни одной строки факта")

    extra_facts, extra_issues, extra_added, superseded = read_extra_facts(projects, facts)
    if superseded:
        facts = [f for f in facts if f["person"] not in superseded]
    facts += extra_facts
    issues += extra_issues

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

    # ── План ────────────────────────────────────────────────────────────────
    plan_block = None
    plan = read_plan(PLAN_SRC, max(f["date"].year for f in facts))
    if plan:
        mapping, ambiguous = match_people(plan["names"], people)
        if os.path.exists(MAPPING_OVERRIDE):
            with open(MAPPING_OVERRIDE, encoding="utf-8") as fh:
                for full, fact_name in json.load(fh).items():
                    ambiguous.pop(fact_name, None)
                    if fact_name:
                        mapping[full] = fact_name
                    else:
                        mapping.pop(full, None)

        # Проекты плана нужно посадить на те же индексы, что и проекты факта
        plan_proj_idx = dict(proj_idx)
        plan_names_out = list(project_names)
        plan_cells = []
        for (who, project, month), hours in sorted(plan["cells"].items()):
            if who not in mapping or month not in m_idx:
                continue
            display = projects.get(key(project), {}).get("display", norm(project))
            if display not in plan_proj_idx:
                plan_proj_idx[display] = len(plan_names_out)
                plan_names_out.append(display)
            plan_cells.append([p_idx[mapping[who]], plan_proj_idx[display],
                               m_idx[month], round(hours, 2)])

        unmatched = sorted(n for n in plan["names"] if n not in mapping)
        plan_block = {
            "sheet": plan["sheet"],
            "cells": plan_cells,
            "calendar": {m: v for m, v in plan["calendar"].items() if m in m_idx},
            "mapping": {v: k for k, v in sorted(mapping.items())},
            "ambiguous": ambiguous,
            # часы, запланированные тем, кого не с чем сравнивать
            "unmatched_hours": round(sum(
                h for (who, pr, m), h in plan["cells"].items()
                if who not in mapping and m in m_idx), 1),
            "unmatched": unmatched,
            "months": sorted({m for (_, _, m) in plan["cells"] if m in m_idx}),
        }
        # проекты могли пополниться теми, что есть только в плане
        project_names = plan_names_out
        proj_idx = plan_proj_idx

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
        "plan": plan_block,
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

    if extra_added:
        print("подключены персональные файлы учёта:")
        for filename, person, rows_count, payload in extra_added:
            if rows_count is None:
                print(f"   {filename}: {payload}")
            else:
                print(f"   {filename} → «{person}»: {rows_count} строк, {payload:.0f} ч")
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

    if plan_block:
        print(f"\nплан: лист «{plan_block['sheet']}», "
              f"{len(plan_block['months'])} мес., {len(plan_block['cells'])} ячеек")
        print("  сопоставлены:", ", ".join(
            f"{v} → {k}" for k, v in sorted(plan_block["mapping"].items())))
        if plan_block["unmatched"]:
            print(f"  без пары в факте ({plan_block['unmatched_hours']:.0f} ч плана): "
                  + ", ".join(plan_block["unmatched"]))
        for fact_name, contenders in plan_block["ambiguous"].items():
            print(f"  НЕОДНОЗНАЧНО: на вкладку «{fact_name}» претендуют "
                  + " и ".join(contenders) + " — уточните в mapping.json")
    else:
        print(f"\nплан не подключён: файла {PLAN_SRC} нет")


if __name__ == "__main__":
    main()
