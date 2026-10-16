"""Сборка «Сводки 2-х час ГПН-Ямал» из книги ПОНМ.

Каждый дневной лист источника ('01'..'31') превращается в лист 'дд.мм' готовой сводки.
"""
import logging
import os
import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string, get_column_letter

from .config import Config
from .errors import ConfigError, SvodkaError
from .manual import load_manual_values
from .safeio import check_source_file, read_json, write_json

log = logging.getLogger("svodka")

FIRST_COL = 4  # столбец D — начало часовых значений в сводке
SERVICE_SHEETS = ("data_sms", "смс")
OWN_NAME_PREFIX = "Сводка 2-х час ГПН-Ямал"

# (столбец источника, первая строка источника, строка в сводке, число столбцов источника)
# Столбец источника (n ячеек по вертикали) -> строка сводки, начиная с D (транспонирование).
# Для 2 столбцов второй попадает в строку ниже.
COPY_RULES = [
    ("P", 89, 10, 2), ("B", 50, 12, 1), ("BK", 31, 24, 2), ("BG", 31, 19, 1),
    ("BW", 69, 20, 1), ("AY", 89, 17, 1), ("BA", 89, 18, 1), ("AH", 89, 13, 2),
    ("X", 154, 31, 1), ("W", 154, 37, 1), ("M", 154, 40, 1), ("AE", 154, 42, 1),
    ("BZ", 89, 54, 1), ("CA", 89, 57, 1), ("CB", 89, 60, 1), ("I", 154, 69, 1),
    ("AD", 154, 77, 1), ("BX", 89, 80, 1), ("BY", 89, 83, 1), ("BV", 89, 86, 1),
    ("CM", 89, 44, 1), ("CN", 89, 46, 1), ("CO", 89, 48, 1),   # АСВ
    ("H", 154, 91, 1), ("AB", 89, 96, 1), ("AH", 154, 94, 1),
    ("BD", 89, 99, 1),    # температура
    ("CJ", 89, 100, 1),   # стабильный конденсат (кондиция)
    ("CK", 89, 101, 1),   # стабильный конденсат (некондиция)
    ("W", 89, 4, 2),
]
# Блок «Уловленная» внизу сводки: (ячейка суточного итога в источнике, строка «текущие сутки»,
# строка «с начала месяца»). Первая запись — УКПГ, вторая — ЦПС. «Текущие сутки» — ячейка D строки,
# «с начала месяца» — по столбцу на каждое число месяца (1 -> D, 2 -> E, ... 31 -> AH).
CAPTURE = [("BC101", 105, 111), ("BN101", 107, 112)]
# строка сводки -> номера столбцов источника, суммируются построчно (строка 154+i)
SUM_154 = {34: [21, 22], 72: [2, 3, 4, 5, 6, 10, 11, 14, 15]}
# нефтеналив (невозврат): строка сводки -> столбцы источника (строка 89+i)
SUM_89 = {15: ["AM", "AP", "AS"], 16: ["AO", "AR", "AU"]}

# Все строки сводки, куда fill_day() пишет 2-часовые значения. Нужны, чтобы очистить их
# в новом дне: лист-образец содержит числа по всем 12 интервалам, а fill_day() пишет
# только наступившие — остальное должно остаться пустым.
DATA_DST_ROWS = set()
for _col, _row0, _dst_row, _ncols in COPY_RULES:
    for _j in range(_ncols):
        DATA_DST_ROWS.add(_dst_row + _j)
DATA_DST_ROWS.update(SUM_154.keys())
DATA_DST_ROWS.update(SUM_89.keys())

MONTH_STEMS = [("янв", 1), ("фев", 2), ("мар", 3), ("апр", 4), ("мая", 5), ("май", 5),
               ("июн", 6), ("июл", 7), ("авг", 8), ("сен", 9), ("окт", 10), ("ноя", 11), ("дек", 12)]
EXCEL_EPOCH = datetime(1899, 12, 30)


def result_name(src_stem: str) -> str:
    return f"{OWN_NAME_PREFIX} — {src_stem}.xlsx"


@dataclass
class RunState:
    """Что накопилось за один запуск сборки."""
    written: dict = field(default_factory=dict)   # (лист, ячейка) -> записанное значение
    skipped: list = field(default_factory=list)   # ячейки внутри объединённых диапазонов
    empty: dict = field(default_factory=dict)     # (диапазон, строка) -> листы без значений
    partial: dict = field(default_factory=dict)   # то же, но часть значений пустая
    strings: dict = field(default_factory=dict)   # диапазоны, где числа были текстом


# ---------- чтение источника ----------
class _Cell:
    def __init__(self, value):
        self.value = value


class XlsbSheet:
    """Лист .xlsb (только чтение) с интерфейсом как у openpyxl."""

    def __init__(self, path, sheet_name):
        from pyxlsb import open_workbook
        self._d = {}
        with open_workbook(path) as wb:
            with wb.get_sheet(sheet_name) as sh:
                for row in sh.rows():
                    for c in row:
                        if c.v is not None:
                            self._d[(c.r + 1, c.c + 1)] = c.v

    def cell(self, row, column):
        return _Cell(self._d.get((row, column)))

    def __getitem__(self, addr):
        letters = "".join(ch for ch in addr if ch.isalpha())
        digits = "".join(ch for ch in addr if ch.isdigit())
        return self.cell(int(digits), column_index_from_string(letters))


class XlsbBook:
    def __init__(self, path):
        from pyxlsb import open_workbook
        self.path = path
        with open_workbook(path) as wb:
            self.sheetnames = list(wb.sheets)

    def __getitem__(self, name):
        return XlsbSheet(self.path, name)


def open_source(path: str):
    if path.lower().endswith(".xlsb"):
        return XlsbBook(path)
    return load_workbook(path, data_only=True, keep_links=False)  # значения, а не формулы


# ---------- вспомогательное ----------
def copy_count(now: datetime) -> int:
    """Индекс последнего доступного 2-часового интервала текущего дня (0..11)."""
    a = now.hour // 2 - 2
    if a == -1:
        a = 11
    elif a == -2:
        a = 10
    return a


def num(v):
    if isinstance(v, (int, float)):
        return v
    try:
        return float(str(v).replace(",", "."))
    except (TypeError, ValueError):
        return 0


def col_values(ws, col, row0, n):
    return [ws[f"{col}{row0 + i}"].value for i in range(n)]


def to_num(v):
    """'1 662,59' (текст) -> 1662.59; остальное без изменений."""
    if isinstance(v, str):
        t = v.replace(" ", "").replace(" ", "").replace(" ", "").replace(",", ".")
        try:
            return float(t)
        except ValueError:
            return v
    return v


def set_cell(ws, coord, value, run: RunState, literal=False):
    """Записывает значение. literal=True — значение из чужого файла: текст никогда не
    превращается в формулу (даже если начинается с '=')."""
    cell = ws[coord]
    if isinstance(cell, MergedCell):
        run.skipped.append((ws.title, coord))
        return None
    cell.value = value
    if literal and isinstance(value, str):
        cell.data_type = "s"
    return cell


def put_row(ws, row, values, run: RunState, start_col=FIRST_COL):
    for i, v in enumerate(values):
        v = to_num(v)
        coord = ws.cell(row=row, column=start_col + i).coordinate
        if set_cell(ws, coord, v, run, literal=True) is not None:
            run.written[(ws.title, coord)] = v


def month_from_filename(path):
    name = os.path.basename(path).lower()
    for stem, m in MONTH_STEMS:
        if stem in name:
            return m
    return None


def find_sheet_date(ws, max_row=6, max_col=3):
    """Настоящая дата листа-источника: сканирует первые строки столбцов A..C
    (у ГПН-Ямал дата лежит в A3 как серийное число Excel)."""
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            v = ws.cell(r, c).value
            if isinstance(v, datetime):
                return v.date()
            if isinstance(v, date):
                return v
            if isinstance(v, (int, float)) and 40000 <= v <= 60000:  # ~2009..2064
                return (EXCEL_EPOCH + timedelta(days=float(v))).date()
    return None


def pick_template(wb, wanted):
    """Лист-образец для новых дней: из настроек, 'sample' или последний 'дд.мм'."""
    if wanted:
        if wanted not in wb.sheetnames:
            raise SvodkaError(f"Нет листа-образца '{wanted}'. Есть: {wb.sheetnames}")
        return wanted
    if "sample" in wb.sheetnames:
        return "sample"
    days = [n for n in wb.sheetnames if re.fullmatch(r"\d{2}\.\d{2}", n)]
    return days[-1] if days else None


def create_daily_sheet(wb, d: date, template, run: RunState):
    name = d.strftime("%d.%m")
    ws = wb.copy_worksheet(wb[template])
    ws.title = name

    names = wb.sheetnames
    days = [i for i, n in enumerate(names) if n != name and re.fullmatch(r"\d{2}\.\d{2}", n)]
    if days:  # ставим сразу после последнего дневного листа
        wb.move_sheet(ws, offset=(max(days) + 1) - names.index(name))

    set_cell(ws, "B1", datetime(d.year, d.month, d.day), run)
    if ws["B1"].number_format == "General":
        ws["B1"].number_format = "DD.MM.YYYY"

    # Лист-образец не пустой — стираем в новом дне все интервальные ячейки.
    for row in DATA_DST_ROWS:
        for c in range(FIRST_COL, FIRST_COL + 12):
            set_cell(ws, f"{get_column_letter(c)}{row}", None, run)
    for _cell, cur_row, month_row in CAPTURE:       # в образце там числа какого-то старого дня
        set_cell(ws, f"{get_column_letter(FIRST_COL)}{cur_row}", None, run)
        for c in range(FIRST_COL, FIRST_COL + 31):
            set_cell(ws, f"{get_column_letter(c)}{month_row}", None, run)

    prev = (d - timedelta(days=1)).strftime("%d.%m")
    if prev in wb.sheetnames:  # отклонение к прошлым суткам: D8 =D4-'пред'!D4, D9 =D5-'пред'!D5
        base = f"=D4-'{prev}'!D4"
        for row in (8, 9):
            for c in range(FIRST_COL, FIRST_COL + 12):
                dest = f"{get_column_letter(c)}{row}"
                set_cell(ws, dest, Translator(base, origin="D8").translate_formula(dest), run)
    else:
        # Листа предыдущих суток нет — формулы из образца ссылаются на чужую дату, очищаем.
        log.warning(f"{name}: нет листа предыдущих суток ('{prev}') — отклонение к прошлым "
                    f"суткам (D8:O9) не посчитать, оставляю пустым")
        for row in (8, 9):
            for c in range(FIRST_COL, FIRST_COL + 12):
                set_cell(ws, f"{get_column_letter(c)}{row}", None, run)
    return ws


def apply_manual(ws, cells: dict, run: RunState) -> None:
    """БП и коэффициенты с датами из файла ручных значений (см. manual.py)."""
    for coord, v in cells.items():
        cell = set_cell(ws, coord, v, run)
        if isinstance(v, datetime):
            if cell is not None and cell.number_format == "General":
                cell.number_format = "DD.MM.YYYY"
        elif cell is not None:   # даты в снимок для сверки не идут: он хранится в JSON
            run.written[(ws.title, coord)] = v


def _note(run: RunState, key, vals, tag):
    empty = [v in (None, "") for v in vals]
    if all(empty):
        run.empty.setdefault(key, []).append(tag)
    elif any(empty):
        run.partial.setdefault(key, []).append(tag)
    if any(isinstance(v, str) and v != "" and isinstance(to_num(v), float) for v in vals):
        run.strings.setdefault(key, []).append(tag)


def fill_day(A, B, cc, run: RunState, tag="", debug=False):
    n = cc + 1
    for col, row0, dst_row, ncols in COPY_RULES:
        c0 = column_index_from_string(col)
        for j in range(ncols):
            cl = get_column_letter(c0 + j)
            vals = col_values(A, cl, row0, n)
            key = (f"{cl}{row0}:{cl}{row0 + n - 1}", dst_row + j)
            _note(run, key, vals, tag)
            if debug:
                log.info(f"  строка {dst_row + j:>3} <- {key[0]:<10} {vals}")
            put_row(B, dst_row + j, vals, run)
    for dst_row, cols in SUM_154.items():
        raw = [[A.cell(row=154 + i, column=c).value for c in cols] for i in range(n)]
        _note(run, ("сумма столбцов " + "+".join(get_column_letter(c) for c in cols) + " строки 154+", dst_row),
              [None if all(x in (None, "") for x in r) else 1 for r in raw], tag)
        if debug:
            log.info(f"  строка {dst_row:>3} <- сумма {[get_column_letter(c) for c in cols]} (154+): {raw}")
        put_row(B, dst_row, [sum(num(x) for x in r) for r in raw], run)
    for dst_row, cols in SUM_89.items():
        raw = [[A[f"{c}{89 + i}"].value for c in cols] for i in range(n)]
        _note(run, ("сумма столбцов " + "+".join(cols) + " строки 89+", dst_row),
              [None if all(x in (None, "") for x in r) else 1 for r in raw], tag)
        if debug:
            log.info(f"  строка {dst_row:>3} <- сумма {cols} (89+): {raw}")
        put_row(B, dst_row, [sum(num(x) for x in r) for r in raw], run)


def read_capture(A, run: RunState, tag="") -> tuple:
    """Суточные итоги УКПГ и ЦПС из листа источника (см. CAPTURE); пустая ячейка -> None."""
    out = []
    for cell, cur_row, _month_row in CAPTURE:
        v = A[cell].value
        _note(run, (cell, cur_row), [v], tag)
        out.append(None if v in (None, "") else to_num(v))
    return tuple(out)


def fill_capture(ws, d: date, captured: dict, run: RunState) -> None:
    """Блок «Уловленная» внизу листа дня d. captured: дата -> итоги из read_capture().
    «Текущие сутки» — итог самого дня d, «с начала месяца» — итоги всех дней этого месяца
    от 1-го числа до d включительно (по столбцу на число)."""
    for k, (_cell, cur_row, month_row) in enumerate(CAPTURE):
        put_row(ws, cur_row, [captured[d][k]], run)
        for e, vals in captured.items():
            if (e.year, e.month) == (d.year, d.month) and e <= d:
                put_row(ws, month_row, [vals[k]], run, start_col=FIRST_COL + e.day - 1)


def log_source_report(run: RunState):
    def block(title, d, level):
        if not d:
            return
        log.log(level, title)
        for (rng, row), tags in sorted(d.items(), key=lambda kv: kv[0][1]):
            shown = ",".join(tags[:8]) + ("…" if len(tags) > 8 else "")
            log.log(level, f"   строка {row:>3} <- {rng}: листы {shown} ({len(tags)})")
    block("ПУСТО в источнике (ни одного значения):", run.empty, logging.WARNING)
    block("Частично пусто в источнике:", run.partial, logging.WARNING)
    block("Числа были текстом (преобразованы в числа):", run.strings, logging.INFO)


def fill_sms(dst_wb, A, b_list, run: RunState):
    S = dst_wb["data_sms"]
    for k, v in enumerate(col_values(A, "AC", 89, 12)):
        set_cell(S, f"A{21 + k}", v, run, literal=True)
    for k, v in enumerate(col_values(A, "X", 89, 12)):
        set_cell(S, f"B{21 + k}", v, run, literal=True)
    for k in range(12):
        set_cell(S, f"C{21 + k}", A["AC101"].value, run, literal=True)
    for k, v in enumerate(col_values(A, "Z", 89, 12)):
        set_cell(S, f"D{21 + k}", v, run, literal=True)
    # E..H — ссылками на лист сводки, т.к. openpyxl не пересчитывает формулы
    for dst_col, row, count in (("E", 34, 12), ("F", 42, 12), ("G", 50, 12), ("H", 53, 13)):
        for k in range(count):
            set_cell(S, f"{dst_col}{21 + k}", f"='{b_list}'!{get_column_letter(FIRST_COL + k)}{row}", run)
    SMS = dst_wb["смс"]
    for k in range(12):
        set_cell(SMS, f"A{9 + k}", f"=data_sms!A{4 + k}", run)


# ---------- сверка с прошлым запуском ----------
def _values_equal(old, new) -> bool:
    if old in (None, "") and new in (None, ""):
        return True
    if isinstance(old, (int, float)) and isinstance(new, (int, float)):
        return abs(old - new) <= 1e-6 * max(1.0, abs(new))
    return old == new


def log_cell_changes(state_path: Path, written: dict) -> None:
    """Пишет в лог только то, что изменилось по сравнению с прошлым запуском."""
    prev = read_json(state_path)
    if not prev:
        log.info(f"Снимок значений для сверки изменений создан впервые ({len(written)} ячеек)")
    else:
        changed = added = 0
        for (sheet, coord), new_v in written.items():
            key = f"{sheet}!{coord}"
            if key not in prev:
                added += 1
                continue
            if not _values_equal(prev[key], new_v):
                log.info(f"  изменение: {sheet}!{coord}: было {prev[key]!r} -> стало {new_v!r}")
                changed += 1
        removed = len(set(prev) - {f"{s}!{c}" for s, c in written})
        if changed or added or removed:
            log.info(f"Изменений по сравнению с прошлым запуском: {changed} (новых ячеек: {added}, "
                     f"исчезло из вывода: {removed})")
        else:
            log.info("Изменений по сравнению с прошлым запуском нет")
    try:
        write_json(state_path, {f"{sheet}!{coord}": v for (sheet, coord), v in written.items()})
    except OSError as e:
        log.warning(f"Не удалось сохранить снимок значений '{state_path}': {e}")


def compare_with_reference(ref_path, run: RunState):
    """Сверка записанных значений с готовой сводкой (по значениям, что сохранил Excel)."""
    ref = load_workbook(ref_path, data_only=True)
    by_row, missing = {}, set()
    total = bad = 0
    for (sheet, coord), v in run.written.items():
        if sheet not in ref.sheetnames:
            missing.add(sheet)
            continue
        total += 1
        rv = ref[sheet][coord].value
        if isinstance(v, (int, float)) and isinstance(rv, (int, float)):
            same = abs(v - rv) <= 1e-6 * max(1.0, abs(rv))
        else:
            same = (v in (None, "") and rv in (None, "")) or v == rv
        if not same:
            bad += 1
            row = int("".join(ch for ch in coord if ch.isdigit()))
            by_row.setdefault(row, []).append((sheet, coord, v, rv))
    log.info(f"Сверка с {ref_path}: сравнено {total}, отличаются {bad}")
    if missing:
        log.info(f"  нет в эталоне листов: {sorted(missing)}")
    for row in sorted(by_row):
        items = by_row[row]
        sh, co, v, rv = items[0]
        log.info(f"  строка {row}: {len(items)} расхождений, пример {sh}!{co}: у нас {v!r}, в эталоне {rv!r}")


# ---------- основная обработка ----------
def build_summary(src, out, cfg: Config, compare=None, debug=None) -> Path:
    """Собирает сводку из src и сохраняет в out. Возвращает путь к сохранённому файлу.

    SvodkaError — понятная остановка (плохой файл, нет образца); остальные
    исключения пробрасываются дальше. Оба случая попадают в лог.
    """
    run = RunState()
    src = Path(src)
    out = Path(out)
    try:
        now = datetime.now(ZoneInfo(cfg.build.timezone)).replace(tzinfo=None)
        today = (now - timedelta(hours=3)).date()          # день «переключается» в 03:00

        if not src.is_file():
            raise SvodkaError(f"Источник '{src}' не найден")
        check_source_file(src, cfg.limits.max_file_mb, cfg.limits.max_unpacked_mb)

        log.info(f"--- Обработка '{src.name}' ---")
        source_obj = open_source(str(src))
        # '00' — последние сутки предыдущего месяца, нужны для «отклонения к прошлым
        # суткам» у первого дня месяца. Дата берётся из самого листа, не из имени.
        days = sorted((int(n), n) for n in source_obj.sheetnames if n.strip().isdigit() and 0 <= int(n) <= 31)
        if not days:
            raise SvodkaError(f"В источнике нет дневных листов ('00'..'31'). Листы: {source_obj.sheetnames}")

        template_path = cfg.paths.template_file
        if not template_path.is_file():
            raise SvodkaError(f"Не найден файл-образец '{template_path}' (параметр template_file в настройках)")
        dst_wb = load_workbook(template_path)  # уже содержит data_sms, образец дня, смс, График
        tmpl = pick_template(dst_wb, cfg.build.template_sheet or None)
        if tmpl is None:
            raise SvodkaError(f"В {template_path} нет листа-образца 'дд.мм'. Листы: {dst_wb.sheetnames}")
        if tmpl == "31.05":
            log.warning(f"Лист-образец для новых дней: '{tmpl}' — в {template_path.name} нет листа 'sample', "
                        f"используется откат на последний 'дд.мм'. Проверьте, что файл-образец не устарел.")
        else:
            log.info(f"Лист-образец для новых дней: '{tmpl}'")

        manual = load_manual_values(cfg.paths.manual_values)
        if manual is None:
            log.warning(f"Файл ручных значений '{cfg.paths.manual_values}' не найден — БП и коэффициенты "
                        f"остаются как в образце (могут быть устаревшими)")
        else:
            log.info(f"Ручные значения (БП, коэффициенты): {len(manual)} ячеек из '{cfg.paths.manual_values.name}'")

        month = cfg.build.month or None
        year = cfg.build.year or None
        fallback_year = year or today.year
        fallback_month = month or month_from_filename(str(src)) or today.month

        last = None
        captured = {}   # дата -> суточные итоги УКПГ и ЦПС; нужны для таблицы «с начала месяца»
        for day, sheet_name in days:
            A = source_obj[sheet_name]
            d = find_sheet_date(A)
            if d is None:
                try:
                    d = date(fallback_year, fallback_month, day)
                except ValueError:
                    continue
                log.warning(f"{sheet_name}: дата в самом листе не найдена, беру {d:%d.%m.%Y}")
            elif d.day != day:
                log.warning(f"имя листа '{sheet_name}' не совпадает с датой внутри него "
                            f"({d:%d.%m.%Y}) — использую дату из листа")
            if month and d.month != month:
                continue
            if year and d.year != year:
                continue
            if d > today:
                log.info(f"{d:%d.%m.%Y}: дата в будущем, пропуск")
                continue
            cc = copy_count(now) if d == today else 11   # прошлые дни — все 12 интервалов
            b_list = d.strftime("%d.%m")
            if b_list not in dst_wb.sheetnames:
                create_daily_sheet(dst_wb, d, tmpl, run)
            is_debug = debug == sheet_name.strip()
            if is_debug:
                log.info(f"=== ОТЛАДКА: лист источника '{sheet_name}' ({d:%d.%m.%Y}), читаем {cc + 1} интервалов ===")
            fill_day(A, dst_wb[b_list], cc, run, sheet_name, debug=is_debug)
            if manual:
                apply_manual(dst_wb[b_list], manual, run)
            captured[d] = read_capture(A, run, sheet_name)
            fill_capture(dst_wb[b_list], d, captured, run)
            if sheet_name.strip() == "00":
                # Лист-якорь предыдущего месяца: нужен только формулам, показывать его не надо.
                dst_wb[b_list].sheet_state = "hidden"
                log.info(f"{b_list}: заполнено интервалов {cc + 1} (лист-якорь '00', скрыт)")
            else:
                last = (A, b_list)
                log.info(f"{b_list}: заполнено интервалов {cc + 1}")

        if last and cfg.build.fill_sms:
            if all(n in dst_wb.sheetnames for n in SERVICE_SHEETS):
                fill_sms(dst_wb, *last, run)
                log.info(f"data_sms и смс заполнены по листу {last[1]}")
            else:
                log.warning(f"Листы {SERVICE_SHEETS} не найдены — data_sms/смс пропущены")

        if "sample" in dst_wb.sheetnames:
            # Служебный лист-образец в результат не попадает (его формулы ссылаются на
            # чужие даты и ломают строгие программы чтения).
            del dst_wb["sample"]

        dst_wb.calculation.fullCalcOnLoad = True
        out.parent.mkdir(parents=True, exist_ok=True)
        tmp = out.with_name(out.name + ".tmp")
        dst_wb.save(tmp)
        os.replace(tmp, out)   # результат появляется сразу целиком
        log_cell_changes(cfg.paths.state / f"{out.name}.cells.json", run.written)
        if cfg.build.detailed_report:
            log_source_report(run)
            if run.skipped:
                uniq = sorted(set(run.skipped))
                log.warning(f"пропущено {len(uniq)} ячеек внутри объединённых диапазонов (значение не записано):")
                for sheet, coord in uniq[:20]:
                    log.warning(f"   {sheet}!{coord}")
                if len(uniq) > 20:
                    log.warning("   ...")
        log.info(f"Сохранено: {out}")
        if compare:
            compare_with_reference(compare, run)
        return out
    except SvodkaError as e:
        log.error(str(e))
        raise
    except ConfigError as e:
        log.error(f"Сводка не собрана: {e}")
        raise
    except Exception:
        log.exception(f"Непредвиденная ошибка при обработке '{src.name}'")
        raise
