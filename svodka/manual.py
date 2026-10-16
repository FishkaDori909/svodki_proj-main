"""Значения, которые вписываются в сводку руками: БП и «Коэффициент в расчёте» с датой.

Хранятся в отдельном файле config/manual_values.toml (параметр manual_values в [paths]).
Названия в файле — это подписи из столбца B сводки; по ним находится ячейка листа.
Что не вписано в файл, остаётся таким, как в образце.
"""
import re
import sys
from datetime import date, datetime
from pathlib import Path

if sys.version_info >= (3, 11):
    import tomllib
else:  # pragma: no cover
    import tomli as tomllib

from .errors import ConfigError

# подпись строки (столбец B) -> ячейка со значением БП
BP_CELLS = {
    "Поступление в 2-х час (тн) - БП": "Q5",
    "Добыча ПНГ - БП": "Q30",
    "Использование ПНГ - БП": "Q33",
    "Закачка ПНГ - БП": "Q39",
    "Конденсат ПНГ - БП": "V52",   # в этой строке Q/R заняты заголовками «Коэффициент»/«Дата»
    "Поставка ПНГ в ЕСГ - БП": "Q68",
    "Добыча ПГ - БП": "Q76",
    "Конденсат ПГ - БП": "Q79",
    "Поставка ПГ в ЕСГ - БП": "Q90",
    "Поставка СОГ в ЕСГ - БП": "Q93",
}
# подпись строки (столбец B) -> строка; коэффициент пишется в Q, дата — в R
COEF_ROWS = {
    "Конденсат ПНГ извлеченный": 53,
    "Стабильный конденсат ПНГ (НТК-1)": 55,
    "Стабильный конденсат ПНГ (НТК-2)": 58,
    "Стабильный конденсат ПНГ (УРГ)": 61,
    "Стабильный конденсат ПГ (тёплый)": 81,
    "Стабильный конденсат ПГ (холодный)": 84,
}
COEF_KEY, DATE_KEY = "Коэффициент", "Дата"


def _is_number(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _parse_date(where: str, v) -> datetime:
    if isinstance(v, datetime):
        v = v.date()
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, str) and re.fullmatch(r"\s*\d{1,2}\.\d{1,2}\.\d{4}\s*", v):
        try:
            return datetime.strptime(v.strip(), "%d.%m.%Y")
        except ValueError:
            pass
    raise ConfigError(f"{where}: дата должна быть вида \"31.05.2026\" (или 2026-05-31), а не {v!r}")


def _check_keys(where: str, got: dict, allowed) -> None:
    unknown = set(got) - set(allowed)
    if unknown:
        raise ConfigError(f"{where}: неизвестные названия: "
                          f"{', '.join(repr(k) for k in sorted(unknown))}. Проверьте написание — "
                          f"название должно совпадать с подписью в сводке (допустимы: {', '.join(allowed)})")


def load_manual_values(path: Path) -> dict | None:
    """Читает файл и возвращает {ячейка: значение}. Файла нет — None. Ошибки — ConfigError."""
    if not path.is_file():
        return None
    try:
        raw = tomllib.loads(path.read_text(encoding="utf-8"))
    except (tomllib.TOMLDecodeError, UnicodeDecodeError) as e:
        raise ConfigError(f"Файл ручных значений '{path}' записан с ошибкой: {e}. "
                          f"Частая причина — дробь через запятую: пишите 0.67, а не 0,67") from e

    _check_keys(f"'{path.name}'", raw, ("bp", "coefficients"))
    bp = raw.get("bp", {})
    coefs = raw.get("coefficients", {})
    for name, section in (("bp", bp), ("coefficients", coefs)):
        if not isinstance(section, dict):
            raise ConfigError(f"'{path.name}' [{name}]: ожидался раздел, а не значение")
    _check_keys(f"'{path.name}' [bp]", bp, BP_CELLS)
    _check_keys(f"'{path.name}' [coefficients]", coefs, COEF_ROWS)

    cells = {}
    for name, v in bp.items():
        if not _is_number(v):
            raise ConfigError(f"'{path.name}' [bp] \"{name}\": нужно число (дробная часть через точку), а не {v!r}")
        cells[BP_CELLS[name]] = v
    for name, entry in coefs.items():
        where = f"'{path.name}' [coefficients] \"{name}\""
        if not isinstance(entry, dict):
            raise ConfigError(f"{where}: пишите в виде {{ \"{COEF_KEY}\" = 0.67, \"{DATE_KEY}\" = \"31.05.2026\" }}")
        _check_keys(where, entry, (COEF_KEY, DATE_KEY))
        if COEF_KEY not in entry or DATE_KEY not in entry:
            raise ConfigError(f"{where}: нужны оба значения — {COEF_KEY} и {DATE_KEY}")
        if not _is_number(entry[COEF_KEY]):
            raise ConfigError(f"{where} {COEF_KEY}: нужно число (дробная часть через точку), а не {entry[COEF_KEY]!r}")
        row = COEF_ROWS[name]
        cells[f"Q{row}"] = entry[COEF_KEY]
        cells[f"R{row}"] = _parse_date(f"{where} {DATE_KEY}", entry[DATE_KEY])
    return cells
