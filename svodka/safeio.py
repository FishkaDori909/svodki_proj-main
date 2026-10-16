"""Безопасная работа с файлами: имена из писем, проверка размера, атомарная запись."""
import json
import os
import unicodedata
import zipfile
from pathlib import Path

from .errors import SvodkaError

SOURCE_EXTS = (".xlsb", ".xlsm", ".xlsx")
_BAD_CHARS = set('<>:"|?*')
_MAX_NAME = 120


def safe_filename(raw: str, allowed_exts=SOURCE_EXTS):
    """Имя файла из письма -> безопасное имя без путей и служебных символов.
    Возвращает None, если расширение не подходит или имя пустое."""
    if not raw:
        return None
    name = raw.replace("\\", "/").split("/")[-1]
    name = unicodedata.normalize("NFC", name)
    name = "".join(ch for ch in name if ch.isprintable() and ch not in _BAD_CHARS).strip(" .")
    stem, dot, ext = name.rpartition(".")
    if not dot or not stem.strip() or f".{ext.lower()}" not in allowed_exts:
        return None
    return f"{stem.strip()[:_MAX_NAME]}.{ext.lower()}"


def inside(path: Path, folder: Path) -> bool:
    """Лежит ли path строго внутри folder (после раскрытия .. и ссылок)."""
    try:
        path.resolve().relative_to(folder.resolve())
        return True
    except ValueError:
        return False


def check_source_file(path: Path, max_file_mb: int, max_unpacked_mb: int) -> None:
    """Отсеивает слишком большие файлы и «бомбы» (маленький архив, огромное содержимое)."""
    size = path.stat().st_size
    if size > max_file_mb * 1024 * 1024:
        raise SvodkaError(f"файл слишком большой: {size // 1024 // 1024} МБ (предел {max_file_mb} МБ)")
    if not zipfile.is_zipfile(path):
        raise SvodkaError("это не книга Excel (файл повреждён или подменён)")
    try:
        with zipfile.ZipFile(path) as z:
            unpacked = sum(i.file_size for i in z.infolist())
    except zipfile.BadZipFile as e:
        raise SvodkaError(f"архив книги повреждён: {e}") from e
    if unpacked > max_unpacked_mb * 1024 * 1024:
        raise SvodkaError(f"содержимое книги слишком велико после распаковки: "
                          f"{unpacked // 1024 // 1024} МБ (предел {max_unpacked_mb} МБ)")


def read_json(path: Path) -> dict:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError, UnicodeDecodeError):
        return {}


def write_json(path: Path, data: dict) -> None:
    """Запись через временный файл: при сбое старый файл остаётся целым."""
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    os.replace(tmp, path)


def unique_path(path: Path) -> Path:
    """Если файл с таким именем уже есть — добавляет к имени метку времени."""
    if not path.exists():
        return path
    from datetime import datetime
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return path.with_name(f"{path.stem}_{ts}{path.suffix}")
