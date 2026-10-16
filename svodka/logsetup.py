import logging
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path

log = logging.getLogger("svodka")
_configured = False


def setup_logging(log_file: Path, level=logging.INFO) -> logging.Logger:
    """Пишет в консоль и в файл (ротация 5 МБ x 5 файлов). Повторный вызов ничего не меняет."""
    global _configured
    if _configured:
        return log
    log.setLevel(level)
    try:
        log_file.parent.mkdir(parents=True, exist_ok=True)
        fh = RotatingFileHandler(log_file, maxBytes=5 * 1024 * 1024, backupCount=5, encoding="utf-8")
        fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S"))
        log.addHandler(fh)
    except OSError as e:
        print(f"ВНИМАНИЕ: не удалось открыть лог-файл '{log_file}': {e}", file=sys.stderr)
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(logging.Formatter("%(message)s"))
    log.addHandler(sh)
    log.propagate = False
    _configured = True
    return log
