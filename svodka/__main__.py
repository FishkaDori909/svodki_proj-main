"""Запуск: python -m svodka <команда>

  check   проверить настройки
  watch   следить за папкой постоянно
  once    один проход по папке и выход (для расписания)
  build   собрать сводку из одного файла
"""
import argparse
import sys
from pathlib import Path

from .builder import build_summary, result_name
from .config import load_config
from .errors import ConfigError
from .logsetup import log, setup_logging
from .manual import load_manual_values


def main(argv=None) -> int:
    p = argparse.ArgumentParser(prog="python -m svodka", description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--config", default=None, help="путь к файлу настроек (по умолчанию config/config.toml)")
    sub = p.add_subparsers(dest="cmd", required=True)
    sub.add_parser("check", help="проверить настройки")
    sub.add_parser("watch", help="следить за папкой постоянно")
    sub.add_parser("once", help="один проход по папке и выход")
    b = sub.add_parser("build", help="собрать сводку из одного файла")
    b.add_argument("src", help="файл ПОНМ (.xlsb / .xlsm / .xlsx)")
    b.add_argument("--out", default=None, help="куда сохранить (по умолчанию в папку результатов)")
    b.add_argument("--debug", default=None, metavar="ЛИСТ", help="показать, что прочитано с листа источника, например 05")
    b.add_argument("--compare", default=None, metavar="ФАЙЛ", help="сверить результат с готовой сводкой")
    args = p.parse_args(argv)

    try:
        cfg = load_config(args.config)
    except ConfigError as e:
        print(f"Ошибка настроек: {e}", file=sys.stderr)
        return 2

    setup_logging(cfg.paths.log_file)
    for w in cfg.warnings:
        log.warning(f"Настройки: {w}")

    if args.cmd == "check":
        log.info(f"Настройки в порядке ({cfg.path}).")
        log.info(f"  Входящие файлы:  {cfg.paths.inbox}")
        log.info(f"  Готовые сводки:  {cfg.paths.results}")
        log.info(f"  Приём из почты:  {'включён' if cfg.mail_in.enabled else 'выключен'}")
        log.info(f"  Отправка почты:  {'включена' if cfg.mail_out.enabled else 'выключена'}")
        if not cfg.paths.template_file.is_file():
            log.error(f"  Не найден файл-образец: {cfg.paths.template_file}")
            return 1
        try:
            manual = load_manual_values(cfg.paths.manual_values)
        except ConfigError as e:
            log.error(f"  Ошибка в файле ручных значений: {e}")
            return 1
        if manual is None:
            log.warning(f"  Ручные значения:  файл {cfg.paths.manual_values} не найден — БП и коэффициенты "
                        f"будут как в образце")
        else:
            log.info(f"  Ручные значения:  {cfg.paths.manual_values} ({len(manual)} ячеек)")
        return 0

    from . import watcher
    if args.cmd == "watch":
        watcher.run_forever(cfg)
        return 0
    if args.cmd == "once":
        watcher.run_once(cfg)
        return 0

    src = Path(args.src)
    out = Path(args.out) if args.out else cfg.paths.results / result_name(src.stem)
    try:
        build_summary(src, out, cfg, compare=args.compare, debug=args.debug)
    except Exception:   # причина уже записана в журнал
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
