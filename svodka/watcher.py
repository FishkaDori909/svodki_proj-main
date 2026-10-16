"""Слежение за папкой с файлами ПОНМ: новый или изменённый файл -> новая сводка."""
import logging
import time
from datetime import datetime

from pathlib import Path

from . import mail_in, mail_out, pdf
from .builder import OWN_NAME_PREFIX, build_summary, result_name
from .config import Config
from .errors import SvodkaError
from .safeio import SOURCE_EXTS, inside, read_json, unique_path, write_json

log = logging.getLogger("svodka")


def file_fingerprint(path: Path) -> dict:
    st = path.stat()
    return {"mtime": st.st_mtime, "size": st.st_size}


def _fmt(fp) -> str:
    if not fp:
        return "нет данных"
    dt = datetime.fromtimestamp(fp["mtime"]).strftime("%Y-%m-%d %H:%M:%S")
    return f"{fp['size']:,} байт, изменён {dt}".replace(",", " ")


def wait_until_stable(path: Path, checks: int = 2, interval: float = 1.0, timeout: float = 30) -> bool:
    """Ждёт, пока файл перестанет меняться (его могут ещё копировать)."""
    last, stable = None, 0
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            cur = file_fingerprint(path)
        except OSError:
            cur = None
        if cur is not None and cur == last:
            stable += 1
            if stable >= checks:
                return True
        else:
            stable = 0
        last = cur
        time.sleep(interval)
    return False


def is_source_candidate(path: Path) -> bool:
    if path.suffix.lower() not in SOURCE_EXTS:
        return False
    if path.name.startswith("~$"):               # временный файл-замок Excel
        return False
    if path.name.startswith(OWN_NAME_PREFIX):    # наша же готовая сводка
        return False
    return True


def _quarantine(cfg: Config, src: Path, reason: str) -> None:
    errors_dir = cfg.paths.errors
    try:
        errors_dir.mkdir(parents=True, exist_ok=True)
        src.rename(unique_path(errors_dir / src.name))
        log.error(f"[{src.name}] файл не принят ({reason}) — перенесён в '{errors_dir}', повторять бессмысленно")
    except OSError as e:
        log.error(f"[{src.name}] файл не принят ({reason}), и убрать его в '{errors_dir}' не удалось ({e}) — "
                  f"будет проверяться снова")


def _send_mail(cfg: Config, src: Path, out_path: Path, sender: str) -> None:
    attachments = [out_path]
    if cfg.mail_out.attach_pdf:
        try:
            attachments.append(pdf.to_pdf(out_path, out_dir=out_path.parent))
        except pdf.PdfConversionError as e:
            log.warning(f"[{src.name}] не удалось сделать PDF ({e}) — письмо уйдёт без него")
    extra = [sender] if (cfg.mail_out.reply_to_sender and sender) else []
    try:
        attached = mail_out.send_report(cfg.mail_out, src.name, attachments, extra_to=extra)
    except Exception as e:
        log.error(f"[{src.name}] сводка собрана, но письмо не отправлено: {e}")
        return
    note = f", также отправителю {sender}" if extra else ""
    log.info(f"[{src.name}] сводка отправлена по почте ({', '.join(a.name for a in attached) or 'без вложений'}){note}")


def process_one(cfg: Config, src: Path) -> None:
    state_path = cfg.paths.state / f"{src.name}.json"
    state = read_json(state_path)
    try:
        current = file_fingerprint(src)
    except OSError as e:
        log.warning(f"[{src.name}] не удалось прочитать атрибуты файла: {e}")
        return

    manual = cfg.paths.manual_values
    try:
        manual_fp = file_fingerprint(manual) if manual.is_file() else None
    except OSError:
        manual_fp = state.get("manual_seen")

    old_fp = state.get("last_seen")
    if current == old_fp and manual_fp == state.get("manual_seen"):
        return
    if current == old_fp:
        log.info(f"[{src.name}] изменились ручные значения ('{manual.name}') — пересобираю сводку")
    elif old_fp is None:
        log.info(f"[{src.name}] новый файл: {_fmt(current)} — собираю сводку")
    else:
        log.info(f"[{src.name}] файл изменился: было ({_fmt(old_fp)}), стало ({_fmt(current)}) — пересобираю сводку")

    if not wait_until_stable(src):
        log.warning(f"[{src.name}] файл всё ещё меняется — проверю на следующем проходе")
        return

    out_path = cfg.paths.results / result_name(src.stem)
    if not inside(out_path, cfg.paths.results):
        _quarantine(cfg, src, "недопустимое имя файла")
        return
    try:
        build_summary(src, out_path, cfg)
    except PermissionError as e:
        log.warning(f"[{src.name}] файл занят другой программой: {e} — попробую на следующей проверке")
        return
    except SvodkaError as e:
        _quarantine(cfg, src, str(e))
        return
    except Exception:
        log.error(f"[{src.name}] пересборка не удалась — повторю на следующей проверке")
        return

    state["last_seen"] = current
    state["manual_seen"] = manual_fp
    state["last_success"] = time.strftime("%Y-%m-%d %H:%M:%S")
    state["out"] = out_path.name
    write_json(state_path, state)
    log.info(f"[{src.name}] сводка пересобрана -> {out_path.name}")

    if cfg.mail_out.enabled:
        _send_mail(cfg, src, out_path, state.get("sender", ""))


def scan_and_process(cfg: Config) -> None:
    inbox = cfg.paths.inbox
    inbox.mkdir(parents=True, exist_ok=True)
    for src in sorted(p for p in inbox.iterdir() if p.is_file() and is_source_candidate(p)):
        process_one(cfg, src)


def cleanup_old_files(cfg: Config) -> None:
    """Удаляет файлы старше keep_days дней из inbox/results/errors/state (0 — не удалять)."""
    days = cfg.watch.keep_days
    if not days:
        return
    cutoff = time.time() - days * 86400
    for folder in (cfg.paths.inbox, cfg.paths.results, cfg.paths.errors, cfg.paths.state):
        if not folder.is_dir():
            continue
        for p in folder.iterdir():
            # сведения о прочитанных письмах нужны всегда, иначе старые письма примутся заново
            if not p.is_file() or p.name.startswith("mbox_"):
                continue
            try:
                if p.stat().st_mtime < cutoff:
                    p.unlink()
                    log.info(f"Удалён старый файл: {p.name}")
            except OSError as e:
                log.warning(f"Не удалось удалить '{p.name}': {e}")


def run_once(cfg: Config) -> None:
    if cfg.mail_in.enabled:
        mail_in.scan_mailbox(cfg)
    scan_and_process(cfg)
    cleanup_old_files(cfg)


def run_forever(cfg: Config) -> None:
    log.info(f"=== Слежу за папкой '{cfg.paths.inbox}' каждые {cfg.watch.interval_seconds:g} с; "
             f"результаты в '{cfg.paths.results}' (Ctrl+C — остановить) ===")
    last_cleanup = 0.0
    try:
        while True:
            try:
                if cfg.mail_in.enabled:
                    mail_in.scan_mailbox(cfg)
                scan_and_process(cfg)
                if time.time() - last_cleanup > 3600:
                    cleanup_old_files(cfg)
                    last_cleanup = time.time()
            except Exception:
                # служба не должна падать из-за одного неудачного прохода
                log.exception("Сбой при проверке папки — продолжаю работу")
            time.sleep(cfg.watch.interval_seconds)
    except KeyboardInterrupt:
        log.info("=== Остановлено пользователем ===")
