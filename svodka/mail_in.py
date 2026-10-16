"""Приём файлов ПОНМ из локального почтового ящика (mbox, например Thunderbird)."""
import email
import logging
import mailbox
import re
from email.header import decode_header
from email.utils import parseaddr
from pathlib import Path

from .builder import OWN_NAME_PREFIX
from .config import Config
from .safeio import SOURCE_EXTS, inside, read_json, safe_filename, unique_path, write_json

log = logging.getLogger("svodka")

OWN_SUBJECT_PREFIX = OWN_NAME_PREFIX + " — "
MAX_PROCESSED_IDS = 20000


def decode_text(value: str) -> str:
    if not value:
        return ""
    out = []
    for text, enc in decode_header(value):
        if isinstance(text, bytes):
            try:
                out.append(text.decode(enc or "utf-8", errors="replace"))
            except (LookupError, TypeError):
                out.append(text.decode("utf-8", errors="replace"))
        else:
            out.append(text)
    return "".join(out)


def sender_allowed(addr: str, allowed: tuple) -> bool:
    """Адрес разрешён, если он есть в списке или его домен указан как '@домен.ru'."""
    addr = addr.strip().lower()
    if not addr or "@" not in addr:
        return False
    domain = "@" + addr.rsplit("@", 1)[1]
    return addr in allowed or domain in allowed


def _clean_addr(addr: str) -> str:
    """Адрес без переводов строк и пробелов (защита от подмены заголовков при ответе)."""
    return addr if re.fullmatch(r"[^\s<>,;\"]+@[^\s<>,;\"]+", addr or "") else ""


def _repair_broken_header_block(raw: bytes) -> bytes:
    """Thunderbird иногда оставляет мусор с X-Mozilla-Keys перед настоящими заголовками."""
    if b"X-Mozilla-Keys" not in raw:
        return raw
    m = re.search(rb"\r?\n(Received:|From:|Message-ID:) ", raw)
    if not m:
        return raw
    return raw[m.start():].lstrip(b"\r\n")


def iter_mbox_files(path: Path):
    if path.is_file():
        yield path
        return
    if not path.is_dir():
        return
    for p in sorted(path.iterdir()):
        if p.is_file() and p.suffix.lower() != ".msf":
            yield p


def _state_path(cfg: Config, mfile: Path) -> Path:
    safe = re.sub(r"[^\w.-]", "_", mfile.name)
    return cfg.paths.state / f"mbox_{safe}.json"


def _fingerprint(path: Path) -> dict:
    st = path.stat()
    return {"mtime": st.st_mtime, "size": st.st_size}


def _extract(cfg: Config, mbox_path: Path, state: dict) -> tuple:
    dest_dir = cfg.paths.inbox
    max_bytes = cfg.limits.max_file_mb * 1024 * 1024
    processed = set(state.get("processed_ids", []))
    extracted = 0
    had_errors = False
    try:
        box = mailbox.mbox(str(mbox_path), create=False)
    except Exception as e:
        log.warning(f"[{mbox_path.name}] не удалось открыть как mbox: {e}")
        return 0, True

    try:
        for key in box.iterkeys():
            key_id = f"{mbox_path.name}:key{key}"
            if key_id in processed:
                continue
            try:
                msg = box.get_message(key)
            except Exception as e:
                log.warning(f"[{mbox_path.name}] письмо #{key} не читается: {e} — попробую снова позже")
                had_errors = True
                continue

            if not msg.is_multipart():
                payload = msg.get_payload(decode=True) or b""
                if b"Content-Disposition: attachment" in payload or b"filename=" in payload:
                    try:
                        raw = bytes(box.get_bytes(key))
                    except Exception:
                        raw = None
                    if raw is not None:
                        repaired = _repair_broken_header_block(raw)
                        if repaired is not raw:
                            try:
                                msg = email.message_from_bytes(repaired)
                                log.info(f"[{mbox_path.name}] письмо #{key}: заголовки были повреждены — восстановил")
                            except Exception as e:
                                log.warning(f"[{mbox_path.name}] письмо #{key}: не удалось перечитать "
                                            f"после восстановления заголовков: {e}")

            msg_id = msg.get("Message-ID") or key_id
            if msg_id in processed:
                processed.add(key_id)
                continue

            subject = decode_text(msg.get("Subject", ""))
            # Своё же отправленное письмо (готовая сводка) не принимаем за новый ПОНМ —
            # иначе сводка бесконечно пересобирала бы сама себя. Проверяем и заголовок,
            # и тему: некоторые серверы стирают нестандартные заголовки при пересылке.
            if msg.get("X-Svodka-Auto") or OWN_SUBJECT_PREFIX in subject:
                log.info(f"[{mbox_path.name}] письмо #{key} (тема: {subject[:80]}) — наше же отправленное, пропускаю")
                processed.update((msg_id, key_id))
                continue

            sender_addr = _clean_addr(parseaddr(decode_text(msg.get("From", "")))[1])
            if cfg.mail_in.restrict_senders and not sender_allowed(sender_addr, cfg.mail_in.allowed_senders):
                log.warning(f"[{mbox_path.name}] письмо #{key} от '{sender_addr or '?'}' не из списка "
                            f"разрешённых отправителей — вложения не принимаю")
                processed.update((msg_id, key_id))
                continue

            save_failed = False
            for part in msg.walk():
                if part.get_content_maintype() == "multipart":
                    continue
                raw_name = part.get_filename()
                if not raw_name:
                    continue
                filename = safe_filename(decode_text(raw_name), SOURCE_EXTS)
                if filename is None:
                    continue
                if filename.startswith(OWN_NAME_PREFIX):
                    log.info(f"[{mbox_path.name}] письмо #{key}: вложение '{filename}' выглядит как "
                             f"уже готовая сводка, а не исходный ПОНМ — пропускаю")
                    continue
                payload = part.get_payload(decode=True)
                if not payload:
                    continue
                if len(payload) > max_bytes:
                    log.warning(f"[{mbox_path.name}] вложение '{filename}' больше {cfg.limits.max_file_mb} МБ — пропускаю")
                    continue

                dest = unique_path(dest_dir / filename)
                if not inside(dest, dest_dir):
                    log.warning(f"[{mbox_path.name}] вложение '{filename}' отклонено: недопустимое имя")
                    continue
                try:
                    dest_dir.mkdir(parents=True, exist_ok=True)
                    dest.write_bytes(payload)
                    # адрес отправителя — рядом, в папке состояния: по нему можно ответить
                    write_json(cfg.paths.state / f"{dest.name}.json", {"sender": sender_addr})
                except OSError as e:
                    log.warning(f"[{mbox_path.name}] не удалось сохранить вложение '{filename}': {e} — "
                                f"попробую снова позже")
                    had_errors = True
                    save_failed = True
                    continue
                log.info(f"[{mbox_path.name}] из письма (тема: {subject[:80]}, от: {sender_addr}) "
                         f"извлечено вложение -> {dest.name}")
                extracted += 1

            if save_failed:
                continue  # не считаем письмо обработанным, иначе вложение потеряется
            processed.update((msg_id, key_id))
    finally:
        try:
            box.close()
        except Exception:
            pass

    state["processed_ids"] = list(processed)[-MAX_PROCESSED_IDS:]
    return extracted, had_errors


def scan_mailbox(cfg: Config) -> None:
    """Достаёт из почты новые вложения-ПОНМ и кладёт их в папку inbox."""
    mbox_path = cfg.mail_in.mbox_path
    if not mbox_path.exists():
        log.warning(f"Почтовый ящик '{mbox_path}' не найден, жду появления")
        return
    for mfile in iter_mbox_files(mbox_path):
        state_path = _state_path(cfg, mfile)
        state = read_json(state_path)
        try:
            current = _fingerprint(mfile)
        except OSError as e:
            log.warning(f"[{mfile.name}] не удалось прочитать атрибуты почтового файла: {e}")
            continue
        if current == state.get("last_seen"):
            continue  # почта не менялась

        extracted, had_errors = _extract(cfg, mfile, state)
        if had_errors:
            log.warning(f"[{mfile.name}] не все письма обработаны — проверю этот файл ещё раз на следующем проходе")
        else:
            state["last_seen"] = current
        try:
            write_json(state_path, state)
        except OSError as e:
            log.warning(f"[{mfile.name}] не удалось сохранить состояние почты: {e}")
        if extracted:
            log.info(f"[{mfile.name}] из почты извлечено новых вложений: {extracted}")
