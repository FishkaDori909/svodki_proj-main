"""Отправка готовой сводки по почте."""
import base64
import logging
import smtplib
import ssl
from email.header import Header
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate
from pathlib import Path

from .config import MailOut

log = logging.getLogger("svodka")


def _encoded_name(name: str) -> str:
    """Имя файла одной строкой в виде '=?utf-8?b?...?=' (RFC 2047) — так его понимают все
    клиенты и шлюзы. Стандартный способ Python (RFC 2231, 'filename*0*=...') режет длинное
    русское имя на куски, которые многие не склеивают: вложение приходит как 'Часть 1.2'."""
    words, chunk = [], ""
    for ch in name:
        if len((chunk + ch).encode("utf-8")) > 45:      # 45 байт -> 60 знаков base64, слово <= 75
            words.append(chunk)
            chunk = ""
        chunk += ch
    if chunk:
        words.append(chunk)
    return " ".join(f"=?utf-8?b?{base64.b64encode(w.encode('utf-8')).decode('ascii')}?=" for w in words)


def _attach(msg: MIMEMultipart, path: Path) -> None:
    part = MIMEApplication(path.read_bytes(), "octet-stream")
    name = _encoded_name(path.name)
    part.replace_header("Content-Type", f'application/octet-stream; name="{name}"')
    part.add_header("Content-Disposition", f'attachment; filename="{name}"')
    msg.attach(part)


def _connect(smtp_cfg) -> smtplib.SMTP:
    context = ssl.create_default_context()
    if not smtp_cfg.verify_certificate:
        context.check_hostname = False    # порядок важен: сначала имя, потом режим проверки
        context.verify_mode = ssl.CERT_NONE
    if smtp_cfg.security == "ssl":
        return smtplib.SMTP_SSL(smtp_cfg.server, smtp_cfg.port, timeout=30, context=context)
    conn = smtplib.SMTP(smtp_cfg.server, smtp_cfg.port, timeout=30)
    if smtp_cfg.security == "starttls":
        conn.starttls(context=context)
    return conn


def send_report(mail_cfg: MailOut, src_name: str, attachments, extra_to=()) -> list:
    """Отправляет письмо с вложениями. Возвращает список реально приложенных файлов.
    Вложения, не влезающие в max_attachment_mb, не прикладываются (их перечень — в тексте)."""
    to_list = list(mail_cfg.to)
    seen = {a.lower() for a in to_list}
    for addr in extra_to:
        if addr and addr.lower() not in seen:
            to_list.append(addr)
            seen.add(addr.lower())

    limit = mail_cfg.max_attachment_mb * 1024 * 1024
    attached, skipped, total = [], [], 0
    for f in map(Path, attachments):
        size = f.stat().st_size
        if total + size > limit:
            skipped.append(f)
        else:
            attached.append(f)
            total += size

    body = f"Автоматически собранная сводка по файлу «{src_name}».\n"
    if attached:
        body += f"Вложения: {', '.join(f.name for f in attached)}\n"
    if skipped:
        body += (f"Не приложено (превышен предел {mail_cfg.max_attachment_mb} МБ на письмо): "
                 f"{', '.join(f.name for f in skipped)}. Файлы лежат в папке результатов на сервере.\n")

    # Письмо собирается классическими email.mime: они оставляют заголовки как написаны (новый
    # EmailMessage переписал бы имя вложения обратно в «куски»). Защиту от подмены заголовков
    # (перевод строки в адресе или теме) поэтому делаем сами.
    subject = f"Сводка 2-х час ГПН-Ямал — {Path(src_name).stem}"
    for value in (mail_cfg.sender, *to_list, subject):
        if "\r" in value or "\n" in value:
            raise ValueError("перевод строки в адресе или теме письма")
    msg = MIMEMultipart()
    msg["From"] = mail_cfg.sender
    msg["To"] = ", ".join(to_list)
    msg["Date"] = formatdate(localtime=True)
    msg["Subject"] = Header(subject, "utf-8")
    # по этой метке программа узнаёт своё письмо и не принимает его за новый ПОНМ
    msg["X-Svodka-Auto"] = "1"
    msg.attach(MIMEText(body, "plain", "utf-8"))
    for f in attached:
        _attach(msg, f)

    conn = _connect(mail_cfg.smtp)
    try:
        if mail_cfg.smtp.login:
            conn.login(mail_cfg.smtp.login, mail_cfg.smtp.password)
        conn.send_message(msg, from_addr=mail_cfg.sender, to_addrs=to_list)
    finally:
        try:
            conn.quit()
        except smtplib.SMTPException:
            pass
    return attached
