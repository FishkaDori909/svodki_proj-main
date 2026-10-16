"""Единственное место, где читаются настройки: config/config.toml.

Пароль к почте можно не хранить в файле, а передать переменной окружения
SVODKA_SMTP_PASSWORD — она важнее значения из файла.
"""
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

if sys.version_info >= (3, 11):
    import tomllib
else:  # pragma: no cover
    import tomli as tomllib

from .errors import ConfigError

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_CONFIG = PROJECT_ROOT / "config" / "config.toml"
ENV_CONFIG = "SVODKA_CONFIG"
ENV_SMTP_PASSWORD = "SVODKA_SMTP_PASSWORD"

_LOCAL_HOSTS = ("localhost", "127.0.0.1", "::1")


@dataclass(frozen=True)
class Paths:
    inbox: Path
    results: Path
    errors: Path
    state: Path
    log_file: Path
    template_file: Path
    manual_values: Path


@dataclass(frozen=True)
class Watch:
    interval_seconds: float
    keep_days: int


@dataclass(frozen=True)
class Build:
    timezone: str
    template_sheet: str
    month: int
    year: int
    fill_sms: bool
    detailed_report: bool


@dataclass(frozen=True)
class Limits:
    max_file_mb: int
    max_unpacked_mb: int


@dataclass(frozen=True)
class MailIn:
    enabled: bool
    mbox_path: Path | None
    restrict_senders: bool
    allowed_senders: tuple


@dataclass(frozen=True)
class Smtp:
    server: str
    port: int
    security: str
    verify_certificate: bool
    login: str
    password: str


@dataclass(frozen=True)
class MailOut:
    enabled: bool
    sender: str
    to: tuple
    reply_to_sender: bool
    attach_pdf: bool
    max_attachment_mb: int
    smtp: Smtp


@dataclass(frozen=True)
class Config:
    path: Path
    paths: Paths
    watch: Watch
    build: Build
    limits: Limits
    mail_in: MailIn
    mail_out: MailOut
    warnings: tuple


# раздел -> {ключ: значение по умолчанию}; тип значения берётся из значения по умолчанию
_SCHEMA = {
    "paths": {
        "inbox": "data/inbox", "results": "data/results", "errors": "data/errors",
        "state": "data/state", "log_file": "logs/svodka.log",
        "template_file": "templates/svodka_template.xlsx",
        "manual_values": "config/manual_values.toml",
    },
    "watch": {"interval_seconds": 60, "keep_days": 30},
    "build": {"timezone": "Asia/Yekaterinburg", "template_sheet": "", "month": 0, "year": 0,
              "fill_sms": True, "detailed_report": False},
    "limits": {"max_file_mb": 50, "max_unpacked_mb": 500},
    "mail_in": {"enabled": False, "mbox_path": "", "restrict_senders": True, "allowed_senders": []},
    "mail_out": {"enabled": False, "from": "", "to": [], "reply_to_sender": False,
                 "attach_pdf": True, "max_attachment_mb": 10},
    "smtp": {"server": "", "port": 587, "security": "starttls", "verify_certificate": True,
             "login": "", "password": ""},
}


def _read_section(raw: dict, name: str) -> dict:
    defaults = _SCHEMA[name]
    section = raw.get(name, {})
    if not isinstance(section, dict):
        raise ConfigError(f"[{name}]: ожидался раздел, а не значение")
    unknown = set(section) - set(defaults)
    if unknown:
        raise ConfigError(f"[{name}]: неизвестные параметры: {', '.join(sorted(unknown))}. "
                          f"Проверьте написание (допустимы: {', '.join(defaults)})")
    out = {}
    for key, default in defaults.items():
        value = section.get(key, default)
        if isinstance(default, bool):
            ok = isinstance(value, bool)
        elif isinstance(default, int):
            ok = isinstance(value, int) and not isinstance(value, bool)
        elif isinstance(default, float):
            ok = isinstance(value, (int, float)) and not isinstance(value, bool)
        elif isinstance(default, list):
            ok = isinstance(value, list) and all(isinstance(v, str) for v in value)
        else:
            ok = isinstance(value, str)
        if not ok:
            raise ConfigError(f"[{name}] {key}: неверный тип значения ({value!r})")
        out[key] = value
    return out


def _resolve(value: str) -> Path:
    p = Path(value).expanduser()
    return p if p.is_absolute() else PROJECT_ROOT / p


def _positive(section: str, key: str, value, minimum=1):
    if value < minimum:
        raise ConfigError(f"[{section}] {key}: должно быть не меньше {minimum}")
    return value


def load_config(path=None) -> Config:
    """Читает и проверяет настройки. Ошибки — ConfigError с понятным текстом."""
    cfg_path = Path(path or os.environ.get(ENV_CONFIG) or DEFAULT_CONFIG)
    if not cfg_path.is_file():
        raise ConfigError(
            f"Файл настроек '{cfg_path}' не найден. Скопируйте config/config.example.toml "
            f"в config/config.toml и впишите свои значения.")
    try:
        raw = tomllib.loads(cfg_path.read_text(encoding="utf-8"))
    except (tomllib.TOMLDecodeError, UnicodeDecodeError) as e:
        raise ConfigError(f"Файл настроек '{cfg_path}' записан с ошибкой: {e}") from e

    unknown = set(raw) - set(_SCHEMA)
    if unknown:
        raise ConfigError(f"Неизвестные разделы в '{cfg_path}': {', '.join(sorted(unknown))}")

    warnings = []
    s = {name: _read_section(raw, name) for name in _SCHEMA}

    paths = Paths(**{k: _resolve(v) for k, v in s["paths"].items()})

    watch = Watch(
        interval_seconds=float(_positive("watch", "interval_seconds", s["watch"]["interval_seconds"], 5)),
        keep_days=_positive("watch", "keep_days", s["watch"]["keep_days"], 0),
    )

    b = s["build"]
    try:
        ZoneInfo(b["timezone"])
    except (ZoneInfoNotFoundError, ValueError) as e:
        raise ConfigError(f"[build] timezone: неизвестный часовой пояс '{b['timezone']}'") from e
    if b["month"] and not 1 <= b["month"] <= 12:
        raise ConfigError("[build] month: должно быть 0 (не задан) или число от 1 до 12")
    if b["year"] and not 2000 <= b["year"] <= 2100:
        raise ConfigError("[build] year: должно быть 0 (не задан) или год, например 2026")
    build = Build(**b)

    lim = s["limits"]
    limits = Limits(max_file_mb=_positive("limits", "max_file_mb", lim["max_file_mb"]),
                    max_unpacked_mb=_positive("limits", "max_unpacked_mb", lim["max_unpacked_mb"]))

    mi = s["mail_in"]
    senders = tuple(a.strip().lower() for a in mi["allowed_senders"] if a.strip())
    if mi["enabled"]:
        if not mi["mbox_path"]:
            raise ConfigError("[mail_in] mbox_path: укажите путь к почтовому ящику")
        if mi["restrict_senders"] and not senders:
            raise ConfigError("[mail_in] allowed_senders: список пуст. Перечислите адреса (или "
                              "'@домен.ru'), от кого разрешено принимать файлы ПОНМ, либо "
                              "отключите проверку: restrict_senders = false")
        if not mi["restrict_senders"]:
            warnings.append("Проверка отправителей отключена ([mail_in] restrict_senders = false): "
                            "вложения принимаются от любого, кто пишет на этот ящик.")
    mail_in = MailIn(enabled=mi["enabled"], mbox_path=_resolve(mi["mbox_path"]) if mi["mbox_path"] else None,
                     restrict_senders=mi["restrict_senders"], allowed_senders=senders)

    sm = s["smtp"]
    password = os.environ.get(ENV_SMTP_PASSWORD) or sm["password"]
    smtp = Smtp(server=sm["server"], port=sm["port"], security=sm["security"].lower(),
                verify_certificate=sm["verify_certificate"], login=sm["login"], password=password)
    mo = s["mail_out"]
    mail_out = MailOut(enabled=mo["enabled"], sender=mo["from"], to=tuple(mo["to"]),
                       reply_to_sender=mo["reply_to_sender"], attach_pdf=mo["attach_pdf"],
                       max_attachment_mb=_positive("mail_out", "max_attachment_mb", mo["max_attachment_mb"]),
                       smtp=smtp)
    if mo["enabled"]:
        missing = [name for name, v in (("[mail_out] from", mo["from"]), ("[mail_out] to", mo["to"]),
                                        ("[smtp] server", sm["server"])) if not v]
        if missing:
            raise ConfigError(f"Отправка почты включена, но не заполнено: {', '.join(missing)}")
        if smtp.security not in ("starttls", "ssl", "none"):
            raise ConfigError("[smtp] security: допустимо 'starttls', 'ssl' или 'none'")
        if smtp.login and not smtp.password:
            raise ConfigError(f"[smtp] указан login, но нет пароля: задайте переменную окружения "
                              f"{ENV_SMTP_PASSWORD} (или password в файле)")
        if smtp.security == "none" and smtp.login and smtp.server not in _LOCAL_HOSTS:
            raise ConfigError("[smtp] security = 'none' с логином и паролем разрешён только для "
                              "локального сервера: иначе пароль пойдёт по сети открытым текстом")
        if smtp.security != "none" and not smtp.verify_certificate:
            warnings.append("Проверка сертификата почтового сервера отключена ([smtp] verify_certificate = "
                            "false): соединение зашифровано, но не проверяется, что отвечает настоящий "
                            "сервер. Используйте только во внутренней сети.")
        if sm["password"] and not os.environ.get(ENV_SMTP_PASSWORD):
            warnings.append(f"Пароль почты записан в файле настроек. Надёжнее хранить его в переменной "
                            f"окружения {ENV_SMTP_PASSWORD}.")
            if os.name == "posix" and cfg_path.stat().st_mode & 0o007:
                warnings.append(f"Файл '{cfg_path}' доступен всем пользователям сервера — "
                                f"выполните: chmod 640 {cfg_path}")

    return Config(path=cfg_path, paths=paths, watch=watch, build=build, limits=limits,
                  mail_in=mail_in, mail_out=mail_out, warnings=tuple(warnings))
