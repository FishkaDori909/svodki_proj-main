class SvodkaError(RuntimeError):
    """Ожидаемая, понятная остановка (не баг): неверный или опасный входной файл и т.п."""


class ConfigError(Exception):
    """Ошибка в файле настроек config/config.toml."""
