from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

from .models import cache_key

CONFIG_FILENAME = "mria-config.txt"
CACHE_FILENAME = "choice-cache.txt"
FIRST_RUN_MESSAGE = (
    "Файлы конфигурации созданы, заполните mria-config.txt и перезапустите программу. "
    "Для сброса кеша выбора удалите файл choice-cache.txt"
)
DEFAULT_CONFIG_TEXT = (
    'Распределение="./Размеченная_программа.xlsx"\n'
    'Партнёры="./Партнёры.xlsx"\n'
)


@dataclass(slots=True)
class ConfigPaths:
    distribution_path: Path
    partners_path: Path
    config_path: Path
    cache_path: Path


class ChoiceCache:
    def __init__(self, path: Path, mappings: dict[str, dict[str, str]]) -> None:
        self.path = path
        self._mappings = mappings

    @classmethod
    def create_default_file(cls, path: Path) -> None:
        default_payload = {"version": 1, "mappings": {}}
        path.write_text(
            json.dumps(default_payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    @classmethod
    def load(cls, path: Path, logger) -> "ChoiceCache":
        if not path.exists():
            logger.info("Файл кеша не найден, создаю новый: %s", path.resolve())
            cls.create_default_file(path)

        raw_text = path.read_text(encoding="utf-8").strip()
        if not raw_text:
            payload: dict[str, object] = {"version": 1, "mappings": {}}
        else:
            payload = json.loads(raw_text)

        mappings_obj = payload.get("mappings")
        if not isinstance(mappings_obj, dict):
            raise ValueError("Некорректный формат choice-cache.txt: ключ 'mappings' должен быть объектом")

        mappings: dict[str, dict[str, str]] = {}
        for key, value in mappings_obj.items():
            if isinstance(value, dict):
                mappings[key] = {str(k): str(v) for k, v in value.items()}

        logger.info("Кеш загружен: %d записей", len(mappings))
        return cls(path=path, mappings=mappings)

    def save(self) -> None:
        payload = {"version": 1, "mappings": self._mappings}
        self.path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def get(self, partner_name: str, event_name: str) -> dict[str, str] | None:
        return self._mappings.get(cache_key(partner_name, event_name))

    def set_skip(self, partner_name: str, event_name: str) -> None:
        self._mappings[cache_key(partner_name, event_name)] = {"action": "skip"}
        self.save()

    def set_no_tz(self, partner_name: str, event_name: str) -> None:
        self._mappings[cache_key(partner_name, event_name)] = {"action": "no_tz"}
        self.save()

    def set_mapping(
        self,
        source_partner: str,
        source_event: str,
        target_partner: str,
        target_event: str,
    ) -> None:
        self._mappings[cache_key(source_partner, source_event)] = {
            "action": "map",
            "partner": target_partner,
            "event": target_event,
        }
        self.save()


def _parse_config_line(line: str) -> tuple[str, str] | None:
    stripped = line.strip()
    if not stripped or stripped.startswith("#"):
        return None
    if "=" not in stripped:
        return None
    key, raw_value = stripped.split("=", 1)
    value = raw_value.strip()
    if (value.startswith('"') and value.endswith('"')) or (value.startswith("'") and value.endswith("'")):
        value = value[1:-1]
    return key.strip(), value


def _read_config_values(config_path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    for raw_line in config_path.read_text(encoding="utf-8").splitlines():
        parsed = _parse_config_line(raw_line)
        if not parsed:
            continue
        key, value = parsed
        values[key] = value
    return values


def load_or_initialize_paths(cwd: Path, logger) -> ConfigPaths | None:
    config_path = cwd / CONFIG_FILENAME
    cache_path = cwd / CACHE_FILENAME

    if not config_path.exists():
        config_path.write_text(DEFAULT_CONFIG_TEXT, encoding="utf-8")
        if not cache_path.exists():
            ChoiceCache.create_default_file(cache_path)
        print(FIRST_RUN_MESSAGE)
        return None

    if not cache_path.exists():
        logger.info("Файл кеша отсутствовал и будет создан: %s", cache_path.resolve())
        ChoiceCache.create_default_file(cache_path)

    values = _read_config_values(config_path)

    distribution_rel = values.get("Распределение")
    partners_rel = values.get("Партнёры")
    if not distribution_rel or not partners_rel:
        raise ValueError(
            "В mria-config.txt должны быть указаны ключи "
            "'Распределение' и 'Партнёры'"
        )

    distribution_path = (cwd / distribution_rel).resolve()
    partners_path = (cwd / partners_rel).resolve()
    logger.info("Путь к файлу 'Распределение': %s", distribution_path)
    logger.info("Путь к файлу 'Партнёры': %s", partners_path)

    return ConfigPaths(
        distribution_path=distribution_path,
        partners_path=partners_path,
        config_path=config_path.resolve(),
        cache_path=cache_path.resolve(),
    )
