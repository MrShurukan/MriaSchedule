from __future__ import annotations

import argparse
import logging
from pathlib import Path

from . import __version__
from .config_cache import ChoiceCache, load_or_initialize_paths
from .excel_parser import load_partners_records, parse_distribution_workbook
from .matching import PartnerMatcher
from .models import OutputScheduleEvent
from .output_writer import resolve_output_path, write_schedule_workbook


def configure_logging() -> logging.Logger:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )
    return logging.getLogger("mria_scheduler")


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="mria-scheduler",
        description=(
            "Автоматическая сводка расписания из Excel файлов "
            "'Распределение' и 'Партнёры'"
        ),
    )
    parser.add_argument(
        "output",
        nargs="?",
        help=(
            "Имя или путь выходного XLSX файла. "
            "По умолчанию: Расписание_<день>_<месяц>_<год>_<час>_<минута>_<секунда>.xlsx"
        ),
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    return parser


def _validate_input_files(distribution_path: Path, partners_path: Path) -> None:
    if not distribution_path.exists():
        raise FileNotFoundError(f"Файл 'Распределение' не найден: {distribution_path}")
    if not partners_path.exists():
        raise FileNotFoundError(f"Файл 'Партнёры' не найден: {partners_path}")


def run(output_name: str | None = None) -> int:
    cwd = Path.cwd().resolve()
    bootstrap_logger = logging.getLogger("mria_scheduler.bootstrap")
    if not bootstrap_logger.handlers:
        bootstrap_logger.addHandler(logging.NullHandler())
    bootstrap_logger.propagate = False

    config_paths = load_or_initialize_paths(cwd=cwd, logger=bootstrap_logger)
    if config_paths is None:
        return 0

    logger = configure_logging()
    logger.info("Запуск mria-scheduler")
    logger.info("Рабочая директория: %s", cwd)
    logger.info("Путь к конфигу: %s", config_paths.config_path)
    logger.info("Путь к кешу: %s", config_paths.cache_path)
    logger.info("Путь к файлу 'Распределение': %s", config_paths.distribution_path)
    logger.info("Путь к файлу 'Партнёры': %s", config_paths.partners_path)
    _validate_input_files(config_paths.distribution_path, config_paths.partners_path)
    logger.info("Проверка существования входных файлов пройдена")

    choice_cache = ChoiceCache.load(config_paths.cache_path, logger=logger)
    partners_records = load_partners_records(config_paths.partners_path, logger=logger)
    color_to_location, day_shift_columns, events, distribution_theme = parse_distribution_workbook(
        config_paths.distribution_path, logger=logger
    )
    logger.info(
        "Считано для обработки: %d столбцов смен и %d заполненных мероприятий",
        len(day_shift_columns),
        len(events),
    )

    matcher = PartnerMatcher(partner_records=partners_records, cache=choice_cache, logger=logger)
    output_events: list[OutputScheduleEvent] = []
    total = len(events)
    for index, event in enumerate(events, start=1):
        logger.info(
            "Обработка %d/%d: день='%s', смена='%s', партнер='%s', мероприятие='%s'",
            index,
            total,
            event.day_label,
            event.shift_label,
            event.partner_name,
            event.event_name,
        )
        match = matcher.resolve(event.partner_name, event.event_name)
        if match is None:
            logger.info("Событие пропущено: (%s | %s)", event.partner_name, event.event_name)
            continue

        if not event.color_key:
            # Отсутствие цвета = свободная/пустая локация.
            location = ""
            logger.info(
                "Для события нет цвета, локация оставлена пустой: "
                "[день=%s, смена=%s, партнер=%s, мероприятие=%s]",
                event.day_label,
                event.shift_label,
                event.partner_name,
                event.event_name,
            )
        else:
            location = color_to_location.get(event.color_key)
            if not location:
                raise ValueError(
                    "Не найдена локация для цвета события "
                    f"({event.color_key}) [день={event.day_label}, смена={event.shift_label}, "
                    f"партнер={event.partner_name}, мероприятие={event.event_name}]"
                )

        output_events.append(
            OutputScheduleEvent(
                day_label=event.day_label,
                shift_label=event.shift_label,
                event_name=event.event_name,
                location=location,
                partner_name=event.partner_name,
                technical_requirements=match.technical_requirements,
                fill=event.fill,
            )
        )

    output_path = resolve_output_path(cwd, output_name)
    logger.info("Выходной файл: %s", output_path)
    write_schedule_workbook(
        path=output_path,
        events=output_events,
        logger=logger,
        source_theme=distribution_theme,
    )
    logger.info("Завершено успешно. Сформировано событий: %d", len(output_events))
    return 0


def main() -> int:
    parser = build_arg_parser()
    args = parser.parse_args()
    try:
        return run(output_name=args.output)
    except Exception as exc:  # pragma: no cover - CLI safety net
        logger = logging.getLogger("mria_scheduler")
        if not logger.handlers:
            logger = configure_logging()
        logger.exception("Завершено с ошибкой: %s", exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
