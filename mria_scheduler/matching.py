from __future__ import annotations

from dataclasses import dataclass

from rapidfuzz import fuzz

from .config_cache import ChoiceCache
from .models import PartnerRecord, normalize_text


@dataclass(slots=True)
class MatchCandidate:
    record: PartnerRecord
    partner_score: float
    event_score: float
    combined_score: float


class PartnerMatcher:
    def __init__(self, partner_records: list[PartnerRecord], cache: ChoiceCache, logger) -> None:
        self.partner_records = partner_records
        self.cache = cache
        self.logger = logger
        self._exact_index: dict[str, list[PartnerRecord]] = {}
        for record in partner_records:
            key = self._exact_key(record.partner_name, record.event_name)
            self._exact_index.setdefault(key, []).append(record)

    @staticmethod
    def _exact_key(partner_name: str, event_name: str) -> str:
        return f"{normalize_text(partner_name)}|||{normalize_text(event_name)}"

    def find_exact(self, partner_name: str, event_name: str) -> PartnerRecord | None:
        matches = self._exact_index.get(self._exact_key(partner_name, event_name), [])
        if not matches:
            return None
        if len(matches) > 1:
            self.logger.warning(
                "Найдено %d точных совпадений для пары (%s | %s). Использую первое.",
                len(matches),
                partner_name,
                event_name,
            )
        return matches[0]

    def _best_fuzzy_candidate(self, partner_name: str, event_name: str) -> MatchCandidate | None:
        partner_norm = normalize_text(partner_name)
        event_norm = normalize_text(event_name)
        if not partner_norm and not event_norm:
            return None

        best_candidate: MatchCandidate | None = None
        for record in self.partner_records:
            partner_score = float(fuzz.WRatio(partner_norm, record.partner_name_norm))
            event_score = float(fuzz.WRatio(event_norm, record.event_name_norm))
            combined_score = round(partner_score * 0.45 + event_score * 0.55, 2)
            candidate = MatchCandidate(
                record=record,
                partner_score=partner_score,
                event_score=event_score,
                combined_score=combined_score,
            )
            if not best_candidate or candidate.combined_score > best_candidate.combined_score:
                best_candidate = candidate
        return best_candidate

    @staticmethod
    def _no_tz_record(source_partner: str, source_event: str) -> PartnerRecord:
        return PartnerRecord(
            row_index=-1,
            event_name=source_event,
            partner_name=source_partner,
            technical_requirements="",
        )

    @staticmethod
    def _table_cell_text(value: object, max_width: int = 72) -> str:
        if value is None:
            return "-"
        text = " ".join(str(value).split())
        if not text:
            return "-"
        if len(text) <= max_width:
            return text
        return text[: max_width - 3] + "..."

    @classmethod
    def _build_ascii_preview_table(cls, rows: list[tuple[str, str, str]]) -> str:
        header = ("Поле", "Нужно найти", "Предложено")
        sanitized_rows = [
            (
                cls._table_cell_text(row[0], max_width=24),
                cls._table_cell_text(row[1], max_width=72),
                cls._table_cell_text(row[2], max_width=72),
            )
            for row in rows
        ]
        widths = [len(header[0]), len(header[1]), len(header[2])]
        for row in sanitized_rows:
            widths[0] = max(widths[0], len(row[0]))
            widths[1] = max(widths[1], len(row[1]))
            widths[2] = max(widths[2], len(row[2]))

        def border() -> str:
            return "+" + "+".join("-" * (width + 2) for width in widths) + "+"

        def format_row(cells: tuple[str, str, str]) -> str:
            return "| " + " | ".join(cells[index].ljust(widths[index]) for index in range(3)) + " |"

        lines = [border(), format_row(header), border()]
        lines.extend(format_row(row) for row in sanitized_rows)
        lines.append(border())
        return "\n".join(lines)

    def _print_match_preview_table(self, source_partner: str, source_event: str, best: MatchCandidate) -> None:
        rows = [
            ("Партнёр", source_partner, best.record.partner_name),
            ("Мероприятие", source_event, best.record.event_name),
            (
                "Сходство",
                "-",
                f"партнёр={best.partner_score:.2f}; мероприятие={best.event_score:.2f}; итог={best.combined_score:.2f}",
            ),
        ]
        print("\nТребуется подтверждение сопоставления:")
        print(self._build_ascii_preview_table(rows))

    def _prompt_yes_no_skip_no_tz(self) -> str:
        while True:
            answer = input(
                "Подтвердить найденное совпадение? "
                "[Да/Нет/Пропустить/Вписать без ТЗ]: "
            ).strip().lower()
            if answer in {"да", "д", "yes", "y"}:
                return "yes"
            if answer in {"нет", "н", "no"}:
                return "no"
            if answer in {"пропустить", "п", "skip", "s"}:
                return "skip"
            if answer in {
                "вписать без тз",
                "без тз",
                "вписатьбезтз",
                "notz",
                "no_tz",
                "ntz",
            }:
                return "no_tz"
            print("Некорректный ответ. Введите: Да, Нет, Пропустить или Вписать без ТЗ.")

    def _prompt_exact_mapping(self, source_partner: str, source_event: str) -> PartnerRecord:
        while True:
            exact_partner = input("Введите точное имя партнёра из файла 'Партнёры': ").strip()
            exact_event = input("Введите точное название мероприятия из файла 'Партнёры': ").strip()
            record = self.find_exact(exact_partner, exact_event)
            if record:
                self.cache.set_mapping(
                    source_partner=source_partner,
                    source_event=source_event,
                    target_partner=record.partner_name,
                    target_event=record.event_name,
                )
                self.logger.info("Ручное сопоставление сохранено в кеш.")
                return record
            print("Точное совпадение не найдено в файле 'Партнёры'. Повторите ввод.")

    def resolve(self, source_partner: str, source_event: str) -> PartnerRecord | None:
        if not self.partner_records:
            raise ValueError("Файл 'Партнёры' пустой: невозможно выполнить сопоставление")

        cached_entry = self.cache.get(source_partner, source_event)
        if cached_entry:
            action = cached_entry.get("action", "")
            if action == "skip":
                self.logger.info("Кеш: пропуск пары (%s | %s)", source_partner, source_event)
                return None
            if action == "no_tz":
                self.logger.info(
                    "Кеш: использован режим 'Вписать без ТЗ' для пары (%s | %s)",
                    source_partner,
                    source_event,
                )
                return self._no_tz_record(source_partner, source_event)
            if action == "map":
                target_partner = cached_entry.get("partner", "")
                target_event = cached_entry.get("event", "")
                cached_match = self.find_exact(target_partner, target_event)
                if cached_match:
                    self.logger.info(
                        "Кеш: использовано сопоставление (%s | %s) -> (%s | %s)",
                        source_partner,
                        source_event,
                        target_partner,
                        target_event,
                    )
                    return cached_match
                self.logger.warning(
                    "Кеш-сопоставление устарело (%s | %s) -> (%s | %s), выполняю подбор заново.",
                    source_partner,
                    source_event,
                    target_partner,
                    target_event,
                )

        exact_match = self.find_exact(source_partner, source_event)
        if exact_match:
            self.logger.info("Точное совпадение найдено без запроса: (%s | %s)", source_partner, source_event)
            self.cache.set_mapping(
                source_partner=source_partner,
                source_event=source_event,
                target_partner=exact_match.partner_name,
                target_event=exact_match.event_name,
            )
            return exact_match

        best = self._best_fuzzy_candidate(source_partner, source_event)
        if not best:
            self.logger.warning("Не удалось найти кандидата для (%s | %s)", source_partner, source_event)
            self.cache.set_skip(source_partner, source_event)
            return None

        self.logger.info(
            "Лучшее fuzzy-совпадение: source=(%s | %s), candidate=(%s | %s), "
            "score_partner=%.2f, score_event=%.2f, score_total=%.2f",
            source_partner,
            source_event,
            best.record.partner_name,
            best.record.event_name,
            best.partner_score,
            best.event_score,
            best.combined_score,
        )

        if best.partner_score == 100 and best.event_score == 100:
            self.logger.info("Идеальное fuzzy-совпадение подтверждено автоматически.")
            self.cache.set_mapping(
                source_partner=source_partner,
                source_event=source_event,
                target_partner=best.record.partner_name,
                target_event=best.record.event_name,
            )
            return best.record

        self._print_match_preview_table(source_partner, source_event, best)

        decision = self._prompt_yes_no_skip_no_tz()
        if decision == "yes":
            self.cache.set_mapping(
                source_partner=source_partner,
                source_event=source_event,
                target_partner=best.record.partner_name,
                target_event=best.record.event_name,
            )
            self.logger.info("Выбор пользователя: Да. Сопоставление сохранено в кеш.")
            return best.record

        if decision == "no":
            self.logger.info("Выбор пользователя: Нет. Запрашиваю ручной ввод.")
            return self._prompt_exact_mapping(source_partner, source_event)

        if decision == "no_tz":
            self.logger.info(
                "Выбор пользователя: Вписать без ТЗ. Будет сохранено в кеш и ТЗ оставлено пустым."
            )
            self.cache.set_no_tz(source_partner, source_event)
            return self._no_tz_record(source_partner, source_event)

        self.logger.info("Выбор пользователя: Пропустить. Значение сохранено в кеш.")
        self.cache.set_skip(source_partner, source_event)
        return None
