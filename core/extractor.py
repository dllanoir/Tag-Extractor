"""PDF Tag Extraction Engine.

Reads engineering one-line diagram PDFs using pdfplumber, extracts equipment
tags via regex, and groups them under spatial headers determined by font size
and position. Uses parallel word extraction for performance.
"""

import logging
import re
from collections.abc import Callable
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

import pdfplumber

from core.models import ExtractionConfig, TagRecord

logger = logging.getLogger(__name__)

# Type alias for the progress callback: (current_page, total_pages) -> None
ProgressCallback = Callable[[int, int], None]


def _extract_words_from_page(
    pdf_path: str, page_index: int
) -> tuple[int, list[dict]]:
    """Extract words from a single PDF page (thread-safe).

    Each thread opens its own PDF handle to avoid shared-state issues
    with pdfplumber/pdfminer internals.

    Args:
        pdf_path: Path to the PDF file.
        page_index: 0-indexed page number.

    Returns:
        Tuple of (page_index, sorted_words_list).
    """
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_index]
        raw_words = page.extract_words(extra_attrs=["size", "fontname"])
    # Sort by visual line (top) then horizontal position (x0)
    sorted_words = sorted(raw_words, key=lambda w: (round(w["top"]), w["x0"]))
    return page_index, sorted_words


class PdfTagExtractor:
    """Extracts engineering tags from PDF one-line diagrams.

    Uses a two-phase approach for performance:
      Phase 1 (parallel): Extract raw words from all pages concurrently
      Phase 2 (sequential): Process headers and tags maintaining state order

    Usage:
        config = ExtractionConfig()
        extractor = PdfTagExtractor(config)
        records = extractor.extract("diagram.pdf")
    """

    def __init__(self, config: ExtractionConfig | None = None) -> None:
        """Initialize the extractor with optional configuration.

        Args:
            config: Extraction parameters. Uses defaults if not provided.
        """
        self._config = config or ExtractionConfig()
        self._tag_regex = re.compile(self._config.tag_pattern)
        self._exclusion_pattern = self._build_exclusion_pattern()

    def _build_exclusion_pattern(self) -> re.Pattern[str]:
        """Build a compiled regex for exclusion keywords.

        Returns:
            Compiled regex that matches any exclusion keyword as a whole word.
        """
        keywords = "|".join(
            re.escape(kw) for kw in self._config.exclusion_keywords
        )
        return re.compile(rf"\b({keywords})\b", re.IGNORECASE)

    def extract(
        self,
        pdf_path: str | Path,
        progress_callback: ProgressCallback | None = None,
    ) -> list[TagRecord]:
        """Extract all engineering tags from the given PDF file.

        Phase 1: Parallel word extraction from all pages using ThreadPoolExecutor.
        Phase 2: Sequential tag/header processing to maintain area/subarea state.

        Args:
            pdf_path: Absolute or relative path to the PDF file.
            progress_callback: Optional callable invoked with
                (current_step, total_steps) for progress tracking.

        Returns:
            List of TagRecord objects with page, area, subarea, and tag data.

        Raises:
            FileNotFoundError: If the PDF file does not exist.
            RuntimeError: If pdfplumber fails to open the file.
        """
        pdf_path = Path(pdf_path)
        if not pdf_path.exists():
            raise FileNotFoundError(f"PDF não encontrado: {pdf_path}")

        logger.info("Iniciando extração do PDF: %s", pdf_path.name)

        # Quick open to get page count
        try:
            with pdfplumber.open(str(pdf_path)) as pdf:
                total_pages = len(pdf.pages)
        except Exception as exc:
            logger.error("Erro ao abrir PDF: %s", exc)
            raise RuntimeError(f"PDF inválido: {exc}") from exc

        logger.info("Total de páginas: %d", total_pages)

        # ── Phase 1: Parallel word extraction ──────────────────────────────
        max_workers = min(4, total_pages)
        logger.info("Fase 1: Extraindo palavras em paralelo (%d processos)...", max_workers)
        all_page_words: list[list[dict]] = [[] for _ in range(total_pages)]
        completed_count = 0

        try:
            with ProcessPoolExecutor(max_workers=max_workers) as executor:
                futures = {
                    executor.submit(
                        _extract_words_from_page, str(pdf_path), i
                    ): i
                    for i in range(total_pages)
                }

                for future in as_completed(futures):
                    page_idx, words = future.result()
                    all_page_words[page_idx] = words
                    completed_count += 1

                    if progress_callback:
                        progress_callback(completed_count, total_pages)

                    logger.debug(
                        "Página %d: %d palavras extraídas",
                        page_idx + 1,
                        len(words),
                    )

        except Exception as exc:
            logger.error("Erro na extração paralela: %s", exc)
            raise RuntimeError(f"Falha na extração: {exc}") from exc

        # ── Phase 2: Sequential header/tag processing ──────────────────────
        logger.info("Fase 2: Processando cabeçalhos e tags sequencialmente...")
        records: list[TagRecord] = []
        current_area = "Geral"
        current_subarea = "Geral"
        area_buffer: list[str] = []
        subarea_buffer: list[str] = []

        for page_idx in range(total_pages):
            page_number = page_idx + 1

            page_records, current_area, current_subarea, area_buffer, subarea_buffer = (
                self._process_page_words(
                    words=all_page_words[page_idx],
                    page_number=page_number,
                    current_area=current_area,
                    current_subarea=current_subarea,
                    area_buffer=area_buffer,
                    subarea_buffer=subarea_buffer,
                )
            )
            records.extend(page_records)

            logger.debug(
                "Página %d: %d tags encontradas",
                page_number,
                len(page_records),
            )

        logger.info(
            "Extração concluída: %d registros extraídos de %d páginas.",
            len(records),
            total_pages,
        )
        return records

    def _process_page_words(
        self,
        words: list[dict],
        page_number: int,
        current_area: str,
        current_subarea: str,
        area_buffer: list[str],
        subarea_buffer: list[str],
    ) -> tuple[list[TagRecord], str, str, list[str], list[str]]:
        """Process pre-extracted words from a single page and extract tags.

        This method handles header classification, tag matching, and location
        extraction. It groups words into visual lines, then for each tag found,
        looks 1-2 lines above for nearby text that represents the tag's
        physical installation location.

        Args:
            words: Pre-extracted and sorted word dicts from pdfplumber.
            page_number: 1-indexed page number.
            current_area: Current area text (e.g., 'TOPSIDE').
            current_subarea: Current sub-area text (e.g., 'M-06 - ...').
            area_buffer: Accumulator for multi-word area headers.
            subarea_buffer: Accumulator for multi-word sub-area headers.

        Returns:
            Tuple of (page_records, updated_area, updated_subarea,
            area_buffer, subarea_buffer).
        """
        config = self._config
        records: list[TagRecord] = []

        # ── Step 1: Group words into visual lines ──────────────────────────
        visual_lines: list[tuple[int, list[dict]]] = []
        current_line_y = -1
        current_line_words: list[dict] = []

        for word_data in words:
            y0 = round(word_data["top"])
            if y0 != current_line_y:
                if current_line_words:
                    visual_lines.append((current_line_y, current_line_words))
                current_line_y = y0
                current_line_words = [word_data]
            else:
                current_line_words.append(word_data)

        if current_line_words:
            visual_lines.append((current_line_y, current_line_words))

        # ── Step 1.5: Pre-calculate levels for this page ───────────────────
        levels = self._extract_levels_from_page(visual_lines)

        # ── Step 2: Process headers and tags line by line ──────────────────
        for line_idx, (y0, line_words) in enumerate(visual_lines):
            accumulated_line_text = " ".join(w["text"].strip() for w in line_words)

            for word_data in line_words:
                text = word_data["text"].strip()
                size = word_data["size"]
                x0 = word_data["x0"]

                # --- Zone and header classification ---
                in_central_zone = config.zone_min_x <= x0 <= config.zone_max_x

                h1_min, h1_max = config.header1_size_range
                h2_min, h2_max = config.header2_size_range

                is_header_1 = (h1_min <= size <= h1_max) and in_central_zone
                is_header_2 = (h2_min <= size <= h2_max) and in_central_zone

                # --- Header buffer logic (identical to original algorithm) ---
                if is_header_1:
                    if subarea_buffer:
                        current_subarea = " ".join(subarea_buffer)
                        subarea_buffer = []
                    area_buffer.append(text)
                    continue

                if is_header_2:
                    if area_buffer:
                        current_area = " ".join(area_buffer)
                        area_buffer = []
                    subarea_buffer.append(text)
                    continue

                # Normal text — flush pending header buffers
                if area_buffer:
                    current_area = " ".join(area_buffer)
                    area_buffer = []
                if subarea_buffer:
                    current_subarea = " ".join(subarea_buffer)
                    subarea_buffer = []

                # --- Tag matching ---
                if self._tag_regex.fullmatch(text):
                    # Exclusion guard: skip tags on lines with FROM/TO
                    if self._exclusion_pattern.search(accumulated_line_text):
                        logger.debug(
                            "Tag '%s' excluída (linha contém keyword de exclusão)",
                            text,
                        )
                        continue

                    # --- Location extraction ---
                    location = self._find_location_above(
                        tag_word=word_data,
                        visual_lines=visual_lines,
                        tag_line_idx=line_idx,
                    )

                    # --- Level mapping ---
                    level = self._find_level_for_tag(y0, levels)

                    records.append(
                        TagRecord(
                            page=page_number,
                            area=current_area,
                            subarea=current_subarea,
                            tag=text,
                            location=location,
                            level=level,
                        )
                    )

        return records, current_area, current_subarea, area_buffer, subarea_buffer

    @staticmethod
    def _is_bold(word: dict) -> bool:
        """Check if a PDF word is rendered in a bold font.

        Args:
            word: A pdfplumber word dict with 'fontname' attribute.

        Returns:
            True if the font name contains 'Bold'.
        """
        fontname = word.get("fontname", "")
        return "Bold" in fontname or "bold" in fontname

    def _find_location_above(
        self,
        tag_word: dict,
        visual_lines: list[tuple[int, list[dict]]],
        tag_line_idx: int,
    ) -> str:
        """Find the physical location text above a tag by spatial proximity.

        Iterates upward through all visual lines within ``location_y_max_distance``
        and collects **bold** words whose horizontal center falls within the
        tag's x-range (plus padding).  This handles multi-line location labels
        (e.g., 'TELECOM CONTROL' on one line, 'ROOM (A707)' on the next) even
        when unrelated visual lines exist in between.

        Args:
            tag_word: The word dict for the matched tag.
            visual_lines: All visual lines on the page as (y, words) tuples.
            tag_line_idx: Index of the tag's line in visual_lines.

        Returns:
            Location string, or empty string if no location was found.
        """
        config = self._config
        tag_x0 = tag_word["x0"]
        tag_x1 = tag_word.get("x1", tag_x0 + 60)
        tag_y = visual_lines[tag_line_idx][0]

        # Search range: tag's full horizontal span + padding on each side
        search_x0 = tag_x0 - config.location_x_tolerance
        search_x1 = tag_x1 + config.location_x_tolerance

        location_parts: list[tuple[float, str]] = []  # (y_distance, text)

        for prev_idx in range(tag_line_idx - 1, -1, -1):
            prev_y, prev_words = visual_lines[prev_idx]
            y_distance = tag_y - prev_y

            if y_distance > config.location_y_max_distance or y_distance < 0:
                break

            # Collect bold words whose center falls within the search range
            nearby_bold: list[dict] = [
                pw for pw in prev_words
                if self._is_bold(pw)
                and search_x0
                <= (pw["x0"] + pw.get("x1", pw["x0"] + 20)) / 2
                <= search_x1
            ]

            if nearby_bold:
                nearby_bold.sort(key=lambda w: w["x0"])
                line_text = " ".join(w["text"].strip() for w in nearby_bold)
                location_parts.append((y_distance, line_text))

        if not location_parts:
            return ""

        # Build location string: lines further above (larger y_distance) first
        location_parts.sort(key=lambda p: p[0], reverse=True)
        return " ".join(part[1] for part in location_parts)

    def _extract_levels_from_page(
        self, visual_lines: list[tuple[int, list[dict]]]
    ) -> list[tuple[int, str]]:
        """Find all level/deck markers on the page's far-left side.

        Returns:
            A list of tuples (y_position, level_text), sorted by y_position.
        """
        levels = []
        config = self._config
        for y, words in visual_lines:
            # Reconstruct any marker text from the left side with the correct font size
            left_words = [
                w for w in words
                if w["x0"] < config.level_x_max
                and config.level_size_min <= w["size"] <= config.level_size_max
            ]
            
            if left_words:
                left_words.sort(key=lambda w: w["x0"])
                level_text = " ".join(w["text"].strip() for w in left_words)
                if level_text.strip():
                    levels.append((y, level_text))

        return levels

    def _find_level_for_tag(self, tag_y: int, levels: list[tuple[int, str]]) -> str:
        """Find the nearest level marker above the tag's y-position.

        Args:
            tag_y: The y-position (top) of the tag.
            levels: List of (y_position, level_text) tuples for the current page.

        Returns:
            The level text, or empty string if no valid level was found above.
        """
        best_level = ""
        best_y = -1

        for level_y, level_text in levels:
            # Level must be physically above or on the same line as the tag
            if level_y <= tag_y:
                if level_y > best_y:
                    best_y = level_y
                    best_level = level_text

        return best_level
