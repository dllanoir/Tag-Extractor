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
        raw_words = page.extract_words(extra_attrs=["size"])
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

        This method handles header classification and tag matching. It receives
        already-extracted and sorted words (from the parallel phase).

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

        current_line_y = -1
        accumulated_line_text = ""

        for word_data in words:
            text = word_data["text"].strip()
            size = word_data["size"]
            x0 = word_data["x0"]
            y0 = round(word_data["top"])

            # --- Line tracker: group words on the same visual line ---
            if y0 != current_line_y:
                current_line_y = y0
                accumulated_line_text = text
            else:
                accumulated_line_text += " " + text

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

                records.append(
                    TagRecord(
                        page=page_number,
                        area=current_area,
                        subarea=current_subarea,
                        tag=text,
                    )
                )

        return records, current_area, current_subarea, area_buffer, subarea_buffer
