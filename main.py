"""PDF Tag Extractor — Application Entry Point.

Launches the CustomTkinter GUI for extracting engineering tags
from one-line diagram PDFs.

Usage:
    python main.py
"""

import logging
import multiprocessing
import sys
from pathlib import Path

from gui.app import TagExtractorApp


def _setup_logging() -> None:
    """Configure the application logging system.

    Writes detailed logs to a file in the `logs/` directory and
    outputs INFO-level messages to the console.
    """
    log_dir = Path(__file__).parent / "logs"
    log_dir.mkdir(exist_ok=True)
    log_file = log_dir / "tag_extractor.log"

    # Root logger configuration
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)

    # File handler — DEBUG level, detailed format
    file_handler = logging.FileHandler(
        str(log_file), mode="a", encoding="utf-8"
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(
        logging.Formatter(
            fmt="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
    )

    # Console handler — INFO level, concise format
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(
        logging.Formatter(
            fmt="%(levelname)-8s | %(message)s",
        )
    )

    root_logger.addHandler(file_handler)
    root_logger.addHandler(console_handler)


def main() -> None:
    """Initialize logging and launch the GUI application."""
    _setup_logging()
    logger = logging.getLogger(__name__)
    logger.info("Iniciando PDF Tag Extractor v1.0")

    app = TagExtractorApp()
    app.mainloop()

    logger.info("Aplicação encerrada.")


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()