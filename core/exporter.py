"""Data Export Module.

Provides export functionality for extracted tag records in TXT (tab-separated)
and XLSX (Excel) formats using pandas and openpyxl.
"""

import logging
from pathlib import Path

import pandas as pd

from core.models import TagRecord

logger = logging.getLogger(__name__)

# Column names for the output files (matching original output format)
_COLUMNS = {
    "page": "PÁGINA",
    "area": "ÁREA",
    "subarea": "SUBÁREA",
    "tag": "TAG",
}


class DataExporter:
    """Exports TagRecord lists to TXT and XLSX file formats.

    Usage:
        exporter = DataExporter()
        df = exporter.to_dataframe(records)
        exporter.export_txt(records, "output.txt")
        exporter.export_xlsx(records, "output.xlsx")
    """

    @staticmethod
    def to_dataframe(records: list[TagRecord]) -> pd.DataFrame:
        """Convert a list of TagRecord objects to a pandas DataFrame.

        Args:
            records: List of extracted tag records.

        Returns:
            DataFrame with renamed columns matching the standard output format.
        """
        if not records:
            logger.warning("Nenhum registro para converter em DataFrame.")
            return pd.DataFrame(columns=list(_COLUMNS.values()))

        data = [
            {
                _COLUMNS["page"]: r.page,
                _COLUMNS["area"]: r.area,
                _COLUMNS["subarea"]: r.subarea,
                _COLUMNS["tag"]: r.tag,
            }
            for r in records
        ]
        return pd.DataFrame(data)

    @staticmethod
    def export_txt(records: list[TagRecord], output_path: str | Path) -> None:
        """Export records to a tab-separated text file.

        Args:
            records: List of extracted tag records.
            output_path: Destination file path (e.g., 'output.txt').

        Raises:
            IOError: If the file cannot be written.
        """
        output_path = Path(output_path)
        logger.info("Exportando %d registros para TXT: %s", len(records), output_path)

        try:
            df = DataExporter.to_dataframe(records)
            df.to_csv(
                str(output_path),
                sep="\t",
                index=False,
                encoding="utf-8",
            )
            logger.info("Exportação TXT concluída com sucesso.")
        except Exception as exc:
            logger.error("Erro ao exportar TXT: %s", exc)
            raise IOError(f"Falha ao salvar TXT: {exc}") from exc

    @staticmethod
    def export_xlsx(records: list[TagRecord], output_path: str | Path) -> None:
        """Export records to an Excel (.xlsx) file with formatting.

        Uses openpyxl engine. Applies auto-fit column widths for readability.

        Args:
            records: List of extracted tag records.
            output_path: Destination file path (e.g., 'output.xlsx').

        Raises:
            IOError: If the file cannot be written.
        """
        output_path = Path(output_path)
        logger.info("Exportando %d registros para XLSX: %s", len(records), output_path)

        try:
            df = DataExporter.to_dataframe(records)

            with pd.ExcelWriter(
                str(output_path), engine="openpyxl"
            ) as writer:
                df.to_excel(writer, index=False, sheet_name="Tags Extraídas")

                # Auto-fit column widths
                worksheet = writer.sheets["Tags Extraídas"]
                for col_idx, column_name in enumerate(df.columns, start=1):
                    max_length = max(
                        len(str(column_name)),
                        df[column_name].astype(str).str.len().max()
                        if not df.empty
                        else 0,
                    )
                    # Add padding and cap at 60 characters
                    adjusted_width = min(max_length + 4, 60)
                    col_letter = worksheet.cell(row=1, column=col_idx).column_letter
                    worksheet.column_dimensions[col_letter].width = adjusted_width

            logger.info("Exportação XLSX concluída com sucesso.")
        except Exception as exc:
            logger.error("Erro ao exportar XLSX: %s", exc)
            raise IOError(f"Falha ao salvar XLSX: {exc}") from exc
