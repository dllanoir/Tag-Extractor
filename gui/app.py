"""PDF Tag Extractor — Modern GUI Application.

CustomTkinter-based graphical interface with dark/light mode toggle,
file selection dialogs, progress tracking, data table visualization,
and threaded PDF extraction to prevent UI freezing.
"""

import logging
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from core.exporter import DataExporter
from core.extractor import PdfTagExtractor
from core.models import ExtractionConfig, TagRecord

logger = logging.getLogger(__name__)

# ─── Theme Constants ───────────────────────────────────────────────────────────
_ACCENT = "#0EA5E9"
_ACCENT_HOVER = "#0284C7"
_ACCENT_DARK = "#075985"
_SUCCESS = "#22C55E"
_SUCCESS_DARK = "#15803D"
_ERROR = "#EF4444"
_ERROR_DARK = "#DC2626"
_TABLE_BG_DARK = "#1E293B"
_TABLE_FG_DARK = "#E2E8F0"
_TABLE_SELECTED_DARK = "#334155"
_TABLE_STRIPE_DARK = "#253347"
_TABLE_BG_LIGHT = "#FFFFFF"
_TABLE_FG_LIGHT = "#1E293B"
_TABLE_SELECTED_LIGHT = "#DBEAFE"
_TABLE_STRIPE_LIGHT = "#F1F5F9"
_TABLE_HEADING_BG_LIGHT = "#0284C7"
_SECONDARY_TEXT_DARK = "#94A3B8"
_SECONDARY_TEXT_LIGHT = "#64748B"
_FONT_FAMILY = "Segoe UI"


class TagExtractorApp(ctk.CTk):
    """Main application window for the PDF Tag Extractor.

    Provides a complete workflow: select PDF → extract → preview → export.
    All extraction runs on a background thread to keep the UI responsive.
    """

    def __init__(self) -> None:
        super().__init__()

        # ── Window Setup ──
        self.title("PDF Tag Extractor — Engineering Diagrams")
        self.geometry("1100x750")
        self.minsize(900, 650)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # ── Window Icon ──
        icon_path = self._resolve_asset("fav.ico")
        if icon_path.exists():
            self.iconbitmap(str(icon_path))
            self.after(200, lambda: self.iconbitmap(str(icon_path)))

        # ── State ──
        self._pdf_path: str = ""
        self._output_dir: str = str(Path.home() / "Desktop")
        self._records: list[TagRecord] = []
        self._format_var = ctk.StringVar(value=".xlsx")
        self._is_dark = True

        # ── Build UI ──
        self._build_header()
        self._build_input_section()
        self._build_output_section()
        self._build_action_bar()
        self._build_progress_section()
        self._build_table()
        self._build_footer()

        # ── Grid Weights ──
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(6, weight=1)  # Table row expands

    # ═══════════════════════════════════════════════════════════════════════════
    # UI CONSTRUCTION
    # ═══════════════════════════════════════════════════════════════════════════

    def _build_header(self) -> None:
        """Build the title bar with app name and theme toggle."""
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=20, pady=(15, 5))
        header.grid_columnconfigure(1, weight=1)

        title = ctk.CTkLabel(
            header,
            text="⚡ PDF Tag Extractor",
            font=ctk.CTkFont(family=_FONT_FAMILY, size=22, weight="bold"),
        )
        title.grid(row=0, column=0, sticky="w")

        subtitle = ctk.CTkLabel(
            header,
            text="Extração automática de tags de diagramas de engenharia",
            font=ctk.CTkFont(family=_FONT_FAMILY, size=12),
            text_color=_SECONDARY_TEXT_DARK,
        )
        subtitle.grid(row=1, column=0, sticky="w", pady=(0, 0))
        self._subtitle = subtitle

        self._theme_btn = ctk.CTkButton(
            header,
            text="☀️ Light",
            width=90,
            height=30,
            corner_radius=15,
            fg_color="transparent",
            border_width=1,
            border_color="gray",
            hover_color=_ACCENT_DARK,
            font=ctk.CTkFont(family=_FONT_FAMILY, size=12),
            command=self._toggle_theme,
        )
        self._theme_btn.grid(row=0, column=2, rowspan=2, sticky="e", padx=(10, 0))

    def _build_input_section(self) -> None:
        """Build the PDF file selection section."""
        frame = ctk.CTkFrame(self, corner_radius=10)
        frame.grid(row=1, column=0, sticky="ew", padx=20, pady=(10, 5))
        frame.grid_columnconfigure(1, weight=1)

        label = ctk.CTkLabel(
            frame,
            text="📄 Arquivo PDF:",
            font=ctk.CTkFont(family=_FONT_FAMILY, size=13, weight="bold"),
        )
        label.grid(row=0, column=0, padx=15, pady=12, sticky="w")

        self._pdf_entry = ctk.CTkEntry(
            frame,
            placeholder_text="Selecione um arquivo PDF...",
            state="disabled",
            font=ctk.CTkFont(family=_FONT_FAMILY, size=12),
            height=36,
        )
        self._pdf_entry.grid(row=0, column=1, padx=(0, 10), pady=12, sticky="ew")

        btn = ctk.CTkButton(
            frame,
            text="📁 Selecionar",
            width=130,
            height=36,
            corner_radius=8,
            fg_color=_ACCENT,
            hover_color=_ACCENT_HOVER,
            font=ctk.CTkFont(family=_FONT_FAMILY, size=13, weight="bold"),
            command=self._select_pdf,
        )
        btn.grid(row=0, column=2, padx=(0, 15), pady=12)

    def _build_output_section(self) -> None:
        """Build the output format and directory selection section."""
        frame = ctk.CTkFrame(self, corner_radius=10)
        frame.grid(row=2, column=0, sticky="ew", padx=20, pady=5)
        frame.grid_columnconfigure(2, weight=1)

        # Format selection
        fmt_label = ctk.CTkLabel(
            frame,
            text="📋 Formato:",
            font=ctk.CTkFont(family=_FONT_FAMILY, size=13, weight="bold"),
        )
        fmt_label.grid(row=0, column=0, padx=(15, 5), pady=12, sticky="w")

        radio_frame = ctk.CTkFrame(frame, fg_color="transparent")
        radio_frame.grid(row=0, column=1, padx=(0, 20), pady=12, sticky="w")

        ctk.CTkRadioButton(
            radio_frame,
            text="Excel (.xlsx)",
            variable=self._format_var,
            value=".xlsx",
            font=ctk.CTkFont(family=_FONT_FAMILY, size=12),
            radiobutton_width=18,
            radiobutton_height=18,
        ).pack(side="left", padx=(0, 15))

        ctk.CTkRadioButton(
            radio_frame,
            text="Texto (.txt)",
            variable=self._format_var,
            value=".txt",
            font=ctk.CTkFont(family=_FONT_FAMILY, size=12),
            radiobutton_width=18,
            radiobutton_height=18,
        ).pack(side="left")

        # Output directory
        self._dir_entry = ctk.CTkEntry(
            frame,
            placeholder_text="Pasta de saída...",
            state="disabled",
            font=ctk.CTkFont(family=_FONT_FAMILY, size=12),
            height=36,
        )
        self._dir_entry.grid(row=0, column=2, padx=(0, 10), pady=12, sticky="ew")
        self._set_entry_text(self._dir_entry, self._output_dir)

        dir_btn = ctk.CTkButton(
            frame,
            text="📂 Pasta",
            width=100,
            height=36,
            corner_radius=8,
            fg_color=_ACCENT,
            hover_color=_ACCENT_HOVER,
            font=ctk.CTkFont(family=_FONT_FAMILY, size=13, weight="bold"),
            command=self._select_output_dir,
        )
        dir_btn.grid(row=0, column=3, padx=(0, 15), pady=12)

    def _build_action_bar(self) -> None:
        """Build the main action button."""
        self._extract_btn = ctk.CTkButton(
            self,
            text="⚙️  EXTRAIR TAGS",
            height=44,
            corner_radius=10,
            fg_color=_ACCENT,
            hover_color=_ACCENT_HOVER,
            font=ctk.CTkFont(family=_FONT_FAMILY, size=15, weight="bold"),
            command=self._start_extraction,
        )
        self._extract_btn.grid(row=3, column=0, sticky="ew", padx=20, pady=(10, 5))

    def _build_progress_section(self) -> None:
        """Build the progress bar and status label."""
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.grid(row=4, column=0, sticky="ew", padx=20, pady=(0, 5))
        frame.grid_columnconfigure(0, weight=1)

        self._progress_bar = ctk.CTkProgressBar(
            frame, height=6, corner_radius=3, progress_color=_ACCENT
        )
        self._progress_bar.grid(row=0, column=0, sticky="ew", pady=(5, 2))
        self._progress_bar.set(0)

        self._status_label = ctk.CTkLabel(
            frame,
            text="Pronto para iniciar.",
            font=ctk.CTkFont(family=_FONT_FAMILY, size=11),
            text_color=_SECONDARY_TEXT_DARK,
        )
        self._status_label.grid(row=1, column=0, sticky="w")

    def _build_table(self) -> None:
        """Build the data preview table using ttk.Treeview with styling."""
        table_frame = ctk.CTkFrame(self, corner_radius=10)
        table_frame.grid(row=6, column=0, sticky="nsew", padx=20, pady=5)
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)

        # Style the Treeview for dark mode
        self._table_style = ttk.Style()
        self._apply_table_theme()

        columns = ("page", "area", "subarea", "tag", "location")
        self._tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            style="Custom.Treeview",
            selectmode="browse",
        )

        # Column configuration
        self._tree.heading("page", text="Página", anchor="center")
        self._tree.heading("area", text="Área", anchor="w")
        self._tree.heading("subarea", text="Subárea", anchor="w")
        self._tree.heading("tag", text="Tag", anchor="w")
        self._tree.heading("location", text="Local", anchor="w")

        self._tree.column("page", width=60, minwidth=50, anchor="center")
        self._tree.column("area", width=130, minwidth=80, anchor="w")
        self._tree.column("subarea", width=300, minwidth=150, anchor="w")
        self._tree.column("tag", width=180, minwidth=120, anchor="w")
        self._tree.column("location", width=220, minwidth=120, anchor="w")

        self._tree.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=10)

        # Scrollbar
        scrollbar = ttk.Scrollbar(
            table_frame, orient="vertical", command=self._tree.yview
        )
        self._tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=0, column=1, sticky="ns", padx=(0, 10), pady=10)

    def _build_footer(self) -> None:
        """Build the footer with record count and export button."""
        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.grid(row=7, column=0, sticky="ew", padx=20, pady=(5, 15))
        footer.grid_columnconfigure(0, weight=1)

        self._count_label = ctk.CTkLabel(
            footer,
            text="Nenhum registro carregado.",
            font=ctk.CTkFont(family=_FONT_FAMILY, size=12),
            text_color=_SECONDARY_TEXT_DARK,
        )
        self._count_label.grid(row=0, column=0, sticky="w")

        self._export_btn = ctk.CTkButton(
            footer,
            text="💾  EXPORTAR",
            width=160,
            height=38,
            corner_radius=8,
            fg_color=_SUCCESS,
            hover_color="#16A34A",
            font=ctk.CTkFont(family=_FONT_FAMILY, size=13, weight="bold"),
            state="disabled",
            command=self._export_data,
        )
        self._export_btn.grid(row=0, column=1, sticky="e")

    # ═══════════════════════════════════════════════════════════════════════════
    # THEME MANAGEMENT
    # ═══════════════════════════════════════════════════════════════════════════

    def _toggle_theme(self) -> None:
        """Toggle between dark and light appearance modes."""
        self._is_dark = not self._is_dark
        secondary = _SECONDARY_TEXT_DARK if self._is_dark else _SECONDARY_TEXT_LIGHT

        if self._is_dark:
            ctk.set_appearance_mode("dark")
            self._theme_btn.configure(
                text="☀️ Light",
                border_color="gray",
                hover_color=_ACCENT_DARK,
                text_color="white",
            )
        else:
            ctk.set_appearance_mode("light")
            self._theme_btn.configure(
                text="🌙 Dark",
                border_color="#CBD5E1",
                hover_color="#E2E8F0",
                text_color="#1E293B",
            )

        # Update secondary-text labels to match the new theme
        self._subtitle.configure(text_color=secondary)
        self._status_label.configure(text_color=secondary)
        self._count_label.configure(text_color=secondary)

        self._apply_table_theme()

    def _apply_table_theme(self) -> None:
        """Apply Treeview styling matching the current theme."""
        if self._is_dark:
            bg, fg, sel, stripe = (
                _TABLE_BG_DARK,
                _TABLE_FG_DARK,
                _TABLE_SELECTED_DARK,
                _TABLE_STRIPE_DARK,
            )
            heading_bg = _ACCENT_DARK
        else:
            bg, fg, sel, stripe = (
                _TABLE_BG_LIGHT,
                _TABLE_FG_LIGHT,
                _TABLE_SELECTED_LIGHT,
                _TABLE_STRIPE_LIGHT,
            )
            heading_bg = _TABLE_HEADING_BG_LIGHT

        self._table_style.theme_use("default")
        self._table_style.configure(
            "Custom.Treeview",
            background=bg,
            foreground=fg,
            fieldbackground=bg,
            font=(_FONT_FAMILY, 11),
            rowheight=28,
            borderwidth=0,
        )
        self._table_style.configure(
            "Custom.Treeview.Heading",
            background=heading_bg,
            foreground="white",
            font=(_FONT_FAMILY, 11, "bold"),
            borderwidth=0,
            relief="flat",
        )
        self._table_style.map(
            "Custom.Treeview",
            background=[("selected", sel)],
            foreground=[("selected", fg)],
        )

        # Store stripe color for zebra rows
        self._stripe_color = stripe
        self._row_bg = bg

    # ═══════════════════════════════════════════════════════════════════════════
    # FILE SELECTION
    # ═══════════════════════════════════════════════════════════════════════════

    def _select_pdf(self) -> None:
        """Open file dialog to select a PDF file."""
        path = filedialog.askopenfilename(
            title="Selecionar PDF de Diagrama",
            filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")],
        )
        if path:
            self._pdf_path = path
            self._set_entry_text(self._pdf_entry, path)
            logger.info("PDF selecionado: %s", path)

    def _select_output_dir(self) -> None:
        """Open directory dialog to choose the export location."""
        directory = filedialog.askdirectory(
            title="Selecionar Pasta de Saída",
            initialdir=self._output_dir,
        )
        if directory:
            self._output_dir = directory
            self._set_entry_text(self._dir_entry, directory)
            logger.info("Pasta de saída: %s", directory)

    # ═══════════════════════════════════════════════════════════════════════════
    # EXTRACTION (THREADED)
    # ═══════════════════════════════════════════════════════════════════════════

    def _start_extraction(self) -> None:
        """Validate inputs and launch the extraction on a background thread."""
        if not self._pdf_path:
            messagebox.showwarning(
                "Atenção", "Por favor, selecione um arquivo PDF primeiro."
            )
            return

        # Disable UI during processing
        self._extract_btn.configure(state="disabled", text="⏳  PROCESSANDO...")
        self._export_btn.configure(state="disabled")
        self._progress_bar.set(0)
        self._status_label.configure(
            text="Iniciando extração...",
            text_color=_ACCENT_DARK if not self._is_dark else _ACCENT,
        )
        self._clear_table()

        thread = threading.Thread(target=self._run_extraction, daemon=True)
        thread.start()

    def _run_extraction(self) -> None:
        """Execute PDF extraction on a background thread.

        Uses `self.after()` to safely update the GUI from the worker thread.
        """
        try:
            config = ExtractionConfig()
            extractor = PdfTagExtractor(config)

            def on_progress(current: int, total: int) -> None:
                progress = current / total
                self.after(0, self._update_progress, progress, current, total)

            records = extractor.extract(self._pdf_path, progress_callback=on_progress)
            self.after(0, self._on_extraction_complete, records)

        except FileNotFoundError as exc:
            self.after(0, self._on_extraction_error, str(exc))
        except RuntimeError as exc:
            self.after(0, self._on_extraction_error, str(exc))
        except Exception as exc:
            logger.exception("Erro inesperado na extração")
            self.after(0, self._on_extraction_error, f"Erro inesperado: {exc}")

    def _update_progress(self, value: float, current: int, total: int) -> None:
        """Update the progress bar and status label (called on main thread)."""
        self._progress_bar.set(value)
        self._status_label.configure(
            text=f"Processando página {current} de {total}...",
            text_color=_ACCENT_DARK if not self._is_dark else _ACCENT,
        )

    def _on_extraction_complete(self, records: list[TagRecord]) -> None:
        """Handle successful extraction completion (called on main thread)."""
        self._records = records
        self._progress_bar.set(1.0)

        # Populate table
        self._populate_table(records)

        count = len(records)
        success_color = _SUCCESS_DARK if not self._is_dark else _SUCCESS
        error_color = _ERROR_DARK if not self._is_dark else _ERROR
        self._status_label.configure(
            text=f"✅ Extração concluída — {count} tags encontradas.",
            text_color=success_color,
        )
        self._count_label.configure(
            text=f"📊 {count} registros carregados.",
            text_color=success_color if count > 0 else error_color,
        )
        self._extract_btn.configure(state="normal", text="⚙️  EXTRAIR TAGS")
        self._export_btn.configure(state="normal" if count > 0 else "disabled")
        logger.info("Extração finalizada: %d registros.", count)

    def _on_extraction_error(self, message: str) -> None:
        """Handle extraction failure (called on main thread)."""
        self._progress_bar.set(0)
        error_color = _ERROR_DARK if not self._is_dark else _ERROR
        self._status_label.configure(text=f"❌ {message}", text_color=error_color)
        self._extract_btn.configure(state="normal", text="⚙️  EXTRAIR TAGS")
        messagebox.showerror("Erro na Extração", message)

    # ═══════════════════════════════════════════════════════════════════════════
    # TABLE MANAGEMENT
    # ═══════════════════════════════════════════════════════════════════════════

    def _clear_table(self) -> None:
        """Remove all rows from the Treeview in a single batch operation."""
        children = self._tree.get_children()
        if children:
            self._tree.delete(*children)

    def _populate_table(self, records: list[TagRecord]) -> None:
        """Insert all records into the Treeview with zebra striping.

        Performance: the Treeview is temporarily hidden during bulk inserts
        to prevent individual re-renders per row (critical for 1000+ records).
        """
        # Hide widget to suppress per-insert rendering
        self._tree.grid_remove()

        self._clear_table()

        self._tree.tag_configure("stripe", background=self._stripe_color)
        self._tree.tag_configure("normal", background=self._row_bg)

        for idx, record in enumerate(records):
            tag = "stripe" if idx % 2 == 1 else "normal"
            self._tree.insert(
                "",
                "end",
                values=(record.page, record.area, record.subarea, record.tag, record.location),
                tags=(tag,),
            )

        # Re-show widget — single render pass for all rows
        self._tree.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=10)

    # ═══════════════════════════════════════════════════════════════════════════
    # EXPORT
    # ═══════════════════════════════════════════════════════════════════════════

    def _export_data(self) -> None:
        """Export the current records to the selected format and directory."""
        if not self._records:
            messagebox.showwarning("Atenção", "Nenhum dado para exportar.")
            return

        fmt = self._format_var.get()
        pdf_name = Path(self._pdf_path).stem
        output_file = Path(self._output_dir) / f"{pdf_name}_tags{fmt}"

        try:
            exporter = DataExporter()
            if fmt == ".txt":
                exporter.export_txt(self._records, output_file)
            else:
                exporter.export_xlsx(self._records, output_file)

            self._status_label.configure(
                text=f"💾 Exportado: {output_file.name}",
                text_color=_SUCCESS_DARK if not self._is_dark else _SUCCESS,
            )
            messagebox.showinfo(
                "Exportação Concluída",
                f"Arquivo salvo com sucesso:\n{output_file}",
            )
        except IOError as exc:
            messagebox.showerror("Erro na Exportação", str(exc))

    # ═══════════════════════════════════════════════════════════════════════════
    # UTILITIES
    # ═══════════════════════════════════════════════════════════════════════════

    @staticmethod
    def _set_entry_text(entry: ctk.CTkEntry, text: str) -> None:
        """Set text on a disabled CTkEntry widget."""
        entry.configure(state="normal")
        entry.delete(0, "end")
        entry.insert(0, text)
        entry.configure(state="disabled")

    @staticmethod
    def _resolve_asset(filename: str) -> Path:
        """Resolve an asset path for both dev and PyInstaller frozen modes.

        When running as a PyInstaller bundle, assets are extracted to a
        temporary directory accessible via sys._MEIPASS. In dev mode,
        assets are resolved relative to the project root.

        Args:
            filename: Asset filename (e.g., 'fav.ico').

        Returns:
            Absolute Path to the asset file.
        """
        if getattr(sys, "frozen", False):
            # Running as PyInstaller bundle
            base = Path(sys._MEIPASS)  # noqa: SLF001
        else:
            # Running in development
            base = Path(__file__).resolve().parent.parent
        return base / filename
