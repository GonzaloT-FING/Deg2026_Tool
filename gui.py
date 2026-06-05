from __future__ import annotations

import pathlib
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from pipelines.activ_pip import run_pipeline as run_activ_pipeline
from pipelines.cic_vol_pip import run_pipeline as run_cv_pipeline
from pipelines.deg_pip import run_pipeline as run_deg_pipeline
from pipelines.eis_pip import run_pipeline as run_eis_pipeline
from pipelines.full_report_pip import generate_full_report
from pipelines.ocp_pip import run_pipeline as run_ocp_pipeline
from pipelines.pol_cur_pip import run_pipeline as run_pc_pipeline
from i18n import LANGUAGE_CODE_BY_LABEL, LANGUAGE_LABELS, translate
from plot_defaults import DEFAULT_PLOT_FONT_DEFAULTS, PlotFontDefaults, parse_plot_font_defaults
from ui_theme import THEME_LABELS, THEME_NAME_BY_LABEL, apply_theme


PIPELINE_OPTIONS = {
    "Activacion": ["V vs t", "V vs I", "dV/dI"],
    "OCP": ["V vs t", "DeltaV"],
    "CV": ["I vs V", "Peak current"],
    "EIS": ["Nyquist plot", "Bode plot", "Series by Pt", "Pre-estabilización", "MultiFit"],
    "PC": ["V vs I", "Series by time", "dV/dI", "Step Stability"],
    "Deg": ["V vs t", "dV/dt"],
    "Analisis multiple": ["EIS", "CV", "PC", "OCP", "Deg"],
}

PIPELINE_DESCRIPTIONS = {
    "Activacion": "Ciclos de activacion respecto al tiempo global o local y respuestas diferenciales.",
    "OCP": "Potencial en circuito abierto a lo largo del tiempo y su derivada.",
    "CV": "Voltametria ciclica y faltan indicadores que quizas se puedan extraer.",
    "EIS": "Espectroscopía de impedancia electroquímica: Nyquist, Bode y composicion visual entre mediciones, valores de corriente, potencial y temperatura para cada punto del espectro.",
    "PC": "Curvas de polarizacion y estabilidad temporal.",
    "Deg": "Series de degradacion galvanostatica y OCP: evolucion temporal de potencial, temperatura y derivadas.",
    "Analisis multiple": "Todas las anteriores como en FQ.",
}

RUNNERS = {
    "EIS": run_eis_pipeline,
    "PC": run_pc_pipeline,
    "OCP": run_ocp_pipeline,
    "CV": run_cv_pipeline,
    "Deg": run_deg_pipeline,
    "Activacion": run_activ_pipeline,
}

NOT_FOUND_MSG = {
    "EIS": "No se encontraron archivos .DTA con 'EISPOT' o 'Est_EIS' en la carpeta de entrada.",
    "PC": "No se encontraron archivos .DTA con 'Curva_Polarizacion_' en la carpeta de entrada.",
    "OCP": "No se encontraron archivos .DTA con 'OCP_' en la carpeta de entrada.",
    "CV": "No se encontraron archivos .DTA con 'Voltametria_ciclica' en la carpeta de entrada.",
    "Deg": "No se encontraron archivos .DTA con 'Degradacion_galvanostatica_19A_60C_' en la carpeta de entrada.",
    "Activacion": "No se encontraron archivos .DTA con 'Activacion_' en la carpeta de entrada.",
}

THEME_LABEL_LIST = [THEME_LABELS[name] for name in ("light", "dark")]
LANGUAGE_LABEL_LIST = [LANGUAGE_LABELS[name] for name in ("es", "en")]


class GamryProtocolApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Gamry Figure Studio")
        self.root.geometry("1024x640")
        self.root.minsize(940, 580)

        self.repo_dir = pathlib.Path(__file__).resolve().parent
        self.default_input_dir = self.repo_dir / "data2/test"
        self.default_output_dir = self.repo_dir / "outputs"
        self.default_output_dir.mkdir(parents=True, exist_ok=True)

        self.style = ttk.Style(self.root)
        self.current_theme = apply_theme(self.root, self.style, "light")

        self.options_window: tk.Toplevel | None = None
        self.selected_options: dict[str, list[str]] = {}

        self.theme_label_var = tk.StringVar(value=THEME_LABELS["light"])
        self.language_label_var = tk.StringVar(value=LANGUAGE_LABELS["es"])
        self.input_dir_var = tk.StringVar(value=str(self.default_input_dir))
        self.output_dir_var = tk.StringVar(value=str(self.default_output_dir))
        self.pipeline_var = tk.StringVar(value=next(iter(PIPELINE_OPTIONS)))
        font_defaults = DEFAULT_PLOT_FONT_DEFAULTS.as_strings()
        self.plot_title_fontsize_var = tk.StringVar(value=font_defaults["title"])
        self.plot_tick_fontsize_var = tk.StringVar(value=font_defaults["tick"])
        self.plot_label_fontsize_var = tk.StringVar(value=font_defaults["label"])
        self.plot_legend_fontsize_var = tk.StringVar(value=font_defaults["legend"])
        self.status_var = tk.StringVar(value="Listo. Seleccione carpetas y un pipeline.")
        self.preview_pipeline_var = tk.StringVar()
        self.preview_description_var = tk.StringVar()
        self.preview_outputs_var = tk.StringVar()
        self.preview_count_var = tk.StringVar()
        self.action_text_var = tk.StringVar()
        self.full_report_text_var = tk.StringVar(value=translate("full_report_button", "es"))

        self.build_main_window()
        self._update_pipeline_preview()
        self.root.after_idle(self._maximize_root_window)

    def _maximize_root_window(self) -> None:
        try:
            self.root.state("zoomed")
        except tk.TclError:
            try:
                self.root.attributes("-zoomed", True)
            except tk.TclError:
                pass

    def build_main_window(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main = ttk.Frame(self.root, style="App.TFrame", padding=(28, 24, 28, 20))
        main.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)

        header = ttk.Frame(main, style="App.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 22))
        header.columnconfigure(0, weight=1)

        title_block = ttk.Frame(header, style="App.TFrame")
        title_block.grid(row=0, column=0, sticky="w")

        ttk.Label(title_block, text="Gamry Figure Studio", style="HeroTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            title_block,
            text="Prepare resultados, ajuste figuras y exporte vistas listas para articulos y presentaciones.",
            style="HeroSubtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))

        theme_panel = ttk.Frame(header, style="Panel.TFrame", padding=(14, 12))
        theme_panel.grid(row=0, column=1, sticky="e")

        ttk.Label(theme_panel, text="App theme", style="PanelLabel.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 10)
        )
        self.theme_combo = ttk.Combobox(
            theme_panel,
            textvariable=self.theme_label_var,
            values=THEME_LABEL_LIST,
            state="readonly",
            width=12,
            style="App.TCombobox",
        )
        self.theme_combo.grid(row=0, column=1, sticky="e")
        self.theme_combo.bind("<<ComboboxSelected>>", self._on_theme_changed)

        ttk.Label(theme_panel, text=translate("language", "es"), style="PanelLabel.TLabel").grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=(8, 0)
        )
        self.language_combo = ttk.Combobox(
            theme_panel,
            textvariable=self.language_label_var,
            values=LANGUAGE_LABEL_LIST,
            state="readonly",
            width=12,
            style="App.TCombobox",
        )
        self.language_combo.grid(row=1, column=1, sticky="e", pady=(8, 0))
        self.language_combo.bind("<<ComboboxSelected>>", self._on_language_changed)

        content = ttk.Frame(main, style="App.TFrame")
        content.grid(row=1, column=0, sticky="nsew")
        content.columnconfigure(0, weight=3)
        content.columnconfigure(1, weight=2)
        content.rowconfigure(0, weight=1)

        workspace_card = ttk.Frame(content, style="Card.TFrame", padding=24)
        workspace_card.grid(row=0, column=0, sticky="nsew", padx=(0, 18))
        workspace_card.columnconfigure(0, weight=1)
        workspace_card.columnconfigure(1, weight=1)

        ttk.Label(workspace_card, text="Workspace", style="CardTitle.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w"
        )
        ttk.Label(
            workspace_card,
            text="Seleccione las carpetas de trabajo y el pipeline que quiere abrir para personalizar las figuras.",
            style="CardMuted.TLabel",
            wraplength=520,
            justify="left",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(8, 18))

        self._build_path_row(
            parent=workspace_card,
            row=2,
            label_text="Carpeta de entrada",
            variable=self.input_dir_var,
            browse_target="input",
        )
        self._build_path_row(
            parent=workspace_card,
            row=4,
            label_text="Carpeta de salida",
            variable=self.output_dir_var,
            browse_target="output",
        )

        ttk.Label(workspace_card, text="Pipeline", style="FieldLabel.TLabel").grid(
            row=6, column=0, sticky="w", pady=(4, 8)
        )
        self.pipeline_combo = ttk.Combobox(
            workspace_card,
            textvariable=self.pipeline_var,
            values=list(PIPELINE_OPTIONS.keys()),
            state="readonly",
            style="App.TCombobox",
        )
        self.pipeline_combo.grid(row=7, column=0, columnspan=3, sticky="ew")
        self.pipeline_combo.bind("<<ComboboxSelected>>", self._update_pipeline_preview)

        ttk.Label(
            workspace_card,
            text="La siguiente ventana permite elegir que figuras abrir y exportar para el pipeline activo.",
            style="CardMuted.TLabel",
            wraplength=520,
            justify="left",
        ).grid(row=8, column=0, columnspan=3, sticky="w", pady=(10, 14))

        plot_defaults_box = ttk.LabelFrame(workspace_card, text="Plot defaults")
        plot_defaults_box.grid(row=9, column=0, columnspan=3, sticky="ew", pady=(0, 20))
        plot_defaults_box.columnconfigure(1, weight=1)
        plot_defaults_box.columnconfigure(3, weight=1)

        ttk.Label(plot_defaults_box, text="Title size", style="FieldLabel.TLabel").grid(
            row=0, column=0, sticky="w", padx=12, pady=(12, 6)
        )
        ttk.Spinbox(
            plot_defaults_box,
            from_=6.0,
            to=50.0,
            increment=0.5,
            textvariable=self.plot_title_fontsize_var,
            width=10,
            style="App.TSpinbox",
        ).grid(row=0, column=1, sticky="w", padx=(0, 12), pady=(12, 6))

        ttk.Label(plot_defaults_box, text="Tick size", style="FieldLabel.TLabel").grid(
            row=0, column=2, sticky="w", padx=(6, 12), pady=(12, 6)
        )
        ttk.Spinbox(
            plot_defaults_box,
            from_=6.0,
            to=40.0,
            increment=0.5,
            textvariable=self.plot_tick_fontsize_var,
            width=10,
            style="App.TSpinbox",
        ).grid(row=0, column=3, sticky="w", padx=(0, 12), pady=(12, 6))

        ttk.Label(plot_defaults_box, text="Label size", style="FieldLabel.TLabel").grid(
            row=1, column=0, sticky="w", padx=12, pady=(6, 12)
        )
        ttk.Spinbox(
            plot_defaults_box,
            from_=6.0,
            to=40.0,
            increment=0.5,
            textvariable=self.plot_label_fontsize_var,
            width=10,
            style="App.TSpinbox",
        ).grid(row=1, column=1, sticky="w", padx=(0, 12), pady=(6, 12))

        ttk.Label(plot_defaults_box, text="Legend size", style="FieldLabel.TLabel").grid(
            row=1, column=2, sticky="w", padx=(6, 12), pady=(6, 12)
        )
        ttk.Spinbox(
            plot_defaults_box,
            from_=6.0,
            to=40.0,
            increment=0.5,
            textvariable=self.plot_legend_fontsize_var,
            width=10,
            style="App.TSpinbox",
        ).grid(row=1, column=3, sticky="w", padx=(0, 12), pady=(6, 12))

        action_row = ttk.Frame(workspace_card, style="Card.TFrame")
        action_row.grid(row=10, column=0, columnspan=3, sticky="ew")
        action_row.columnconfigure(0, weight=1)

        ttk.Button(
            action_row,
            textvariable=self.action_text_var,
            style="Accent.TButton",
            command=self.pipeline_selected,
        ).grid(row=0, column=0, sticky="w")
        ttk.Button(
            action_row,
            textvariable=self.full_report_text_var,
            style="Accent.TButton",
            command=self.full_report_selected,
        ).grid(row=0, column=1, sticky="w", padx=(12, 0))
        ttk.Button(
            action_row,
            text="Restaurar rutas",
            style="Subtle.TButton",
            command=self.reset_defaults,
        ).grid(row=0, column=2, sticky="w", padx=(12, 0))

        preview_card = ttk.Frame(content, style="Card.TFrame", padding=24)
        preview_card.grid(row=0, column=1, sticky="nsew")
        preview_card.columnconfigure(0, weight=1)

        ttk.Label(preview_card, text="Pipeline preview", style="CardTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(preview_card, textvariable=self.preview_pipeline_var, style="CardValue.TLabel").grid(
            row=1, column=0, sticky="w", pady=(10, 4)
        )
        ttk.Label(preview_card, textvariable=self.preview_count_var, style="FieldLabel.TLabel").grid(
            row=2, column=0, sticky="w", pady=(0, 10)
        )
        ttk.Label(
            preview_card,
            textvariable=self.preview_description_var,
            style="CardBody.TLabel",
            wraplength=300,
            justify="left",
        ).grid(row=3, column=0, sticky="w")

        ttk.Separator(preview_card).grid(row=4, column=0, sticky="ew", pady=18)

        ttk.Label(preview_card, text="Salidas disponibles", style="FieldLabel.TLabel").grid(
            row=5, column=0, sticky="w"
        )
        ttk.Label(
            preview_card,
            textvariable=self.preview_outputs_var,
            style="CardBody.TLabel",
            justify="left",
            wraplength=300,
        ).grid(row=6, column=0, sticky="w", pady=(8, 18))

        viewport = ttk.Frame(preview_card, style="PlotViewport.TFrame", padding=18)
        viewport.grid(row=7, column=0, sticky="ew")
        viewport.columnconfigure(0, weight=1)

        ttk.Label(viewport, text="White figure surface", style="PlotViewportTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            viewport,
            text=(
                "The plot viewport should stay white in every mode so exported figures keep the same visual ground "
                "you expect in papers and slides."
            ),
            style="PlotViewportBody.TLabel",
            wraplength=260,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))

        status_frame = ttk.Frame(main, style="Status.TFrame", padding=(16, 12))
        status_frame.grid(row=2, column=0, sticky="ew", pady=(18, 0))
        status_frame.columnconfigure(0, weight=1)

        ttk.Label(status_frame, textvariable=self.status_var, style="Status.TLabel").grid(
            row=0, column=0, sticky="w"
        )

    def _build_path_row(
        self,
        parent: ttk.Frame,
        row: int,
        label_text: str,
        variable: tk.StringVar,
        browse_target: str,
    ) -> None:
        ttk.Label(parent, text=label_text, style="FieldLabel.TLabel").grid(
            row=row, column=0, sticky="w", pady=(0, 8)
        )
        entry = ttk.Entry(parent, textvariable=variable, style="App.TEntry")
        entry.grid(row=row + 1, column=0, columnspan=2, sticky="ew", padx=(0, 10))
        ttk.Button(
            parent,
            text="Buscar...",
            style="Subtle.TButton",
            command=lambda target=browse_target: self.browse_button(target),
        ).grid(row=row + 1, column=2, sticky="ew")

    def _update_pipeline_preview(self, _event=None) -> None:
        selected_pipeline = self.pipeline_var.get().strip()
        options = PIPELINE_OPTIONS.get(selected_pipeline, [])
        description = PIPELINE_DESCRIPTIONS.get(
            selected_pipeline,
            "Seleccione un pipeline para ver sus vistas disponibles.",
        )

        self.preview_pipeline_var.set(selected_pipeline or "Pipeline")
        self.preview_description_var.set(description)
        self.preview_count_var.set(f"{len(options)} salidas graficas disponibles")
        self.preview_outputs_var.set(
            "\n".join(f"- {option}" for option in options) if options else "Sin vistas configuradas."
        )
        self.action_text_var.set(f"Configurar {selected_pipeline}" if selected_pipeline else "Configurar pipeline")
        self.full_report_text_var.set(translate("full_report_button", self._current_language()))

    def _apply_theme(self, theme_name: str) -> None:
        self.current_theme = apply_theme(self.root, self.style, theme_name)
        self.root.configure(bg=self.current_theme.window_bg)
        if self.options_window is not None and self.options_window.winfo_exists():
            self.options_window.configure(bg=self.current_theme.window_bg)

    def _on_theme_changed(self, _event=None) -> None:
        selected_label = self.theme_label_var.get().strip()
        theme_name = THEME_NAME_BY_LABEL.get(selected_label, "light")
        self._apply_theme(theme_name)
        self.set_status(f"Theme activo: {selected_label}.")

    def _current_language(self) -> str:
        selected_label = self.language_label_var.get().strip()
        return LANGUAGE_CODE_BY_LABEL.get(selected_label, "es")

    def _on_language_changed(self, _event=None) -> None:
        language = self._current_language()
        self.full_report_text_var.set(translate("full_report_button", language))
        self.set_status(
            translate(
                "language_status",
                language,
                language=self.language_label_var.get().strip(),
            )
        )

    def set_status(self, message: str) -> None:
        self.status_var.set(message)

    def reset_defaults(self) -> None:
        self.input_dir_var.set(str(self.default_input_dir))
        self.output_dir_var.set(str(self.default_output_dir))
        self.pipeline_var.set(next(iter(PIPELINE_OPTIONS)))
        self.language_label_var.set(LANGUAGE_LABELS["es"])
        font_defaults = DEFAULT_PLOT_FONT_DEFAULTS.as_strings()
        self.plot_title_fontsize_var.set(font_defaults["title"])
        self.plot_tick_fontsize_var.set(font_defaults["tick"])
        self.plot_label_fontsize_var.set(font_defaults["label"])
        self.plot_legend_fontsize_var.set(font_defaults["legend"])
        self._update_pipeline_preview()
        self.set_status("Rutas y estilos de plot restaurados.")

    def browse_button(self, button_type: str) -> None:
        if button_type == "input":
            title = "Select input folder"
            start_dir = self.input_dir_var.get().strip() or str(self.default_input_dir)
        else:
            title = "Select output folder"
            start_dir = self.output_dir_var.get().strip() or str(self.default_output_dir)

        folder = filedialog.askdirectory(title=title, initialdir=start_dir)
        if not folder:
            self.set_status("Seleccion cancelada.")
            return

        if button_type == "input":
            self.input_dir_var.set(folder)
            self.set_status("Carpeta de entrada seleccionada.")
        else:
            self.output_dir_var.set(folder)
            self.set_status("Carpeta de salida seleccionada.")

    def full_report_selected(self) -> None:
        input_dir_text = self.input_dir_var.get().strip()
        output_dir_text = self.output_dir_var.get().strip()
        language = self._current_language()

        if not input_dir_text or not output_dir_text:
            self.set_status("Complete las carpetas de entrada y salida antes de continuar.")
            return

        input_dir = pathlib.Path(input_dir_text)
        output_dir = pathlib.Path(output_dir_text)

        if not input_dir.exists() or not input_dir.is_dir():
            self.set_status("La carpeta de entrada no existe o no es valida.")
            return

        try:
            plot_font_defaults = self._current_plot_font_defaults()
        except ValueError as exc:
            self.set_status(f"Error en estilos de plot: {exc}")
            return

        output_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir_var.set(str(output_dir))

        progress_window = tk.Toplevel(self.root)
        progress_window.title(translate("full_report", language))
        progress_window.geometry("440x150")
        progress_window.resizable(False, False)
        progress_window.transient(self.root)
        progress_window.configure(bg=self.current_theme.window_bg)
        progress_window.protocol("WM_DELETE_WINDOW", lambda: None)

        panel = ttk.Frame(progress_window, style="Card.TFrame", padding=18)
        panel.pack(fill="both", expand=True, padx=12, pady=12)
        message_var = tk.StringVar(value=translate("full_report_generating", language))
        ttk.Label(panel, textvariable=message_var, style="CardBody.TLabel", wraplength=380).pack(
            anchor="w",
            fill="x",
            pady=(0, 12),
        )
        progress_bar = ttk.Progressbar(panel, mode="indeterminate", length=380)
        progress_bar.pack(fill="x")
        progress_bar.start(12)

        self.set_status(translate("full_report_generating", language))

        def report_progress(message: str, _done: int | None = None, _total: int | None = None) -> None:
            self.root.after(0, lambda msg=message: message_var.set(msg))

        def finish_success(path: pathlib.Path) -> None:
            if progress_window.winfo_exists():
                progress_bar.stop()
                progress_window.destroy()
            status = f"{translate('full_report_completed', language)}: {path}"
            self.set_status(status)
            messagebox.showinfo(translate("full_report", language), status)

        def finish_error(error_message: str) -> None:
            if progress_window.winfo_exists():
                progress_bar.stop()
                progress_window.destroy()
            self.set_status(f"{translate('full_report_failed', language)}: {error_message}")
            messagebox.showerror(translate("full_report", language), error_message)

        def worker() -> None:
            try:
                exported_path = generate_full_report(
                    input_dir,
                    output_dir,
                    language=language,
                    font_defaults=plot_font_defaults,
                    progress_callback=report_progress,
                )
            except Exception as exc:
                import traceback

                print(traceback.format_exc())
                self.root.after(0, lambda msg=f"{type(exc).__name__}: {exc}": finish_error(msg))
                return

            self.root.after(0, lambda path=exported_path: finish_success(path))

        threading.Thread(target=worker, daemon=True).start()

    def pipeline_selected(self) -> None:
        input_dir_text = self.input_dir_var.get().strip()
        output_dir_text = self.output_dir_var.get().strip()
        selected_pipeline = self.pipeline_var.get().strip()

        if not input_dir_text or not output_dir_text or not selected_pipeline:
            self.set_status("Complete todos los campos antes de continuar.")
            return

        input_dir = pathlib.Path(input_dir_text)
        output_dir = pathlib.Path(output_dir_text)

        if not input_dir.exists() or not input_dir.is_dir():
            self.set_status("La carpeta de entrada no existe o no es valida.")
            return

        output_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir_var.set(str(output_dir))
        self.set_status(f"{selected_pipeline}: configure las vistas que desea abrir.")
        self.open_pipeline_window(selected_pipeline)

    def _close_options_window(self) -> None:
        if self.options_window is None:
            return
        if self.options_window.winfo_exists():
            self.options_window.destroy()
        self.options_window = None

    def open_pipeline_window(self, selected_pipeline: str) -> None:
        options = PIPELINE_OPTIONS.get(selected_pipeline, [])

        if self.options_window is not None and self.options_window.winfo_exists():
            self.options_window.destroy()

        self.options_window = tk.Toplevel(self.root)
        self.options_window.title(f"{selected_pipeline} outputs")
        self.options_window.geometry("460x440")
        self.options_window.minsize(420, 360)
        self.options_window.transient(self.root)
        self.options_window.grab_set()
        self.options_window.configure(bg=self.current_theme.window_bg)
        self.options_window.protocol("WM_DELETE_WINDOW", self._close_options_window)

        outer = ttk.Frame(self.options_window, style="App.TFrame", padding=(22, 22, 22, 18))
        outer.pack(fill="both", expand=True)

        card = ttk.Frame(outer, style="Card.TFrame", padding=22)
        card.pack(fill="both", expand=True)
        card.columnconfigure(0, weight=1)

        ttk.Label(card, text=selected_pipeline, style="CardValue.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            card,
            text="Seleccione las vistas que quiere generar y abrir para este pipeline.",
            style="CardMuted.TLabel",
            wraplength=360,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(8, 16))

        selected_count_var = tk.StringVar(value="0 vistas seleccionadas")
        previous_selection = set(self.selected_options.get(selected_pipeline, []))
        option_vars: list[tuple[str, tk.BooleanVar]] = []

        options_frame = ttk.Frame(card, style="Card.TFrame")
        options_frame.grid(row=2, column=0, sticky="nsew")
        options_frame.columnconfigure(0, weight=1)
        card.rowconfigure(2, weight=1)

        if options:
            for row_index, option in enumerate(options):
                var = tk.BooleanVar(value=option in previous_selection)
                chk = ttk.Checkbutton(
                    options_frame,
                    text=option,
                    variable=var,
                    style="Card.TCheckbutton",
                    command=lambda: self._update_selected_count(selected_count_var, option_vars),
                )
                chk.grid(row=row_index, column=0, sticky="w", pady=(0, 8))
                option_vars.append((option, var))
        else:
            ttk.Label(
                options_frame,
                text="Este pipeline todavia no tiene vistas configuradas en el lanzador.",
                style="CardMuted.TLabel",
                wraplength=360,
                justify="left",
            ).grid(row=0, column=0, sticky="w")

        self._update_selected_count(selected_count_var, option_vars)

        ttk.Label(card, textvariable=selected_count_var, style="FieldLabel.TLabel").grid(
            row=3, column=0, sticky="w", pady=(12, 0)
        )

        if selected_pipeline not in RUNNERS:
            ttk.Label(
                card,
                text="Este pipeline aun no tiene un runner asignado desde el lanzador principal.",
                style="CardMuted.TLabel",
                wraplength=360,
                justify="left",
            ).grid(row=4, column=0, sticky="w", pady=(8, 0))

        buttons = ttk.Frame(card, style="Card.TFrame")
        buttons.grid(row=5, column=0, sticky="ew", pady=(20, 0))
        buttons.columnconfigure(0, weight=1)

        ttk.Button(
            buttons,
            text="Cancelar",
            style="Subtle.TButton",
            command=self._close_options_window,
        ).grid(row=0, column=0, sticky="w")

        def confirm() -> None:
            try:
                plot_font_defaults = self._current_plot_font_defaults()
            except ValueError as exc:
                self.set_status(f"Error en estilos de plot: {exc}")
                return

            selected = [name for name, var in option_vars if var.get()]
            self.selected_options[selected_pipeline] = selected

            if selected:
                self.set_status(f"{selected_pipeline}: {len(selected)} vista(s) seleccionada(s).")
            else:
                self.set_status(f"{selected_pipeline}: no se seleccionaron vistas.")

            runner = RUNNERS.get(selected_pipeline)
            if runner is None:
                self.set_status(f"{selected_pipeline}: pipeline no implementado aun.")
                self._close_options_window()
                return

            input_dir = pathlib.Path(self.input_dir_var.get().strip())
            output_dir = pathlib.Path(self.output_dir_var.get().strip())
            self._close_options_window()

            try:
                exported_files = runner(
                    input_dir,
                    output_dir,
                    selected,
                    font_defaults=plot_font_defaults,
                    language=self._current_language(),
                )
                if exported_files:
                    self.set_status(
                        f"{selected_pipeline} ejecutado: {len(exported_files)} archivo(s) .xlsx creado(s)."
                    )
                else:
                    self.set_status(NOT_FOUND_MSG.get(selected_pipeline, "No se encontraron archivos."))
            except Exception as exc:
                import traceback

                print(traceback.format_exc())
                self.set_status(f"Error en {selected_pipeline}: {type(exc).__name__}: {exc}")

        ttk.Button(
            buttons,
            text="Ejecutar pipeline",
            style="Accent.TButton",
            command=confirm,
        ).grid(row=0, column=1, sticky="e", padx=(12, 0))

    @staticmethod
    def _update_selected_count(
        target_var: tk.StringVar,
        option_vars: list[tuple[str, tk.BooleanVar]],
    ) -> None:
        selected_count = sum(1 for _name, var in option_vars if var.get())
        target_var.set(f"{selected_count} vista(s) seleccionada(s)")

    def _current_plot_font_defaults(self) -> PlotFontDefaults:
        return parse_plot_font_defaults(
            title=self.plot_title_fontsize_var.get().strip(),
            tick=self.plot_tick_fontsize_var.get().strip(),
            label=self.plot_label_fontsize_var.get().strip(),
            legend=self.plot_legend_fontsize_var.get().strip(),
        )

    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    app = GamryProtocolApp()
    app.run()
