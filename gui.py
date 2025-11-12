import math
import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

AUTO_PRIMARY_OPTION = "Automatico (più piccolo)"
ADDRESS_MODE_OPTIONS = {
    "Siatel (dettagliato)": "siatel",
    "Siatel compatto": "compatto",
}
COMPARE_MODE_OPTIONS = {
    "Siatel compatto": "compatto",
    "Siatel (dettagliato)": "dettagliato",
}
SPLIT_HEADER_OPTIONS = {
    "Intestazione in tutti i file": "repeat",
    "Intestazione solo nel primo file": "first-only",
    "Intestazione formattata (tutti i file)": "formatted",
    "Intestazione formattata (solo primo file)": "formatted-first-only",
    "Nessuna intestazione": "none",
}
SPLIT_TARGET_BYTES = 3 * 1024 * 1024

EXCEL_FILETYPES = [
    ("File Excel", "*.xlsx *.xlsm *.xls"),
    ("Tutti i file", "*.*"),
]
CSV_FILETYPES = [
    ("File CSV", "*.csv"),
    ("Tutti i file", "*.*"),
]
SPLIT_FILETYPES = [
    ("File Excel o CSV", "*.xlsx *.xlsm *.xls *.csv"),
    ("File CSV", "*.csv"),
    ("File Excel", "*.xlsx *.xlsm *.xls"),
    ("Tutti i file", "*.*"),
]

CONVERT_DIRECTION_OPTIONS = {
    "Excel -> CSV": ("csv", EXCEL_FILETYPES),
    "CSV -> Excel": ("excel", CSV_FILETYPES),
}

SCRIPTS = {
    "Unisci file": "unisci_file.py",
    "Dividi file": "spilit_file.py",
    "Merge file": "marge_file.py",
    "Dividi indirizzi": "dividi_indirizzi.py",
    "Converti file": "converti_file.py",
    "Confronta indirizzi": "confronta_indirizzi.py",
}


def ensure_directories_exist() -> None:
    """Ensure that input and output directories are available."""
    for directory in (INPUT_DIR, OUTPUT_DIR):
        os.makedirs(directory, exist_ok=True)


class ScriptRunnerGUI(tk.Tk):
    """GUI evoluta per configurare ed eseguire gli script Excel."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Excel Script Runner")
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass
        self.configure(bg="#f4f6fb")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        # Stato Unisci file
        self.unisci_files: list[str] = []
        self.unisci_output_var = tk.StringVar(value="unione.xlsx")

        # Stato Dividi file
        self.split_file_var = tk.StringVar(value="")
        self.chunk_size_var = tk.StringVar(value="100")
        self.split_header_mode_var = tk.StringVar(value=next(iter(SPLIT_HEADER_OPTIONS)))
        self.split_suggestion_var = tk.StringVar(value="")
        self._split_suggestion_file: str | None = None

        # Stato Merge file
        self.merge_files: list[str | None] = [None, None]
        self.merge_labels: list[ttk.Label] = []
        self.merge_primary_var = tk.StringVar(value=AUTO_PRIMARY_OPTION)
        self.merge_primary_options: dict[str, str] = {AUTO_PRIMARY_OPTION: AUTO_PRIMARY_OPTION}

        # Stato Confronta indirizzi
        self.compare_files: list[str | None] = [None, None]
        self.compare_labels: list[ttk.Label] = []
        default_compare_mode = next(iter(COMPARE_MODE_OPTIONS))
        self.compare_mode_var = tk.StringVar(value=default_compare_mode)
        self.compare_map_var = tk.StringVar(value="")
        self.compare_map_label: ttk.Label | None = None
        self.compare_run_button: ttk.Button | None = None

        # Stato Converti file
        self.convert_files: list[str] = []
        self.convert_direction_var = tk.StringVar(value=next(iter(CONVERT_DIRECTION_OPTIONS)))
        self.convert_csv_delimiter_var = tk.StringVar(value=",")
        self.convert_delimiter_entry = None
        self.convert_listbox = None
        self.convert_run_button = None

        # Stato Dividi indirizzi
        default_address_mode = next(iter(ADDRESS_MODE_OPTIONS))
        self.address_mode_var = tk.StringVar(value=default_address_mode)
        self.address_file_var = tk.StringVar(value="")

        self._build_widgets()
        self._center_window()

    # ------------------------------------------------------------------ UI build
    def _build_widgets(self) -> None:
        header_style = "Header.TLabel"
        self.style.configure(header_style, font=("Segoe UI", 16, "bold"), background="#f4f6fb")

        main_frame = ttk.Frame(self, padding=15)
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)

        header_label = ttk.Label(main_frame, text="Excel Script Runner", style=header_style)
        header_label.grid(row=0, column=0, sticky="w")

        top_buttons = ttk.Frame(main_frame)
        top_buttons.grid(row=1, column=0, sticky="w", pady=(8, 12))

        ttk.Button(
            top_buttons,
            text="Apri cartella input",
            command=lambda: self._open_directory(INPUT_DIR),
        ).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(
            top_buttons,
            text="Apri cartella output",
            command=lambda: self._open_directory(OUTPUT_DIR),
        ).grid(row=0, column=1)

        notebook = ttk.Notebook(main_frame)
        notebook.grid(row=2, column=0, sticky="nsew")

        self._build_unisci_tab(notebook)
        self._build_split_tab(notebook)
        self._build_merge_tab(notebook)
        self._build_compare_tab(notebook)
        self._build_convert_tab(notebook)
        self._build_address_tab(notebook)

        log_frame = ttk.LabelFrame(main_frame, text="Log esecuzione")
        log_frame.grid(row=3, column=0, sticky="nsew", pady=(12, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, height=12, wrap="word", state=tk.DISABLED)
        self.log_text.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def _build_unisci_tab(self, notebook: ttk.Notebook) -> None:
        frame = ttk.Frame(notebook, padding=12)
        frame.columnconfigure(0, weight=1)
        notebook.add(frame, text="Unisci file")

        ttk.Label(
            frame,
            text="Seleziona uno o più file Excel da concatenare nell'ordine indicato.",
        ).grid(row=0, column=0, sticky="w")

        list_frame = ttk.Frame(frame)
        list_frame.grid(row=1, column=0, sticky="nsew", pady=8)
        list_frame.columnconfigure(0, weight=1)

        self.unisci_listbox = tk.Listbox(list_frame, height=8, activestyle="dotbox")
        self.unisci_listbox.grid(row=0, column=0, sticky="nsew")

        list_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.unisci_listbox.yview)
        list_scroll.grid(row=0, column=1, sticky="ns")
        self.unisci_listbox.configure(yscrollcommand=list_scroll.set)

        buttons_frame = ttk.Frame(frame)
        buttons_frame.grid(row=2, column=0, sticky="w")

        ttk.Button(
            buttons_frame,
            text="Aggiungi file…",
            command=self._add_unisci_files,
        ).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(
            buttons_frame,
            text="Rimuovi selezionato",
            command=self._remove_unisci_selected,
        ).grid(row=0, column=1, padx=(0, 6))
        ttk.Button(
            buttons_frame,
            text="Svuota elenco",
            command=self._clear_unisci_files,
        ).grid(row=0, column=2)

        output_frame = ttk.Frame(frame)
        output_frame.grid(row=3, column=0, sticky="w", pady=(10, 0))

        ttk.Label(output_frame, text="Nome file di output:").grid(row=0, column=0, padx=(0, 6))
        ttk.Entry(output_frame, width=30, textvariable=self.unisci_output_var).grid(
            row=0, column=1, sticky="w"
        )

        self.unisci_run_button = ttk.Button(
            frame, text="Esegui unione", command=self._run_unisci_script
        )
        self.unisci_run_button.grid(row=4, column=0, pady=(12, 0), sticky="e")

    def _build_split_tab(self, notebook: ttk.Notebook) -> None:
        frame = ttk.Frame(notebook, padding=12)
        frame.columnconfigure(0, weight=1)
        notebook.add(frame, text="Dividi file")

        ttk.Label(
            frame,
            text="Scegli il file da suddividere e imposta il numero di righe per blocco.",
        ).grid(row=0, column=0, sticky="w")

        select_frame = ttk.Frame(frame)
        select_frame.grid(row=1, column=0, sticky="w", pady=8)

        ttk.Label(select_frame, text="File selezionato:").grid(row=0, column=0, padx=(0, 6))
        self.split_file_label = ttk.Label(select_frame, text="(nessun file)")
        self.split_file_label.grid(row=0, column=1, sticky="w")

        ttk.Button(
            frame,
            text="Scegli file…",
            command=self._choose_split_file,
        ).grid(row=2, column=0, sticky="w")

        chunk_frame = ttk.Frame(frame)
        chunk_frame.grid(row=3, column=0, sticky="w", pady=(10, 0))
        ttk.Label(chunk_frame, text="Righe per file:").grid(row=0, column=0, padx=(0, 6))
        self.chunk_size_spinbox = ttk.Spinbox(
            chunk_frame,
            from_=10,
            to=100000,
            increment=10,
            textvariable=self.chunk_size_var,
            width=10,
        )
        self.chunk_size_spinbox.grid(row=0, column=1, sticky="w")
        ttk.Label(
            chunk_frame,
            textvariable=self.split_suggestion_var,
            foreground="#555555",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        header_frame = ttk.Frame(frame)
        header_frame.grid(row=4, column=0, sticky="w", pady=(10, 0))
        ttk.Label(header_frame, text="Gestione intestazione:").grid(row=0, column=0, padx=(0, 6))
        ttk.Combobox(
            header_frame,
            state="readonly",
            width=30,
            textvariable=self.split_header_mode_var,
            values=list(SPLIT_HEADER_OPTIONS.keys()),
        ).grid(row=0, column=1, sticky="w")

        self.split_run_button = ttk.Button(
            frame, text="Esegui divisione", command=self._run_split_script
        )
        self.split_run_button.grid(row=5, column=0, pady=(12, 0), sticky="e")

    def _build_merge_tab(self, notebook: ttk.Notebook) -> None:
        frame = ttk.Frame(notebook, padding=12)
        frame.columnconfigure(0, weight=1)
        notebook.add(frame, text="Merge file")

        ttk.Label(
            frame,
            text="Seleziona i due file da confrontare sulla colonna 'match'.",
        ).grid(row=0, column=0, sticky="w")

        for idx in range(2):
            row = idx + 1
            label_text = "File A:" if idx == 0 else "File B:"
            file_row = ttk.Frame(frame)
            file_row.grid(row=row, column=0, sticky="w", pady=6)
            ttk.Label(file_row, text=label_text).grid(row=0, column=0, padx=(0, 6))

            label = ttk.Label(file_row, text="(nessun file)")
            label.grid(row=0, column=1, sticky="w")
            self.merge_labels.append(label)

            ttk.Button(
                file_row,
                text="Scegli…",
                command=lambda index=idx: self._choose_merge_file(index),
            ).grid(row=0, column=2, padx=(8, 0))

        primary_frame = ttk.Frame(frame)
        primary_frame.grid(row=3, column=0, sticky="w", pady=(10, 0))
        ttk.Label(primary_frame, text="Lato completo:").grid(row=0, column=0, padx=(0, 6))
        self.merge_primary_combobox = ttk.Combobox(
            primary_frame,
            state="readonly",
            width=30,
            textvariable=self.merge_primary_var,
            values=[AUTO_PRIMARY_OPTION],
        )
        self.merge_primary_combobox.grid(row=0, column=1, sticky="w")

        self.merge_run_button = ttk.Button(
            frame, text="Esegui merge", command=self._run_merge_script
        )
        self.merge_run_button.grid(row=4, column=0, pady=(12, 0), sticky="e")

    def _build_compare_tab(self, notebook: ttk.Notebook) -> None:
        frame = ttk.Frame(notebook, padding=12)
        frame.columnconfigure(0, weight=1)
        notebook.add(frame, text="Confronta indirizzi")

        ttk.Label(
            frame,
            text="Confronta gli indirizzi dei due file sulla colonna 'match'.",
        ).grid(row=0, column=0, sticky="w")

        for idx in range(2):
            row = idx + 1
            file_frame = ttk.Frame(frame)
            file_frame.grid(row=row, column=0, sticky="w", pady=6)
            ttk.Label(file_frame, text=f"File {'A' if idx == 0 else 'B'}:").grid(
                row=0, column=0, padx=(0, 6)
            )
            label = ttk.Label(file_frame, text="(nessun file)")
            label.grid(row=0, column=1, sticky="w")
            self.compare_labels.append(label)
            ttk.Button(
                file_frame,
                text="Scegli…",
                command=lambda index=idx: self._choose_compare_file(index),
            ).grid(row=0, column=2, padx=(8, 0))

        mode_frame = ttk.Frame(frame)
        mode_frame.grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Label(mode_frame, text="Modalità di indirizzo:").grid(row=0, column=0, padx=(0, 6))
        ttk.Combobox(
            mode_frame,
            state="readonly",
            width=28,
            textvariable=self.compare_mode_var,
            values=list(COMPARE_MODE_OPTIONS.keys()),
        ).grid(row=0, column=1, sticky="w")

        comune_frame = ttk.LabelFrame(frame, text="Mappa equivalenze comuni (opzionale)")
        comune_frame.grid(row=4, column=0, sticky="we", pady=(10, 0))
        comune_frame.columnconfigure(1, weight=1)
        ttk.Label(comune_frame, text="File selezionato:").grid(row=0, column=0, padx=(0, 6))
        self.compare_map_label = ttk.Label(comune_frame, text="(nessun file)")
        self.compare_map_label.grid(row=0, column=1, sticky="w")
        buttons = ttk.Frame(comune_frame)
        buttons.grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Button(
            buttons,
            text="Scegli CSV…",
            command=self._choose_compare_map_file,
        ).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(
            buttons,
            text="Rimuovi",
            command=self._clear_compare_map_file,
        ).grid(row=0, column=1)
        ttk.Label(
            comune_frame,
            text="Se non selezioni nulla viene usato automaticamente 'comuni_equivalenze.csv'.",
            foreground="#555555",
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))

        self.compare_run_button = ttk.Button(
            frame, text="Esegui confronto", command=self._run_compare_script
        )
        self.compare_run_button.grid(row=5, column=0, pady=(12, 0), sticky="e")

    def _build_convert_tab(self, notebook: ttk.Notebook) -> None:
        frame = ttk.Frame(notebook, padding=12)
        frame.columnconfigure(0, weight=1)
        notebook.add(frame, text="Converti file")

        ttk.Label(
            frame,
            text="Seleziona i file da convertire e il formato di destinazione.",
        ).grid(row=0, column=0, sticky="w")

        direction_frame = ttk.Frame(frame)
        direction_frame.grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Label(direction_frame, text="Tipo conversione:").grid(row=0, column=0, padx=(0, 6))
        direction_combobox = ttk.Combobox(
            direction_frame,
            state="readonly",
            width=25,
            textvariable=self.convert_direction_var,
            values=list(CONVERT_DIRECTION_OPTIONS.keys()),
        )
        direction_combobox.grid(row=0, column=1, sticky="w")
        direction_combobox.bind("<<ComboboxSelected>>", self._on_convert_direction_changed)

        list_frame = ttk.Frame(frame)
        list_frame.grid(row=2, column=0, sticky="nsew", pady=8)
        list_frame.columnconfigure(0, weight=1)

        self.convert_listbox = tk.Listbox(list_frame, height=8, activestyle="dotbox")
        self.convert_listbox.grid(row=0, column=0, sticky="nsew")

        list_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.convert_listbox.yview)
        list_scroll.grid(row=0, column=1, sticky="ns")
        self.convert_listbox.configure(yscrollcommand=list_scroll.set)

        buttons_frame = ttk.Frame(frame)
        buttons_frame.grid(row=3, column=0, sticky="w")

        ttk.Button(
            buttons_frame,
            text="Aggiungi file…",
            command=self._add_convert_files,
        ).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(
            buttons_frame,
            text="Rimuovi selezionato",
            command=self._remove_convert_selected,
        ).grid(row=0, column=1, padx=(0, 6))
        ttk.Button(
            buttons_frame,
            text="Svuota elenco",
            command=self._clear_convert_files,
        ).grid(row=0, column=2)

        delimiter_frame = ttk.Frame(frame)
        delimiter_frame.grid(row=4, column=0, sticky="w", pady=(10, 0))
        ttk.Label(delimiter_frame, text="Delimitatore CSV:").grid(row=0, column=0, padx=(0, 6))
        self.convert_delimiter_entry = ttk.Entry(
            delimiter_frame,
            width=5,
            textvariable=self.convert_csv_delimiter_var,
        )
        self.convert_delimiter_entry.grid(row=0, column=1, sticky="w")

        self.convert_run_button = ttk.Button(
            frame, text="Esegui conversione", command=self._run_convert_script
        )
        self.convert_run_button.grid(row=5, column=0, pady=(12, 0), sticky="e")

        self._on_convert_direction_changed()

    def _build_address_tab(self, notebook: ttk.Notebook) -> None:
        frame = ttk.Frame(notebook, padding=12)
        frame.columnconfigure(0, weight=1)
        notebook.add(frame, text="Dividi indirizzi")

        ttk.Label(
            frame,
            text="Seleziona il file da elaborare e scegli la modalità di divisione.",
        ).grid(row=0, column=0, sticky="w")

        file_frame = ttk.Frame(frame)
        file_frame.grid(row=1, column=0, sticky="w", pady=8)

        ttk.Label(file_frame, text="File selezionato:").grid(row=0, column=0, padx=(0, 6))
        self.address_file_label = ttk.Label(file_frame, text="(nessun file)")
        self.address_file_label.grid(row=0, column=1, sticky="w")

        ttk.Button(
            frame,
            text="Scegli file…",
            command=self._choose_address_file,
        ).grid(row=2, column=0, sticky="w")

        mode_frame = ttk.Frame(frame)
        mode_frame.grid(row=3, column=0, sticky="w", pady=(10, 0))
        ttk.Label(mode_frame, text="Modalità:").grid(row=0, column=0, padx=(0, 6))
        self.address_mode_combobox = ttk.Combobox(
            mode_frame,
            state="readonly",
            textvariable=self.address_mode_var,
            values=list(ADDRESS_MODE_OPTIONS.keys()),
            width=30,
        )
        self.address_mode_combobox.grid(row=0, column=1, sticky="w")

        self.address_run_button = ttk.Button(
            frame, text="Esegui divisione indirizzi", command=self._run_address_script
        )
        self.address_run_button.grid(row=4, column=0, pady=(12, 0), sticky="e")

    # ------------------------------------------------------------------ Helpers
    def _friendly_name(self, path: str | None) -> str:
        if not path:
            return "(nessun file)"
        return os.path.basename(path)

    def _add_unisci_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Seleziona file da unire", filetypes=EXCEL_FILETYPES
        )
        for path in paths:
            normalized = os.path.abspath(path)
            if normalized not in self.unisci_files:
                self.unisci_files.append(normalized)
        self._refresh_unisci_list()

    def _remove_unisci_selected(self) -> None:
        selection = self.unisci_listbox.curselection()
        if not selection:
            messagebox.showinfo("Rimozione file", "Seleziona un elemento da rimuovere.")
            return
        index = selection[0]
        del self.unisci_files[index]
        self._refresh_unisci_list()

    def _clear_unisci_files(self) -> None:
        self.unisci_files.clear()
        self._refresh_unisci_list()

    def _refresh_unisci_list(self) -> None:
        self.unisci_listbox.delete(0, tk.END)
        for path in self.unisci_files:
            self.unisci_listbox.insert(tk.END, path)

    def _add_convert_files(self) -> None:
        target, filetypes = self._get_convert_config()
        dialog_title = (
            "Seleziona file Excel da convertire in CSV"
            if target == "csv"
            else "Seleziona file CSV da convertire in Excel"
        )
        paths = filedialog.askopenfilenames(title=dialog_title, filetypes=filetypes)
        if not paths:
            return
        updated = False
        for path in paths:
            normalized = os.path.abspath(path)
            if normalized not in self.convert_files:
                self.convert_files.append(normalized)
                updated = True
        if updated:
            self._refresh_convert_list()

    def _remove_convert_selected(self) -> None:
        if not self.convert_listbox:
            return
        selection = self.convert_listbox.curselection()
        if not selection:
            messagebox.showinfo("Converti file", "Seleziona un elemento da rimuovere.")
            return
        index = selection[0]
        del self.convert_files[index]
        self._refresh_convert_list()

    def _clear_convert_files(self) -> None:
        self.convert_files.clear()
        self._refresh_convert_list()

    def _refresh_convert_list(self) -> None:
        if not self.convert_listbox:
            return
        self.convert_listbox.delete(0, tk.END)
        for path in self.convert_files:
            self.convert_listbox.insert(tk.END, path)

    def _get_convert_config(self):
        display = self.convert_direction_var.get()
        target, filetypes = CONVERT_DIRECTION_OPTIONS.get(
            display, next(iter(CONVERT_DIRECTION_OPTIONS.values()))
        )
        return target, filetypes

    def _on_convert_direction_changed(self, *_: object) -> None:
        target, _ = self._get_convert_config()
        if self.convert_delimiter_entry:
            state = tk.NORMAL if target == "csv" else tk.DISABLED
            self.convert_delimiter_entry.config(state=state)

    def _choose_split_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Seleziona file da dividere", filetypes=SPLIT_FILETYPES
        )
        if path:
            absolute = os.path.abspath(path)
            self.split_file_var.set(absolute)
            self.split_file_label.config(text=self._friendly_name(absolute))
            self._request_split_suggestion(absolute)
        else:
            self._split_suggestion_file = None
            self.split_suggestion_var.set("")

    def _choose_merge_file(self, index: int) -> None:
        path = filedialog.askopenfilename(
            title="Seleziona file Excel", filetypes=EXCEL_FILETYPES
        )
        if path:
            absolute = os.path.abspath(path)
            self.merge_files[index] = absolute
            self.merge_labels[index].config(text=self._friendly_name(absolute))
            self._update_merge_primary_options()

    def _update_merge_primary_options(self) -> None:
        options = [AUTO_PRIMARY_OPTION]
        self.merge_primary_options = {AUTO_PRIMARY_OPTION: AUTO_PRIMARY_OPTION}
        for path in self.merge_files:
            if path:
                display = os.path.basename(path)
                options.append(display)
                self.merge_primary_options[display] = os.path.basename(path)
        current = self.merge_primary_var.get()
        if current not in options:
            self.merge_primary_var.set(AUTO_PRIMARY_OPTION)
        self.merge_primary_combobox.configure(values=options)

    def _choose_compare_file(self, index: int) -> None:
        path = filedialog.askopenfilename(
            title="Seleziona file con indirizzi", filetypes=EXCEL_FILETYPES
        )
        if path:
            absolute = os.path.abspath(path)
            self.compare_files[index] = absolute
            self.compare_labels[index].config(text=self._friendly_name(absolute))

    def _choose_compare_map_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Seleziona CSV equivalenze comuni", filetypes=CSV_FILETYPES
        )
        if path:
            absolute = os.path.abspath(path)
            self.compare_map_var.set(absolute)
            if self.compare_map_label:
                self.compare_map_label.config(text=self._friendly_name(absolute))

    def _clear_compare_map_file(self) -> None:
        self.compare_map_var.set("")
        if self.compare_map_label:
            self.compare_map_label.config(text="(nessun file)")

    def _choose_address_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Seleziona file con indirizzi", filetypes=EXCEL_FILETYPES
        )
        if path:
            absolute = os.path.abspath(path)
            self.address_file_var.set(absolute)
            self.address_file_label.config(text=self._friendly_name(absolute))

    def _open_directory(self, directory: str) -> None:
        if not os.path.isdir(directory):
            os.makedirs(directory, exist_ok=True)
        try:
            if sys.platform.startswith("darwin"):
                subprocess.Popen(["open", directory])
            elif os.name == "nt":
                subprocess.Popen(["explorer", directory])
            else:
                subprocess.Popen(["xdg-open", directory])
        except Exception as exc:  # pragma: no cover - dipende dal sistema operativo
            messagebox.showerror("Errore", f"Impossibile aprire la cartella: {exc}")

    # ------------------------------------------------------------------ Run logic
    def _run_unisci_script(self) -> None:
        if not self.unisci_files:
            messagebox.showwarning("Unisci file", "Seleziona almeno un file da unire.")
            return
        output_name = self.unisci_output_var.get().strip() or "unione.xlsx"
        args = ["--files", *self.unisci_files, "--output-name", output_name]
        self._execute_script("Unisci file", args, self.unisci_run_button)

    def _run_split_script(self) -> None:
        file_path = self.split_file_var.get().strip()
        if not file_path:
            messagebox.showwarning("Dividi file", "Seleziona il file da dividere.")
            return
        try:
            chunk_size = int(self.chunk_size_var.get())
        except ValueError:
            messagebox.showerror("Dividi file", "Il numero di righe deve essere un intero.")
            return
        if chunk_size <= 0:
            messagebox.showerror("Dividi file", "Il numero di righe deve essere positivo.")
            return
        header_display = self.split_header_mode_var.get()
        header_mode = SPLIT_HEADER_OPTIONS.get(header_display, "repeat")
        args = ["--file", file_path, "--chunk-size", str(chunk_size), "--header-mode", header_mode]
        self._execute_script("Dividi file", args, self.split_run_button)

    def _run_merge_script(self) -> None:
        if not all(self.merge_files):
            messagebox.showwarning("Merge file", "Seleziona entrambi i file da confrontare.")
            return
        args = ["--files", *(self.merge_files[0:2])]
        primary_display = self.merge_primary_var.get()
        primary_value = self.merge_primary_options.get(primary_display)
        if primary_value and primary_value != AUTO_PRIMARY_OPTION:
            args.extend(["--primary", primary_value])
        self._execute_script("Merge file", args, self.merge_run_button)

    def _run_compare_script(self) -> None:
        if not all(self.compare_files):
            messagebox.showwarning(
                "Confronta indirizzi", "Seleziona entrambi i file da confrontare."
            )
            return
        if not self.compare_run_button:
            return
        mode_label = self.compare_mode_var.get()
        mode_value = COMPARE_MODE_OPTIONS.get(mode_label, "compatto")
        args = ["--files", *(self.compare_files[0:2]), "--mode", mode_value]
        comune_map = self.compare_map_var.get().strip()
        if comune_map:
            args.extend(["--comune-map", comune_map])
        self._execute_script("Confronta indirizzi", args, self.compare_run_button)

    def _run_convert_script(self) -> None:
        if not self.convert_files:
            messagebox.showwarning("Converti file", "Seleziona almeno un file da convertire.")
            return
        if not self.convert_run_button:
            return
        target, _ = self._get_convert_config()
        args = ["--files", *self.convert_files, "--to", target]
        if target == "csv":
            delimiter = self.convert_csv_delimiter_var.get().strip() or ","
            args.extend(["--csv-delimiter", delimiter])
        self._execute_script("Converti file", args, self.convert_run_button)

    def _run_address_script(self) -> None:
        file_path = self.address_file_var.get().strip()
        if not file_path:
            messagebox.showwarning(
                "Dividi indirizzi", "Seleziona il file contenente 'indirizzo_completo'."
            )
            return
        mode_label = self.address_mode_var.get()
        mode_value = ADDRESS_MODE_OPTIONS.get(mode_label, "siatel")
        args = ["--file", file_path, "--mode", mode_value]
        self._execute_script("Dividi indirizzi", args, self.address_run_button)

    def _execute_script(self, script_key: str, args: list[str], button: ttk.Button) -> None:
        script_filename = SCRIPTS.get(script_key)
        if not script_filename:
            messagebox.showerror("Errore", f"Script non riconosciuto: {script_key}")
            return
        script_path = os.path.join(BASE_DIR, script_filename)
        if not os.path.isfile(script_path):
            messagebox.showerror("Errore", f"Il file {script_path} non esiste.")
            return

        button.config(state=tk.DISABLED)
        self.append_log(f"> {script_key} {args}\n")

        thread = threading.Thread(
            target=self._run_script_thread,
            args=(script_key, script_path, args, button),
            daemon=True,
        )
        thread.start()

    def _request_split_suggestion(self, path: str) -> None:
        self._split_suggestion_file = path
        self.split_suggestion_var.set("Calcolo suggerimento per 3 MB…")
        thread = threading.Thread(
            target=self._compute_split_suggestion, args=(path,), daemon=True
        )
        thread.start()

    def _compute_split_suggestion(self, path: str) -> None:
        suggestion = self._generate_split_suggestion(path)
        self.after(0, self._apply_split_suggestion, path, suggestion)

    def _apply_split_suggestion(self, path: str, message: str) -> None:
        if path != self._split_suggestion_file:
            return
        self.split_suggestion_var.set(message)

    def _generate_split_suggestion(self, path: str) -> str:
        try:
            file_size = os.path.getsize(path)
        except OSError:
            return "Suggerimento 3 MB: non disponibile."
        data_rows = self._count_data_rows(path)
        if not data_rows:
            return "Suggerimento 3 MB: non disponibile."
        average_row_bytes = file_size / max(data_rows, 1)
        if average_row_bytes <= 0:
            return "Suggerimento 3 MB: non disponibile."

        suggested_rows = max(int(SPLIT_TARGET_BYTES / average_row_bytes), 1)
        suggested_rows = min(suggested_rows, data_rows)
        approx_chunks = max(math.ceil(data_rows / suggested_rows), 1)
        approx_size_mb = (average_row_bytes * suggested_rows) / (1024 * 1024)

        return (
            f"Suggerimento 3 MB: ~{suggested_rows} righe per file "
            f"(~{approx_chunks} file totali, ~{approx_size_mb:.1f} MB ciascuno)"
        )

    def _count_data_rows(self, path: str) -> int | None:
        ext = os.path.splitext(path)[1].lower()
        if ext == ".csv":
            return self._count_csv_rows(path)
        if ext in {".xlsx", ".xlsm", ".xls"}:
            return self._count_excel_rows(path)
        return None

    def _count_csv_rows(self, path: str) -> int | None:
        try:
            with open(path, "rb") as handle:
                total_rows = sum(1 for _ in handle)
        except OSError:
            return None
        data_rows = total_rows - 1 if total_rows > 0 else 0
        return data_rows if data_rows > 0 else None

    def _count_excel_rows(self, path: str) -> int | None:
        try:
            import openpyxl

            workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
            sheet = workbook.active
            max_row = sheet.max_row or 0
            workbook.close()
            data_rows = max(max_row - 1, 0)
            if data_rows > 0:
                return data_rows
        except Exception:
            pass

        try:
            import pandas as pd

            dataframe = pd.read_excel(path, usecols=[0])
            data_rows = len(dataframe)
            return data_rows if data_rows > 0 else None
        except Exception:
            return None

    def _run_script_thread(
        self, script_key: str, script_path: str, args: list[str], button: ttk.Button
    ) -> None:
        try:
            result = subprocess.run(
                [sys.executable, script_path, *args],
                capture_output=True,
                text=True,
                cwd=BASE_DIR,
            )
            success = result.returncode == 0
            stdout = result.stdout.strip()
            stderr = result.stderr.strip()
        except Exception as exc:  # pragma: no cover - dipende dal sistema host
            success = False
            stdout = ""
            stderr = str(exc)

        self.after(
            0,
            self._on_script_complete,
            script_key,
            button,
            success,
            stdout,
            stderr,
            os.path.basename(script_path),
        )

    def _on_script_complete(
        self,
        script_key: str,
        button: ttk.Button,
        success: bool,
        stdout: str,
        stderr: str,
        script_name: str,
    ) -> None:
        status_text = "completata" if success else "fallita"
        self.append_log(f"Esecuzione {status_text} per {script_name}.\n")
        if stdout:
            self.append_log(f"[STDOUT]\n{stdout}\n")
        if stderr:
            self.append_log(f"[STDERR]\n{stderr}\n")

        if success:
            messagebox.showinfo(script_key, f"{script_key} eseguito correttamente.")
        else:
            messagebox.showerror(script_key, f"Errore durante l'esecuzione di {script_key}.")

        button.config(state=tk.NORMAL)

    def append_log(self, message: str) -> None:
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _center_window(self) -> None:
        self.update_idletasks()
        width = self.winfo_width() or self.winfo_reqwidth()
        height = self.winfo_height() or self.winfo_reqheight()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.geometry(f"{width}x{height}+{x}+{y}")


def main() -> None:
    ensure_directories_exist()
    app = ScriptRunnerGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
