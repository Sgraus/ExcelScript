import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import messagebox, ttk

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
AUTO_PRIMARY_OPTION = "Automatico (più piccolo)"
ADDRESS_MODE_OPTIONS = {
    "Siatel (dettagliato)": "siatel",
    "Siatel compatto": "compatto",
}

SCRIPTS = {
    "Unisci file": "unisci_file.py",
    "Dividi file": "spilit_file.py",
    "Merge file": "marge_file.py",
    "Dividi indirizzi": "dividi_indirizzi.py",
}


def ensure_directories_exist() -> None:
    """Ensure that input and output directories are available."""
    for directory in (INPUT_DIR, OUTPUT_DIR):
        os.makedirs(directory, exist_ok=True)


class ScriptRunnerGUI(tk.Tk):
    """Simple GUI to inspect input/output folders and execute scripts."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Excel Script Runner")
        self.resizable(False, False)

        self.selected_script_key: str | None = None
        self.primary_file_var = tk.StringVar(value=AUTO_PRIMARY_OPTION)
        default_address_mode = next(iter(ADDRESS_MODE_OPTIONS))
        self.address_mode_var = tk.StringVar(value=default_address_mode)

        self._build_widgets()
        self.refresh_file_lists()
        self._center_window()

    def _build_widgets(self) -> None:
        main_frame = ttk.Frame(self, padding=10)
        main_frame.grid(row=0, column=0, sticky="nsew")

        # Script selection frame
        scripts_frame = ttk.LabelFrame(main_frame, text="Script disponibili")
        scripts_frame.grid(row=0, column=0, rowspan=2, padx=(0, 10), sticky="ns")

        self.scripts_list = tk.Listbox(scripts_frame, height=len(SCRIPTS))
        self.scripts_list.grid(row=0, column=0, padx=5, pady=5)
        for name in SCRIPTS:
            self.scripts_list.insert(tk.END, name)
        self.scripts_list.bind("<<ListboxSelect>>", self.on_script_select)

        self.run_button = ttk.Button(
            scripts_frame,
            text="Esegui script selezionato",
            command=self.execute_selected_script,
            state=tk.DISABLED,
        )
        self.run_button.grid(row=1, column=0, padx=5, pady=(0, 5), sticky="ew")

        self.merge_options_frame = ttk.LabelFrame(scripts_frame, text="Opzioni Merge")
        self.merge_options_frame.grid(row=2, column=0, padx=5, pady=(0, 5), sticky="ew")
        self.primary_combobox = ttk.Combobox(
            self.merge_options_frame,
            textvariable=self.primary_file_var,
            state="readonly",
            width=28,
        )
        self.primary_combobox.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        self.merge_options_frame.grid_remove()

        self.address_options_frame = ttk.LabelFrame(
            scripts_frame, text="Opzioni Dividi indirizzi"
        )
        self.address_options_frame.grid(
            row=3, column=0, padx=5, pady=(0, 5), sticky="ew"
        )
        self.address_mode_combobox = ttk.Combobox(
            self.address_options_frame,
            textvariable=self.address_mode_var,
            state="readonly",
            values=list(ADDRESS_MODE_OPTIONS.keys()),
            width=28,
        )
        self.address_mode_combobox.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        self.address_options_frame.grid_remove()

        # Input/output frame
        io_frame = ttk.Frame(main_frame)
        io_frame.grid(row=0, column=1, sticky="nsew")

        input_frame = ttk.LabelFrame(io_frame, text="File in input")
        input_frame.grid(row=0, column=0, padx=(0, 5), sticky="nsew")
        self.input_list = tk.Listbox(input_frame, width=40, height=10)
        self.input_list.grid(row=0, column=0, padx=5, pady=5)

        output_frame = ttk.LabelFrame(io_frame, text="File in output")
        output_frame.grid(row=0, column=1, padx=(5, 0), sticky="nsew")
        self.output_list = tk.Listbox(output_frame, width=40, height=10)
        self.output_list.grid(row=0, column=0, padx=5, pady=5)

        # Log frame
        log_frame = ttk.LabelFrame(main_frame, text="Log esecuzione")
        log_frame.grid(row=1, column=1, pady=(10, 0), sticky="nsew")
        self.log_text = tk.Text(log_frame, width=80, height=12, state=tk.DISABLED)
        self.log_text.grid(row=0, column=0, padx=5, pady=5)

    def on_script_select(self, event: tk.Event) -> None:  # type: ignore[override]
        selection = self.scripts_list.curselection()
        if not selection:
            self.selected_script_key = None
            self.run_button.config(state=tk.DISABLED)
            return
        index = selection[0]
        self.selected_script_key = self.scripts_list.get(index)
        self.run_button.config(state=tk.NORMAL)
        self.append_log(f"Script selezionato: {self.selected_script_key}\n")
        self.refresh_file_lists()
        if self.selected_script_key == "Merge file":
            self.merge_options_frame.grid()
            self._update_primary_file_options()
        else:
            self.merge_options_frame.grid_remove()
        if self.selected_script_key == "Dividi indirizzi":
            self.address_options_frame.grid()
        else:
            self.address_options_frame.grid_remove()

    def execute_selected_script(self) -> None:
        if not self.selected_script_key:
            messagebox.showwarning("Nessuno script", "Seleziona uno script da eseguire.")
            return
        script_path = os.path.join(BASE_DIR, SCRIPTS[self.selected_script_key])
        if not os.path.isfile(script_path):
            messagebox.showerror("Errore", f"Il file {script_path} non esiste.")
            return

        self.append_log(f"Esecuzione di: {self.selected_script_key}\n")
        extra_args: list[str] = []
        if self.selected_script_key == "Merge file":
            primary_choice = self.primary_file_var.get()
            if primary_choice and primary_choice != AUTO_PRIMARY_OPTION:
                extra_args = ["--primary", primary_choice]
                self.append_log(f"File principale selezionato: {primary_choice}\n")
        elif self.selected_script_key == "Dividi indirizzi":
            selected_label = self.address_mode_var.get()
            mode_value = ADDRESS_MODE_OPTIONS.get(selected_label, "siatel")
            extra_args = ["--mode", mode_value]
            self.append_log(f"Modalità dividi indirizzi: {selected_label}\n")
        self.run_button.config(state=tk.DISABLED)

        thread = threading.Thread(
            target=self._run_script, args=(script_path, extra_args), daemon=True
        )
        thread.start()

    def _run_script(self, script_path: str, extra_args: list[str]) -> None:
        try:
            result = subprocess.run(
                [sys.executable, script_path, *extra_args],
                capture_output=True,
                text=True,
                cwd=BASE_DIR,
            )
            success = result.returncode == 0
            stdout = result.stdout.strip()
            stderr = result.stderr.strip()
        except Exception as exc:  # pragma: no cover - unforeseen execution errors
            success = False
            stdout = ""
            stderr = str(exc)

        self.after(
            0,
            self._on_script_complete,
            success,
            stdout,
            stderr,
            os.path.basename(script_path),
        )

    def _on_script_complete(
        self, success: bool, stdout: str, stderr: str, script_name: str
    ) -> None:
        status = "completata" if success else "fallita"
        self.append_log(f"Esecuzione {status} per {script_name}.\n")
        if stdout:
            self.append_log(f"[STDOUT]\n{stdout}\n")
        if stderr:
            self.append_log(f"[STDERR]\n{stderr}\n")
        if success:
            messagebox.showinfo("Completato", f"{script_name} eseguito correttamente.")
        else:
            messagebox.showerror(
                "Errore", f"Si è verificato un errore eseguendo {script_name}."
            )

        self.run_button.config(state=tk.NORMAL)
        self.refresh_file_lists()

    def refresh_file_lists(self) -> None:
        self._populate_listbox(self.input_list, INPUT_DIR)
        self._populate_listbox(self.output_list, OUTPUT_DIR)
        self._update_primary_file_options()

    def _populate_listbox(self, listbox: tk.Listbox, directory: str) -> None:
        try:
            entries = sorted(os.listdir(directory))
        except FileNotFoundError:
            os.makedirs(directory, exist_ok=True)
            entries = []
        listbox.delete(0, tk.END)
        if entries:
            for entry in entries:
                listbox.insert(tk.END, entry)
        else:
            listbox.insert(tk.END, "(nessun file)")

    def append_log(self, message: str) -> None:
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _update_primary_file_options(self) -> None:
        if not hasattr(self, "primary_combobox"):
            return
        try:
            entries = [
                f
                for f in sorted(os.listdir(INPUT_DIR))
                if os.path.isfile(os.path.join(INPUT_DIR, f))
                and os.path.splitext(f)[1].lower() in {".xlsx", ".xls"}
            ]
        except FileNotFoundError:
            entries = []
        options = [AUTO_PRIMARY_OPTION] + entries
        current_value = self.primary_file_var.get()
        self.primary_combobox["values"] = options
        if current_value not in options:
            self.primary_file_var.set(AUTO_PRIMARY_OPTION)

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
    if SCRIPTS:
        app.scripts_list.selection_set(0)
        app.on_script_select(None)
    app.mainloop()


if __name__ == "__main__":
    main()
