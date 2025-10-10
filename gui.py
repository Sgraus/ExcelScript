import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import messagebox, ttk

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

SCRIPTS = {
    "Unisci file": "unisci_file.py",
    "Dividi file": "spilit_file.py",
    "Merge file": "marge_file.py",
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

        self._build_widgets()
        self.refresh_file_lists()

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

    def execute_selected_script(self) -> None:
        if not self.selected_script_key:
            messagebox.showwarning("Nessuno script", "Seleziona uno script da eseguire.")
            return
        script_path = os.path.join(BASE_DIR, SCRIPTS[self.selected_script_key])
        if not os.path.isfile(script_path):
            messagebox.showerror("Errore", f"Il file {script_path} non esiste.")
            return

        self.append_log(f"Esecuzione di: {self.selected_script_key}\n")
        self.run_button.config(state=tk.DISABLED)

        thread = threading.Thread(
            target=self._run_script, args=(script_path,), daemon=True
        )
        thread.start()

    def _run_script(self, script_path: str) -> None:
        try:
            result = subprocess.run(
                [sys.executable, script_path],
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


def main() -> None:
    ensure_directories_exist()
    app = ScriptRunnerGUI()
    if SCRIPTS:
        app.scripts_list.selection_set(0)
        app.on_script_select(None)
    app.mainloop()


if __name__ == "__main__":
    main()
