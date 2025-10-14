import argparse
import csv
import os
from typing import Optional

import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
SUPPORTED_EXTS = {".xlsx", ".xlsm", ".xls", ".csv"}
HEADER_MODE_REPEAT = "repeat"
HEADER_MODE_FIRST_ONLY = "first-only"
HEADER_MODE_NONE = "none"
HEADER_MODE_CHOICES = (HEADER_MODE_REPEAT, HEADER_MODE_FIRST_ONLY, HEADER_MODE_NONE)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Divide un file in blocchi più piccoli da 100 righe (default)."
    )
    parser.add_argument(
        "--file",
        help=(
            "Percorso del file da suddividere. Se non specificato viene preso il "
            "primo file supportato presente nella cartella 'input'."
        ),
    )
    parser.add_argument(
        "--chunk-size",
        type=int,
        default=100,
        help="Numero di righe per ciascun file generato (default: 100).",
    )
    parser.add_argument(
        "--header-mode",
        choices=HEADER_MODE_CHOICES,
        default=HEADER_MODE_REPEAT,
        help=(
            "Gestione dell'intestazione nei file generati: "
            "'repeat' per ripeterla ovunque, "
            "'first-only' per mantenerla solo nel primo file, "
            "'none' per ometterla sempre."
        ),
    )
    return parser.parse_args()


def _select_file_from_input() -> str:
    if not os.path.isdir(INPUT_DIR):
        raise SystemExit(f"Errore: la cartella di input non esiste: {INPUT_DIR}")

    for entry in sorted(os.listdir(INPUT_DIR)):
        if os.path.splitext(entry)[1].lower() in SUPPORTED_EXTS:
            return os.path.join(INPUT_DIR, entry)

    raise SystemExit(
        "Errore: nessun file supportato trovato nella cartella 'input'. "
        f"Formati accettati: {', '.join(sorted(SUPPORTED_EXTS))}"
    )


def _resolve_file(path: Optional[str]) -> str:
    if path:
        absolute = os.path.abspath(path)
        if not os.path.isfile(absolute):
            raise SystemExit(f"Errore: il file '{path}' non esiste.")
        if os.path.splitext(absolute)[1].lower() not in SUPPORTED_EXTS:
            raise SystemExit(
                f"Errore: il file '{path}' non ha un'estensione supportata "
                f"({', '.join(sorted(SUPPORTED_EXTS))})."
            )
        return absolute
    return _select_file_from_input()


def _detect_csv_delimiter(file_path: str) -> str:
    """
    Prova a rilevare il delimitatore principale del file CSV
    restituendo una virgola come fallback.
    """
    try:
        with open(file_path, "r", newline="", encoding="utf-8-sig") as handle:
            sample = handle.read(4096)
            if not sample:
                return ","
            sniffer = csv.Sniffer()
            dialect = sniffer.sniff(sample, delimiters=[",", ";", "\t", "|", ":"])
            return dialect.delimiter
    except (OSError, csv.Error):
        return ","


def _load_dataframe(file_path: str) -> tuple[pd.DataFrame, str | None]:
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext == ".csv":
            delimiter = _detect_csv_delimiter(file_path)
            df = pd.read_csv(file_path, sep=delimiter, engine="python")
            return df, delimiter
        df = pd.read_excel(file_path, header=0)
        return df, None
    except Exception as exc:  # pragma: no cover - dipende dai file utente
        raise SystemExit(f"Errore durante la lettura di '{file_path}': {exc}") from exc


def main() -> None:
    args = parse_args()
    file_path = _resolve_file(args.file)

    if args.chunk_size <= 0:
        raise SystemExit("Errore: --chunk-size deve essere un numero positivo.")

    df, delimiter = _load_dataframe(file_path)
    if df.empty:
        raise SystemExit("Errore: il file selezionato è vuoto.")

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    base_name = os.path.basename(file_path)
    file_ext = os.path.splitext(base_name)[1].lower()
    chunk_size = args.chunk_size

    total_chunks = (len(df) + chunk_size - 1) // chunk_size

    for index, start in enumerate(range(0, len(df), chunk_size), start=1):
        chunk = df.iloc[start : start + chunk_size].copy()
        output_file = os.path.join(OUTPUT_DIR, f"{index} di {total_chunks} {base_name}")
        write_header = args.header_mode == HEADER_MODE_REPEAT or (
            args.header_mode == HEADER_MODE_FIRST_ONLY and index == 1
        )
        try:
            if file_ext == ".csv":
                chunk.to_csv(
                    output_file,
                    index=False,
                    header=write_header,
                    sep=delimiter or ",",
                )
            else:
                chunk.to_excel(output_file, index=False, header=write_header)
        except Exception as exc:  # pragma: no cover - dipende dai file utente
            raise SystemExit(f"Errore durante il salvataggio di '{output_file}': {exc}") from exc
        print(f"Creato: {output_file}")


if __name__ == "__main__":
    main()
