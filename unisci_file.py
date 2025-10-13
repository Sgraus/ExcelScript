import argparse
import os
from typing import Iterable, List

import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
SUPPORTED_EXTS = {".xlsx", ".xls", ".xlsm"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Unisce più file Excel in un unico file."
    )
    parser.add_argument(
        "--files",
        nargs="+",
        help=(
            "Percorsi dei file Excel da unire. Se non specificati, vengono usati tutti "
            "i file supportati presenti nella cartella 'input'."
        ),
    )
    parser.add_argument(
        "--output-name",
        default="unione.xlsx",
        help="Nome del file di output (verrà creato nella cartella 'output').",
    )
    return parser.parse_args()


def _collect_files_from_input() -> List[str]:
    if not os.path.isdir(INPUT_DIR):
        raise SystemExit(f"Errore: la cartella di input non esiste: {INPUT_DIR}")
    files = [
        os.path.join(INPUT_DIR, entry)
        for entry in os.listdir(INPUT_DIR)
        if os.path.splitext(entry)[1].lower() in SUPPORTED_EXTS
    ]
    if not files:
        raise SystemExit(
            "Errore: nessun file Excel trovato nella cartella 'input'. "
            f"Formati supportati: {', '.join(sorted(SUPPORTED_EXTS))}"
        )
    return files


def _validate_files(files: Iterable[str]) -> List[str]:
    validated: List[str] = []
    for path in files:
        absolute = os.path.abspath(path)
        if not os.path.isfile(absolute):
            raise SystemExit(f"Errore: il file '{path}' non esiste.")
        if os.path.splitext(absolute)[1].lower() not in SUPPORTED_EXTS:
            raise SystemExit(
                f"Errore: il file '{path}' non ha un'estensione supportata "
                f"({', '.join(sorted(SUPPORTED_EXTS))})."
            )
        validated.append(absolute)
    return validated


def main() -> None:
    args = parse_args()
    files = _collect_files_from_input() if not args.files else _validate_files(args.files)

    dataframes: List[pd.DataFrame] = []
    for file_path in files:
        try:
            df = pd.read_excel(file_path)
            dataframes.append(df)
        except Exception as exc:  # pragma: no cover - dipende dai file utente
            raise SystemExit(f"Errore durante la lettura di '{file_path}': {exc}") from exc

    if not dataframes:
        raise SystemExit("Errore: nessun dato è stato caricato dai file selezionati.")

    combined = pd.concat(dataframes, ignore_index=True)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_name = args.output_name.strip() or "unione.xlsx"
    if not output_name.lower().endswith((".xlsx", ".xlsm", ".xls")):
        output_name += ".xlsx"
    output_path = os.path.join(OUTPUT_DIR, output_name)

    try:
        combined.to_excel(output_path, index=False)
    except Exception as exc:  # pragma: no cover - dipende dai file utente
        raise SystemExit(f"Errore durante il salvataggio dell'output: {exc}") from exc

    print("File elaborati:")
    for file_path in files:
        print(f" - {file_path}")
    print(f"File di output salvato in: {output_path}")


if __name__ == "__main__":
    main()
