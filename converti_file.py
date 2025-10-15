import argparse
import csv
import os
from typing import Iterable, List, Sequence

import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
EXCEL_EXTS = {".xlsx", ".xlsm", ".xls"}
CSV_EXTS = {".csv"}
SUPPORTED_EXTS = EXCEL_EXTS | CSV_EXTS
TARGET_CHOICES = ("csv", "excel")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Converte uno o più file tra i formati Excel e CSV."
    )
    parser.add_argument(
        "--files",
        nargs="+",
        help=(
            "Percorsi dei file da convertire. Se assente, vengono utilizzati i file "
            "presenti nella cartella 'input' compatibili con la conversione richiesta."
        ),
    )
    parser.add_argument(
        "--to",
        choices=TARGET_CHOICES,
        help=(
            "Formato di destinazione: 'csv' per convertire file Excel in CSV, "
            "'excel' per convertire file CSV in Excel. "
            "Se omesso, viene dedotto dal tipo dei file specificati."
        ),
    )
    parser.add_argument(
        "--output-dir",
        default=OUTPUT_DIR,
        help="Cartella di destinazione per i file convertiti (default: cartella 'output').",
    )
    parser.add_argument(
        "--csv-delimiter",
        default=",",
        help="Delimitatore da utilizzare per i CSV generati (default: ',').",
    )
    return parser.parse_args()


def _collect_files_from_input(source_type: str) -> List[str]:
    directory = INPUT_DIR
    if not os.path.isdir(directory):
        raise SystemExit(f"Errore: la cartella di input non esiste: {directory}")

    expected_exts = EXCEL_EXTS if source_type == "excel" else CSV_EXTS
    files = [
        os.path.join(directory, entry)
        for entry in sorted(os.listdir(directory))
        if os.path.splitext(entry)[1].lower() in expected_exts
    ]

    if not files:
        readable = ", ".join(sorted(expected_exts))
        raise SystemExit(
            "Errore: nessun file compatibile trovato nella cartella 'input'. "
            f"Estensioni attese: {readable}"
        )
    return files


def _validate_files(files: Iterable[str]) -> List[str]:
    validated: List[str] = []
    for path in files:
        absolute = os.path.abspath(path)
        if not os.path.isfile(absolute):
            raise SystemExit(f"Errore: il file '{path}' non esiste.")
        ext = os.path.splitext(absolute)[1].lower()
        if ext not in SUPPORTED_EXTS:
            readable = ", ".join(sorted(SUPPORTED_EXTS))
            raise SystemExit(
                f"Errore: il file '{path}' non ha un'estensione supportata "
                f"({readable})."
            )
        validated.append(absolute)
    return validated


def _detect_target_mode(files: Sequence[str], explicit: str | None) -> str:
    if explicit:
        return explicit

    excel_files = [path for path in files if os.path.splitext(path)[1].lower() in EXCEL_EXTS]
    csv_files = [path for path in files if os.path.splitext(path)[1].lower() in CSV_EXTS]

    if excel_files and not csv_files:
        return "csv"
    if csv_files and not excel_files:
        return "excel"

    raise SystemExit(
        "Errore: impossibile dedurre il formato di destinazione. "
        "Specifica il parametro '--to'."
    )


def _ensure_output_directory(directory: str) -> str:
    absolute = os.path.abspath(directory)
    os.makedirs(absolute, exist_ok=True)
    return absolute


def _detect_csv_delimiter(file_path: str) -> str:
    try:
        with open(file_path, "r", encoding="utf-8-sig", newline="") as handle:
            sample = handle.read(4096)
            if not sample:
                return ","
            dialect = csv.Sniffer().sniff(sample, delimiters=[",", ";", "\t", "|", ":"])
            return dialect.delimiter
    except (OSError, csv.Error):
        return ","


def _convert_excel_to_csv(file_path: str, output_dir: str, delimiter: str) -> str:
    try:
        dataframe = pd.read_excel(file_path, header=0)
    except Exception as exc:  # pragma: no cover - dipende dai file utente
        raise SystemExit(f"Errore durante la lettura di '{file_path}': {exc}") from exc

    if dataframe.empty:
        raise SystemExit(f"Errore: il file '{file_path}' non contiene dati.")

    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_path = os.path.join(output_dir, f"{base_name}.csv")

    try:
        dataframe.to_csv(
            output_path,
            index=False,
            sep=delimiter,
            float_format="%.15g",  # evita suffisso .0 per valori numerici interi
        )
    except Exception as exc:  # pragma: no cover - dipende dai file utente
        raise SystemExit(f"Errore durante il salvataggio di '{output_path}': {exc}") from exc

    return output_path


def _convert_csv_to_excel(file_path: str, output_dir: str) -> str:
    delimiter = _detect_csv_delimiter(file_path)
    try:
        dataframe = pd.read_csv(file_path, sep=delimiter, engine="python")
    except Exception as exc:  # pragma: no cover - dipende dai file utente
        raise SystemExit(f"Errore durante la lettura di '{file_path}': {exc}") from exc

    if dataframe.empty:
        raise SystemExit(f"Errore: il file '{file_path}' non contiene dati.")

    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_path = os.path.join(output_dir, f"{base_name}.xlsx")

    try:
        dataframe.to_excel(output_path, index=False)
    except Exception as exc:  # pragma: no cover - dipende dai file utente
        raise SystemExit(f"Errore durante il salvataggio di '{output_path}': {exc}") from exc

    return output_path


def main() -> None:
    args = parse_args()
    csv_delimiter = args.csv_delimiter

    if args.files:
        files = _validate_files(args.files)
    else:
        if not args.to:
            raise SystemExit(
                "Errore: quando non si specificano file in ingresso è necessario usare '--to'."
            )
        source_type = "excel" if args.to == "csv" else "csv"
        files = _collect_files_from_input(source_type)

    target = _detect_target_mode(files, args.to)
    output_dir = _ensure_output_directory(args.output_dir)

    print("File convertiti:")
    for file_path in files:
        ext = os.path.splitext(file_path)[1].lower()
        if target == "csv" and ext not in EXCEL_EXTS:
            raise SystemExit(
                f"Errore: il file '{file_path}' non è un file Excel, ma è richiesta la conversione in CSV."
            )
        if target == "excel" and ext not in CSV_EXTS:
            raise SystemExit(
                f"Errore: il file '{file_path}' non è un file CSV, ma è richiesta la conversione in Excel."
            )

        if target == "csv":
            output_path = _convert_excel_to_csv(file_path, output_dir, csv_delimiter)
        else:
            output_path = _convert_csv_to_excel(file_path, output_dir)

        print(f" - {file_path} -> {output_path}")


if __name__ == "__main__":
    main()
