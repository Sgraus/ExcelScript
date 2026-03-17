import argparse
import ast
import json
import os
from typing import Any

import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Estrae i campi JSON dalla colonna 'data' di un file Excel e li salva "
            "in nuove colonne."
        )
    )
    parser.add_argument(
        "--file",
        help=(
            "Percorso file Excel da elaborare. Se omesso, viene usato il primo file "
            "Excel trovato nella cartella 'input'."
        ),
    )
    parser.add_argument(
        "--column",
        default="data",
        help="Nome della colonna che contiene il JSON (default: data).",
    )
    parser.add_argument(
        "--output-name",
        default="parse_json.xlsx",
        help="Nome del file di output nella cartella di output (default: parse_json.xlsx).",
    )
    parser.add_argument(
        "--output-dir",
        default=OUTPUT_DIR,
        help="Cartella di destinazione (default: cartella 'output').",
    )
    return parser.parse_args()


def _resolve_input_file(path: str | None) -> str:
    if path:
        absolute = os.path.abspath(path)
        if not os.path.isfile(absolute):
            raise SystemExit(f"Errore: il file '{path}' non esiste.")
        extension = os.path.splitext(absolute)[1].lower()
        if extension not in EXCEL_EXTENSIONS:
            readable = ", ".join(sorted(EXCEL_EXTENSIONS))
            raise SystemExit(
                f"Errore: il file '{path}' non è un Excel supportato ({readable})."
            )
        return absolute

    if not os.path.isdir(INPUT_DIR):
        raise SystemExit(f"Errore: cartella input non trovata: {INPUT_DIR}")

    excel_files = [
        os.path.join(INPUT_DIR, entry)
        for entry in sorted(os.listdir(INPUT_DIR))
        if os.path.splitext(entry)[1].lower() in EXCEL_EXTENSIONS
    ]
    if not excel_files:
        raise SystemExit("Errore: nessun file Excel trovato nella cartella 'input'.")
    return excel_files[0]


def _normalize_json_key(prefix: str, key: str) -> str:
    normalized = key.strip().replace(" ", "_")
    return f"{prefix}.{normalized}" if prefix else normalized


def _flatten_json(value: Any, prefix: str = "") -> dict[str, Any]:
    if isinstance(value, dict):
        flattened: dict[str, Any] = {}
        for key, nested_value in value.items():
            new_key = _normalize_json_key(prefix, str(key))
            flattened.update(_flatten_json(nested_value, new_key))
        return flattened

    if isinstance(value, list):
        return {prefix: json.dumps(value, ensure_ascii=False)} if prefix else {"value": value}

    if prefix:
        return {prefix: value}
    return {"value": value}


def _parse_json_cell(raw_value: Any) -> dict[str, Any]:
    if pd.isna(raw_value):
        return {}

    if isinstance(raw_value, dict):
        parsed = raw_value
    elif isinstance(raw_value, str):
        text = raw_value.strip()
        if not text:
            return {}
        try:
            parsed = json.loads(text)
        except json.JSONDecodeError:
            try:
                parsed = ast.literal_eval(text)
            except (ValueError, SyntaxError):
                return {"raw_json": text}
    else:
        return {"raw_json": str(raw_value)}

    return _flatten_json(parsed)


def main() -> None:
    args = parse_args()
    input_file = _resolve_input_file(args.file)

    try:
        dataframe = pd.read_excel(input_file)
    except Exception as exc:  # pragma: no cover - dipende dai file utente
        raise SystemExit(f"Errore durante la lettura del file Excel: {exc}") from exc

    target_column = None
    for column in dataframe.columns:
        if str(column).strip().lower() == args.column.strip().lower():
            target_column = column
            break

    if target_column is None:
        raise SystemExit(
            f"Errore: colonna '{args.column}' non trovata. Colonne disponibili: "
            f"{', '.join(map(str, dataframe.columns))}"
        )

    parsed_rows = [_parse_json_cell(value) for value in dataframe[target_column]]
    discovered_keys = sorted({key for row in parsed_rows for key in row})

    if not discovered_keys:
        raise SystemExit(
            "Errore: nessuna chiave JSON trovata nella colonna indicata."
        )

    renamed_columns: dict[str, str] = {}
    for key in discovered_keys:
        column_name = key
        if column_name in dataframe.columns:
            column_name = f"json.{key}"
        renamed_columns[key] = column_name

    normalized = pd.DataFrame(
        [{renamed_columns[key]: row.get(key) for key in discovered_keys} for row in parsed_rows]
    )
    result = pd.concat([dataframe, normalized], axis=1)

    output_dir = os.path.abspath(args.output_dir)
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, args.output_name)

    try:
        result.to_excel(output_path, index=False)
    except Exception as exc:  # pragma: no cover - dipende dai file utente
        raise SystemExit(f"Errore durante il salvataggio dell'output: {exc}") from exc

    print(f"File elaborato: {input_file}")
    print(f"Colonna JSON: {target_column}")
    print(f"Nuove colonne aggiunte: {len(discovered_keys)}")
    print(f"Output salvato in: {output_path}")


if __name__ == "__main__":
    main()
