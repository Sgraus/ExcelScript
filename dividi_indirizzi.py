import argparse
import os
import sys
from typing import Literal

import pandas as pd

from dividi_indirizzi_helper import (
    dividi_indirizzo_siatel,
    dividi_indirizzo_siatel_compatto,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
SUPPORTED_EXTS = {".xlsx", ".xls", ".xlsm"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Divide gli indirizzi presenti nella colonna 'indirizzo_completo' "
            "di un file Excel e aggiunge le colonne risultanti."
        )
    )
    parser.add_argument(
        "--mode",
        choices=("siatel", "compatto"),
        default="siatel",
        help="Seleziona la modalità di divisione: 'siatel' (dettagliata) o 'compatto'.",
    )
    return parser.parse_args()


def _select_input_file() -> str:
    if not os.path.isdir(INPUT_DIR):
        raise SystemExit(f"Errore: la cartella di input non esiste: {INPUT_DIR}")

    entries = [
        f
        for f in os.listdir(INPUT_DIR)
        if os.path.splitext(f)[1].lower() in SUPPORTED_EXTS
    ]
    if not entries:
        raise SystemExit(
            "Errore: nessun file Excel trovato nella cartella 'input'. "
            "Formati supportati: .xlsx, .xlsm, .xls"
        )
    if len(entries) > 1:
        raise SystemExit(
            f"Errore: trovati {len(entries)} file Excel nella cartella 'input'. "
            "Lascia un solo file da elaborare."
        )
    return os.path.join(INPUT_DIR, entries[0])


def dividi_indirizzi(
    dataframe: pd.DataFrame, mode: Literal["siatel", "compatto"]
) -> pd.DataFrame:
    column_name = "indirizzo_completo"
    if column_name not in dataframe.columns:
        raise SystemExit(
            "Errore: la colonna 'indirizzo_completo' non è presente nel file da elaborare."
        )

    if mode == "siatel":
        splitter = dividi_indirizzo_siatel
        expected_columns = [
            "indirizzo",
            "civico",
            "scala",
            "interno",
            "piano",
            "estensione",
            "esito",
        ]
    else:
        splitter = dividi_indirizzo_siatel_compatto
        expected_columns = [
            "indirizzo_diviso",
            "civico_diviso",
            "specifica_civico_diviso",
            "esito",
        ]

    records = [splitter(value) for value in dataframe[column_name]]
    result_df = pd.DataFrame(records)

    for col in expected_columns:
        if col not in result_df.columns:
            result_df[col] = ""

    updated_df = dataframe.copy()
    for col in expected_columns:
        updated_df[col] = result_df[col].astype(str)

    return updated_df


def main() -> None:
    args = parse_args()
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    input_path = _select_input_file()

    try:
        df = pd.read_excel(input_path)
    except Exception as exc:  # pragma: no cover - dipende dai file utente
        raise SystemExit(f"Errore durante la lettura del file Excel: {exc}") from exc

    updated_df = dividi_indirizzi(df, args.mode)

    base_name = os.path.splitext(os.path.basename(input_path))[0]
    suffix = "_siatel" if args.mode == "siatel" else "_compatto"
    output_path = os.path.join(OUTPUT_DIR, f"{base_name}{suffix}.xlsx")

    try:
        updated_df.to_excel(output_path, index=False)
    except Exception as exc:  # pragma: no cover - dipende dai file utente
        raise SystemExit(f"Errore durante il salvataggio dell'output: {exc}") from exc

    print(f"File di input: {input_path}")
    print(f"Modalità di divisione: {args.mode}")
    print(f"File salvato in: {output_path}")


if __name__ == "__main__":
    main()
