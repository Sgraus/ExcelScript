import argparse
import os
from typing import Iterable, List, Tuple

import pandas as pd

# Imposta la cartella contenente i file Excel
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
ALLOWED_EXT = {".xlsx", ".xls"}


def find_excel_files(folder: str) -> List[str]:
    """
    Restituisce la lista dei file Excel validi (xlsx/xls) nella cartella,
    escludendo file temporanei (prefisso '~$').
    """
    if not os.path.isdir(folder):
        raise SystemExit(f"Errore: la cartella di input non esiste: {folder}")
    files = [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if os.path.isfile(os.path.join(folder, f))
    ]
    excel_files = [
        path
        for path in files
        if os.path.splitext(path)[1].lower() in ALLOWED_EXT
        and not os.path.basename(path).startswith("~$")
    ]
    return excel_files


def read_excel_or_fail(path: str) -> pd.DataFrame:
    """
    Legge un file Excel e controlla la presenza della colonna 'match'.
    """
    try:
        df = pd.read_excel(path)
    except Exception as e:  # pragma: no cover - dipende dai file utente
        raise SystemExit(f"Errore durante la lettura di '{os.path.basename(path)}': {e}")

    if "match" not in df.columns:
        raise SystemExit(
            f"Errore: la colonna 'match' manca nel file: {os.path.basename(path)}"
        )

    # Normalizza la colonna 'match' a stringa, rimuovendo spazi
    df["match"] = df["match"].astype(str).str.strip()
    return df


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Esegue un merge LEFT tra due file Excel sulla colonna 'match'. "
            "Se non vengono specificati i file, usa quelli presenti nella cartella 'input'."
        )
    )
    parser.add_argument(
        "--files",
        nargs=2,
        metavar=("FILE_A", "FILE_B"),
        help="Percorsi dei due file Excel da unire.",
    )
    parser.add_argument(
        "--primary",
        help=(
            "Nome o percorso del file da usare come lato completo del merge. "
            "Se non specificato viene scelto automaticamente il file con meno righe."
        ),
    )
    return parser.parse_args()


def _validate_files(files: Iterable[str]) -> List[str]:
    validated: List[str] = []
    for raw_path in files:
        absolute = os.path.abspath(raw_path)
        if not os.path.isfile(absolute):
            raise SystemExit(f"Errore: il file '{raw_path}' non esiste.")
        if os.path.splitext(absolute)[1].lower() not in ALLOWED_EXT:
            raise SystemExit(
                f"Errore: il file '{raw_path}' non ha un'estensione supportata "
                f"({', '.join(sorted(ALLOWED_EXT))})."
            )
        validated.append(absolute)
    return validated


def _resolve_files(args: argparse.Namespace) -> List[str]:
    if args.files:
        return _validate_files(args.files)

    excel_files = find_excel_files(INPUT_DIR)
    if len(excel_files) != 2:
        raise SystemExit(
            "Errore: per eseguire il merge servono esattamente 2 file Excel. "
            "Specificare i file tramite --files oppure lasciare nella cartella 'input' "
            "solo i due file da unire."
        )
    return [os.path.abspath(path) for path in excel_files]


def _match_candidate(candidate: str, path: str) -> bool:
    filename = os.path.basename(path)
    base_no_ext = os.path.splitext(filename)[0]
    normalized_set = {
        os.path.normcase(candidate.strip()),
        os.path.normcase(os.path.basename(candidate.strip())),
        os.path.normcase(os.path.splitext(os.path.basename(candidate.strip()))[0]),
    }
    return os.path.normcase(filename) in normalized_set or os.path.normcase(
        base_no_ext
    ) in normalized_set or os.path.normcase(os.path.abspath(path)) in normalized_set


def _pick_base_table(
    data: List[Tuple[str, pd.DataFrame]], primary_choice: str | None
) -> Tuple[Tuple[str, pd.DataFrame], Tuple[str, pd.DataFrame]]:
    base_entry = None
    other_entry = None

    if primary_choice:
        for entry in data:
            if _match_candidate(primary_choice, entry[0]):
                base_entry = entry
                break
        if base_entry is None:
            raise SystemExit(
                f"Errore: il file principale '{primary_choice}' non coincide con i file selezionati."
            )
        other_entry = next((e for e in data if e is not base_entry), None)
        if other_entry is None:  # pragma: no cover - situazione impossibile
            raise SystemExit("Errore interno: non è stato trovato il secondo file.")
    else:
        data_sorted = sorted(data, key=lambda item: len(item[1]))
        base_entry, other_entry = data_sorted[0], data_sorted[1]

    return base_entry, other_entry


def main() -> None:
    args = parse_args()
    primary_choice = args.primary.strip() if args.primary else None

    excel_files = _resolve_files(args)

    # Leggi entrambi i file e valida la colonna 'match'
    data = [(path, read_excel_or_fail(path)) for path in excel_files]

    base_entry, other_entry = _pick_base_table(data, primary_choice)

    base_path, base_df = base_entry
    other_path, other_df = other_entry
    base_name = os.path.splitext(os.path.basename(base_path))[0]
    other_name = os.path.splitext(os.path.basename(other_path))[0]

    # Merge LEFT usando come "sinistra" la tabella più piccola
    merged = base_df.merge(
        other_df,
        on="match",
        how="left",
        suffixes=(f"_{base_name}", f"_{other_name}"),
    )

    # Porta 'match' come prima colonna
    cols = ["match"] + [c for c in merged.columns if c != "match"]
    merged = merged.loc[:, cols]

    # Scrivi l'output
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out_path = os.path.join(OUTPUT_DIR, f"merge_{base_name}_{other_name}.xlsx")
    try:
        merged.to_excel(out_path, index=False)
    except Exception as e:  # pragma: no cover - dipende dai file utente
        raise SystemExit(f"Errore durante il salvataggio dell'output: {e}")

    if primary_choice:
        print(f"File principale selezionato manualmente: {os.path.basename(base_path)}")
    print(
        f"Merge completato. File base: {base_name} (righe: {len(base_df)}). "
        f"Altro file: {other_name} (righe: {len(other_df)})."
    )
    print(f"Totale righe in output: {len(merged)}")
    print(f"File salvato in: {out_path}")


if __name__ == "__main__":
    main()
