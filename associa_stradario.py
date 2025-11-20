import argparse
import os
from dataclasses import dataclass
from difflib import SequenceMatcher
from typing import List, Sequence

import pandas as pd

from confronta_indirizzi import (
    MODE_CONFIGS,
    ComuneResolver,
    _find_comune_columns,
    _first_non_empty,
    _load_comune_resolver,
    _normalize_text,
    _safe_string,
    _strip_street_prefix,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
SIMILARITY_THRESHOLD = 0.88
FUZZY_TIE_MARGIN = 0.02
ADDRESS_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}
STRADARIO_EXTENSIONS = ADDRESS_EXTENSIONS | {".csv"}
FLAG_COLUMNS = [
    "codice_via",
    "flag_codice_via",
    "flag_codice_via_motivo",
]


@dataclass(frozen=True)
class StradarioEntry:
    comune_key: str
    codice_via: str
    via_clean: str
    via_norm: str


class StradarioMatcher:
    def __init__(self, entries: Sequence[StradarioEntry]) -> None:
        self.entries_by_comune: dict[str, list[StradarioEntry]] = {}
        self.entries_by_exact: dict[tuple[str, str], list[StradarioEntry]] = {}
        for entry in entries:
            self.entries_by_comune.setdefault(entry.comune_key, []).append(entry)
            if entry.via_norm:
                key = (entry.comune_key, entry.via_norm)
                self.entries_by_exact.setdefault(key, []).append(entry)

    def match(self, comune_key: str, via_clean: str, via_norm: str) -> tuple[str, str]:
        if not comune_key:
            return "", "comune_assente"

        candidates = self.entries_by_comune.get(comune_key)
        if not candidates:
            return "", "comune_non_trovato"

        if not via_norm:
            return "", "via_assente"

        exact_entries = self.entries_by_exact.get((comune_key, via_norm))
        if exact_entries:
            if len(exact_entries) == 1:
                return exact_entries[0].codice_via, "exact"
            return "", "via_ambigua"

        best_ratio = 0.0
        best_entries: list[StradarioEntry] = []
        for entry in candidates:
            if not entry.via_clean:
                continue
            ratio = SequenceMatcher(
                None, entry.via_clean.casefold(), via_clean.casefold()
            ).ratio()
            if ratio > best_ratio + FUZZY_TIE_MARGIN:
                best_ratio = ratio
                best_entries = [entry]
            elif abs(ratio - best_ratio) <= FUZZY_TIE_MARGIN:
                best_entries.append(entry)

        if best_ratio < SIMILARITY_THRESHOLD:
            return "", "via_non_trovata"

        if len(best_entries) == 1:
            return best_entries[0].codice_via, "fuzzy"

        return "", "via_ambigua"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Associa ad ogni indirizzo il CODICE_VIA proveniente da un file stradario "
            "utilizzando un confronto flessibile su comune e via."
        )
    )
    parser.add_argument(
        "--addresses",
        required=True,
        help="File Excel con gli indirizzi già divisi (compatto o dettagliato).",
    )
    parser.add_argument(
        "--stradario",
        required=True,
        help="File Excel/CSV dello stradario con le colonne richieste.",
    )
    parser.add_argument(
        "--mode",
        choices=tuple(MODE_CONFIGS.keys()),
        default="compatto",
        help="Modalità del file indirizzi (compatto o dettagliato).",
    )
    parser.add_argument(
        "--output",
        help="Percorso del file Excel di output (predefinito nella cartella 'output').",
    )
    parser.add_argument(
        "--comune-map",
        help=(
            "Percorso CSV con colonne 'alias' e 'canonical' per gestire comuni equivalenti. "
            "Se omesso viene usato automaticamente 'comuni_equivalenze.csv' se presente."
        ),
    )
    return parser.parse_args()


def _validate_file(path: str, allowed_exts: set[str], description: str) -> str:
    absolute = os.path.abspath(path)
    if not os.path.isfile(absolute):
        raise SystemExit(f"Errore: il file {description} '{path}' non esiste.")
    ext = os.path.splitext(absolute)[1].lower()
    if ext not in allowed_exts:
        readable = ", ".join(sorted(allowed_exts))
        raise SystemExit(
            f"Errore: il file {description} '{path}' non ha un'estensione supportata ({readable})."
        )
    return absolute


def _read_table(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".csv":
            return pd.read_csv(path)
        return pd.read_excel(path)
    except Exception as exc:  # pragma: no cover - dipende da file utente
        raise SystemExit(f"Errore durante la lettura di '{os.path.basename(path)}': {exc}") from exc


def _canonical_comune(value: str, resolver: ComuneResolver | None) -> str:
    text = _safe_string(value)
    if not text:
        return ""
    return resolver.canonical(text) if resolver else _normalize_text(text)


def _build_stradario_entries(
    dataframe: pd.DataFrame,
    resolver: ComuneResolver | None,
) -> List[StradarioEntry]:
    if dataframe.empty:
        raise SystemExit("Errore: il file stradario non contiene righe utili.")

    column_map = {col.casefold(): col for col in dataframe.columns}
    comune_col = column_map.get("descrizione_comune")
    toponimo_col = column_map.get("toponimo")
    descrizione_col = column_map.get("descrizione_via")
    codice_col = column_map.get("codice_via")

    missing = [
        name
        for name, column in [
            ("DESCRIZIONE_COMUNE", comune_col),
            ("TOPONIMO", toponimo_col),
            ("DESCRIZIONE_VIA", descrizione_col),
            ("CODICE_VIA", codice_col),
        ]
        if column is None
    ]
    if missing:
        readable = ", ".join(missing)
        raise SystemExit(
            f"Errore: il file stradario deve contenere le colonne richieste: {readable}."
        )

    entries: list[StradarioEntry] = []
    for comune, toponimo, via_desc, codice in dataframe[
        [comune_col, toponimo_col, descrizione_col, codice_col]
    ].itertuples(index=False, name=None):
        comune_key = _canonical_comune(comune, resolver)
        codice_via = _safe_string(codice)
        if not comune_key or not codice_via:
            continue

        via_text = " ".join(
            value
            for value in [_safe_string(toponimo), _safe_string(via_desc)]
            if value
        ).strip()
        if not via_text:
            continue

        via_clean = _strip_street_prefix(via_text)
        via_norm = _normalize_text(via_clean)

        entries.append(
            StradarioEntry(
                comune_key=comune_key,
                codice_via=codice_via,
                via_clean=via_clean,
                via_norm=via_norm,
            )
        )

    if not entries:
        raise SystemExit("Errore: nessuna riga valida trovata nello stradario.")

    return entries


def _assign_codice_via(
    dataframe: pd.DataFrame,
    comune_columns: Sequence[str],
    via_column: str,
    matcher: StradarioMatcher,
    resolver: ComuneResolver | None,
) -> pd.DataFrame:
    def compute_row(row: pd.Series) -> pd.Series:
        comune_value = _first_non_empty(row, comune_columns)
        comune_key = _canonical_comune(comune_value, resolver)
        via_value = _safe_string(row.get(via_column, ""))
        via_clean = _strip_street_prefix(via_value) if via_value else ""
        via_norm = _normalize_text(via_clean) if via_clean else ""

        codice, status = matcher.match(comune_key, via_clean, via_norm)
        if status == "exact":
            flag = "certo"
            motivo = ""
        elif status == "fuzzy":
            flag = "simile"
            motivo = ""
        else:
            flag = "non_trovato"
            motivo = status or "via_non_trovata"

        return pd.Series([codice, flag, motivo], index=FLAG_COLUMNS)

    flags = dataframe.apply(compute_row, axis=1)
    result_df = dataframe.copy()
    for column in FLAG_COLUMNS:
        result_df[column] = flags[column]
    return result_df


def _ensure_output_path(path: str | None, address_path: str, mode: str) -> str:
    if path:
        absolute = os.path.abspath(path)
        directory = os.path.dirname(absolute)
        if directory and not os.path.isdir(directory):
            os.makedirs(directory, exist_ok=True)
        return absolute

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    stem = os.path.splitext(os.path.basename(address_path))[0]
    filename = f"associazione_stradario_{stem}_{mode}.xlsx"
    return os.path.join(OUTPUT_DIR, filename)


def main() -> None:
    args = parse_args()
    mode = args.mode
    if mode not in MODE_CONFIGS:
        raise SystemExit(f"Errore: modalità non supportata: {mode}")
    config = MODE_CONFIGS[mode]

    address_path = _validate_file(args.addresses, ADDRESS_EXTENSIONS, "indirizzi")
    stradario_path = _validate_file(args.stradario, STRADARIO_EXTENSIONS, "stradario")

    resolver = _load_comune_resolver(args.comune_map)

    address_df = _read_table(address_path)
    stradario_df = _read_table(stradario_path)

    comune_columns = _find_comune_columns(address_df, os.path.basename(address_path))
    if config.via_column not in address_df.columns:
        raise SystemExit(
            f"Errore: nel file indirizzi manca la colonna richiesta '{config.via_column}'."
        )

    entries = _build_stradario_entries(stradario_df, resolver)
    matcher = StradarioMatcher(entries)

    enriched_df = _assign_codice_via(
        address_df,
        comune_columns,
        config.via_column,
        matcher,
        resolver,
    )

    output_path = _ensure_output_path(args.output, address_path, mode)
    try:
        enriched_df.to_excel(output_path, index=False)
    except Exception as exc:  # pragma: no cover - dipende dai file utente
        raise SystemExit(f"Errore durante il salvataggio del file di output: {exc}") from exc

    flag_counts = enriched_df["flag_codice_via"].value_counts().to_dict()
    print(f"File indirizzi: {os.path.basename(address_path)} (modalità: {mode})")
    print(f"File stradario: {os.path.basename(stradario_path)}")
    print(f"Righe elaborate: {len(enriched_df)}")
    print(
        "Esiti rintraccio codice via: "
        + ", ".join(f"{key}={value}" for key, value in flag_counts.items())
    )
    print(f"File generato: {output_path}")


if __name__ == "__main__":
    main()
