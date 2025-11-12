import argparse
import os
import re
from dataclasses import dataclass
from difflib import SequenceMatcher
from typing import Callable, Iterable, List, Sequence

import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
ALLOWED_EXT = {".xlsx", ".xls", ".xlsm"}
COMUNE_KEYWORD = "comune"
DEFAULT_COMUNE_MAP = os.path.join(BASE_DIR, "comuni_equivalenze.csv")

NORMALIZE_RE = re.compile(r"[^A-Z0-9]", re.UNICODE)
EXTRA_STRIP_CHARS = "\xa0\u2007\u202f\u200b\u200c\u200d\u2060\ufeff"


@dataclass(frozen=True)
class ModeConfig:
    address_columns: Sequence[str]
    via_column: str
    civico_column: str
    rest_columns: Sequence[str]


MODE_CONFIGS: dict[str, ModeConfig] = {
    "compatto": ModeConfig(
        address_columns=("indirizzo_diviso", "civico_diviso", "specifica_civico_diviso"),
        via_column="indirizzo_diviso",
        civico_column="civico_diviso",
        rest_columns=("specifica_civico_diviso",),
    ),
    "dettagliato": ModeConfig(
        address_columns=("indirizzo", "civico", "scala", "interno", "piano", "estensione"),
        via_column="indirizzo",
        civico_column="civico",
        rest_columns=("scala", "interno", "piano", "estensione"),
    ),
}


@dataclass
class PreparedTable:
    path: str
    label: str
    dataframe: pd.DataFrame
    comune_columns: Sequence[str]
    address_columns: Sequence[str]
    via_column: str
    civico_column: str
    rest_columns: Sequence[str]


@dataclass
class ComuneResolver:
    lookup: dict[str, str]

    def canonical(self, text: str) -> str:
        normalized = _normalize_text(text)
        if not normalized:
            return ""
        return self.lookup.get(normalized, normalized)


FLAG_COLUMNS = [
    "flag_comune",
    "flag_via",
    "flag_civico",
    "flag_dettagli",
    "flag_generale",
]

STREET_PREFIXES = {
    "VIA",
    "V",
    "VIALE",
    "VLE",
    "PIAZZA",
    "PZ",
    "PZZA",
    "PIAZZALE",
    "PLE",
    "LARGO",
    "LGO",
    "CORSO",
    "CRS",
    "STRADA",
    "STR",
    "SDA",
    "CONTRADA",
    "CTR",
    "LOCALITA",
    "LOC",
    "BORGO",
    "BGO",
    "VICOLO",
    "VICO",
    "SALITA",
    "SAL",
    "PIAZZETTA",
    "PZTTA",
    "TRAVERSA",
    "TRAV",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Confronta le colonne degli indirizzi di due file Excel con la stessa colonna 'match'. "
            "Produce un file con i dettagli degli indirizzi di entrambi i file e flag sintetici."
        )
    )
    parser.add_argument(
        "--files",
        nargs=2,
        metavar=("FILE_A", "FILE_B"),
        help=(
            "Percorsi dei due file Excel da confrontare. "
            "Se omesso vengono usati i due file presenti nella cartella 'input'."
        ),
    )
    parser.add_argument(
        "--mode",
        choices=tuple(MODE_CONFIGS.keys()),
        default="compatto",
        help="Seleziona il set di colonne da confrontare: 'compatto' o 'dettagliato'.",
    )
    parser.add_argument(
        "--output",
        help=(
            "Percorso del file Excel di output. "
            "Se non specificato viene creato nella cartella 'output'."
        ),
    )
    parser.add_argument(
        "--comune-map",
        help=(
            "Percorso di un CSV con colonne 'alias' e 'canonical' per forzare l'equivalenza dei comuni. "
            "Se presente, viene unito alla tabella predefinita 'comuni_equivalenze.csv' nella cartella dello script."
        ),
    )
    return parser.parse_args()


def find_excel_files(folder: str) -> List[str]:
    if not os.path.isdir(folder):
        raise SystemExit(f"Errore: la cartella di input non esiste: {folder}")
    files = [
        os.path.join(folder, entry)
        for entry in sorted(os.listdir(folder))
        if os.path.isfile(os.path.join(folder, entry))
    ]
    excel_files = [
        path
        for path in files
        if os.path.splitext(path)[1].lower() in ALLOWED_EXT
        and not os.path.basename(path).startswith("~$")
    ]
    return excel_files


def _validate_files(files: Iterable[str]) -> List[str]:
    validated: List[str] = []
    for raw in files:
        absolute = os.path.abspath(raw)
        if not os.path.isfile(absolute):
            raise SystemExit(f"Errore: il file '{raw}' non esiste.")
        if os.path.splitext(absolute)[1].lower() not in ALLOWED_EXT:
            readable = ", ".join(sorted(ALLOWED_EXT))
            raise SystemExit(
                f"Errore: il file '{raw}' non ha un'estensione supportata ({readable})."
            )
        validated.append(absolute)
    return validated


def _resolve_files(args: argparse.Namespace) -> List[str]:
    if args.files:
        return _validate_files(args.files)

    excel_files = find_excel_files(INPUT_DIR)
    if len(excel_files) != 2:
        raise SystemExit(
            "Errore: per il confronto servono esattamente due file Excel. "
            "Specificare i file con --files oppure lasciare nella cartella 'input' "
            "solo i due file da confrontare."
        )
    return [os.path.abspath(path) for path in excel_files]


def _read_dataframe(path: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
    except Exception as exc:  # pragma: no cover - dipende dai file utente
        raise SystemExit(f"Errore durante la lettura di '{os.path.basename(path)}': {exc}") from exc
    if "match" not in df.columns:
        raise SystemExit(
            f"Errore: il file '{os.path.basename(path)}' non contiene la colonna 'match'."
        )
    df = df.copy()
    df["match"] = df["match"].astype(str).str.strip()
    return df


def _normalize_label(path: str) -> str:
    stem = os.path.splitext(os.path.basename(path))[0]
    normalized = re.sub(r"\W+", "_", stem, flags=re.UNICODE).strip("_")
    return normalized or "file"


def _find_comune_columns(df: pd.DataFrame, filename: str) -> List[str]:
    comune_cols = [col for col in df.columns if COMUNE_KEYWORD in col.casefold()]
    if not comune_cols:
        raise SystemExit(
            f"Errore: il file '{filename}' non contiene colonne con la parola '{COMUNE_KEYWORD}'."
        )
    return comune_cols


def _prepare_table(path: str, config: ModeConfig) -> PreparedTable:
    dataframe = _read_dataframe(path)
    filename = os.path.basename(path)
    comune_columns = _find_comune_columns(dataframe, filename)
    missing = [col for col in config.address_columns if col not in dataframe.columns]
    if missing:
        readable = ", ".join(missing)
        raise SystemExit(
            f"Errore: nel file '{filename}' mancano le colonne richieste per la modalità selezionata: {readable}."
        )

    selected_cols = ["match"] + list(comune_columns) + list(config.address_columns)
    subset = dataframe.loc[:, selected_cols].copy()
    label = _normalize_label(path)
    rename_map = {
        col: col if col == "match" else f"{col}_{label}"
        for col in subset.columns
    }
    renamed = subset.rename(columns=rename_map)

    comune_renamed = [rename_map[col] for col in comune_columns]
    address_renamed = [rename_map[col] for col in config.address_columns]
    rest_renamed = [rename_map[col] for col in config.rest_columns]

    via_col = rename_map[config.via_column]
    civico_col = rename_map[config.civico_column]

    return PreparedTable(
        path=path,
        label=label,
        dataframe=renamed,
        comune_columns=comune_renamed,
        address_columns=address_renamed,
        via_column=via_col,
        civico_column=civico_col,
        rest_columns=rest_renamed,
    )


def _load_comune_resolver(extra_path: str | None) -> ComuneResolver | None:
    candidate_paths: list[str] = []
    if os.path.isfile(DEFAULT_COMUNE_MAP):
        candidate_paths.append(DEFAULT_COMUNE_MAP)
    if extra_path:
        absolute = os.path.abspath(extra_path)
        if not os.path.isfile(absolute):
            raise SystemExit(
                f"Errore: il file CSV specificato per i comuni non esiste: {extra_path}"
            )
        candidate_paths.append(absolute)

    mapping: dict[str, str] = {}
    for path in candidate_paths:
        mapping.update(_read_comune_mapping(path))

    if not mapping:
        return None

    return ComuneResolver(mapping)


def _read_comune_mapping(path: str) -> dict[str, str]:
    try:
        dataframe = pd.read_csv(path, sep=None, engine="python", dtype=str)
    except Exception as exc:  # pragma: no cover - dipende dai file utente
        raise SystemExit(
            f"Errore durante la lettura del file CSV delle equivalenze comuni '{path}': {exc}"
        ) from exc

    if dataframe.empty:
        return {}

    columns = {col.casefold(): col for col in dataframe.columns}
    alias_column = columns.get("alias")
    canonical_column = columns.get("canonical")
    if not alias_column or not canonical_column:
        raise SystemExit(
            f"Errore: il file '{os.path.basename(path)}' deve contenere le colonne 'alias' e 'canonical'."
        )

    mapping: dict[str, str] = {}
    for alias_value, canonical_value in dataframe[[alias_column, canonical_column]].itertuples(
        index=False, name=None
    ):
        alias_text = _safe_string(alias_value)
        canonical_text = _safe_string(canonical_value)
        alias_norm = _normalize_text(alias_text)
        canonical_norm = _normalize_text(canonical_text)
        if not alias_norm or not canonical_norm:
            continue
        mapping[alias_norm] = canonical_norm
        mapping.setdefault(canonical_norm, canonical_norm)

    return mapping


def _safe_string(value) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value)
    text = text.strip()
    text = text.strip(EXTRA_STRIP_CHARS)
    return text


def _normalize_text(value: str) -> str:
    return NORMALIZE_RE.sub("", value.upper())


def _strip_street_prefix(value: str) -> str:
    text = _safe_string(value)
    if not text:
        return ""
    cleaned = re.sub(r"[\.,]", " ", text.upper())
    tokens = [token for token in cleaned.split() if token]
    while tokens:
        head = tokens[0].replace(".", "").replace("'", "")
        if head in STREET_PREFIXES:
            tokens = tokens[1:]
            continue
        break
    return " ".join(tokens)


def _compare_values(
    value_a,
    value_b,
    *,
    resolver: ComuneResolver | None = None,
    preprocess: Callable[[str], str] | None = None,
    similarity_threshold: float = 0.88,
) -> str:
    text_a = _safe_string(value_a)
    text_b = _safe_string(value_b)

    if preprocess:
        text_a = preprocess(text_a)
        text_b = preprocess(text_b)

    if text_a == "" and text_b == "":
        return "vuoti"
    if text_a == "":
        return "solo_file_2"
    if text_b == "":
        return "solo_file_1"

    if text_a.casefold() == text_b.casefold():
        return "uguale"

    norm_a = resolver.canonical(text_a) if resolver else _normalize_text(text_a)
    norm_b = resolver.canonical(text_b) if resolver else _normalize_text(text_b)
    if norm_a and norm_a == norm_b:
        return "uguale"

    if text_a and text_b:
        ratio = SequenceMatcher(None, text_a.casefold(), text_b.casefold()).ratio()
        if ratio >= similarity_threshold:
            return "simile"
    return "diverso"


def _compose_rest_value(row: pd.Series, columns: Sequence[str]) -> str:
    values = [_safe_string(row.get(col, "")) for col in columns]
    clean_values = [value for value in values if value]
    return " | ".join(clean_values)


def _first_non_empty(row: pd.Series, columns: Sequence[str]) -> str:
    for column in columns:
        value = _safe_string(row.get(column, ""))
        if value:
            return value
    return ""


def _overall_flag(flags: Sequence[str]) -> str:
    if all(flag in {"uguale", "vuoti"} for flag in flags):
        return "coincidenza"
    if any(flag == "diverso" for flag in flags):
        return "differenza"
    if any(flag == "simile" for flag in flags):
        return "simile"
    if any(flag.startswith("solo_file") for flag in flags):
        return "incompleto"
    return "parziale"


def _build_flags(
    merged: pd.DataFrame,
    left: PreparedTable,
    right: PreparedTable,
    comune_resolver: ComuneResolver | None,
) -> pd.DataFrame:
    def compute_flags(row: pd.Series) -> pd.Series:
        comune_left_value = _first_non_empty(row, left.comune_columns)
        comune_right_value = _first_non_empty(row, right.comune_columns)
        comune_flag = _compare_values(
            comune_left_value,
            comune_right_value,
            resolver=comune_resolver,
        )
        via_flag = _compare_values(
            row[left.via_column],
            row[right.via_column],
            preprocess=_strip_street_prefix,
        )
        civico_flag = _compare_values(row[left.civico_column], row[right.civico_column])

        rest_left = _compose_rest_value(row, left.rest_columns)
        rest_right = _compose_rest_value(row, right.rest_columns)
        rest_flag = _compare_values(rest_left, rest_right)

        overall = _overall_flag((comune_flag, via_flag, civico_flag, rest_flag))
        return pd.Series(
            [comune_flag, via_flag, civico_flag, rest_flag, overall],
            index=FLAG_COLUMNS,
        )

    return merged.apply(compute_flags, axis=1)


def _ensure_output_path(path: str | None, file_a: PreparedTable, file_b: PreparedTable, mode: str) -> str:
    if path:
        output_path = os.path.abspath(path)
        directory = os.path.dirname(output_path)
        if directory and not os.path.isdir(directory):
            os.makedirs(directory, exist_ok=True)
        return output_path

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    filename = f"confronto_{file_a.label}_{file_b.label}_{mode}.xlsx"
    return os.path.join(OUTPUT_DIR, filename)


def main() -> None:
    args = parse_args()
    mode = args.mode
    config = MODE_CONFIGS[mode]
    comune_resolver = _load_comune_resolver(args.comune_map)

    selected_files = _resolve_files(args)
    left_table = _prepare_table(selected_files[0], config)
    right_table = _prepare_table(selected_files[1], config)

    merged = left_table.dataframe.merge(right_table.dataframe, on="match", how="inner")
    if merged.empty:
        raise SystemExit(
            "Nessuna riga con valori di 'match' presenti in entrambi i file. "
            "Nulla da confrontare."
        )

    flag_df = _build_flags(merged, left_table, right_table, comune_resolver)

    ordered_columns = (
        ["match"]
        + list(left_table.comune_columns)
        + list(right_table.comune_columns)
        + list(left_table.address_columns)
        + list(right_table.address_columns)
    )

    output_df = merged.loc[:, ordered_columns].copy()
    for column in FLAG_COLUMNS:
        output_df[column] = flag_df[column]

    output_path = _ensure_output_path(args.output, left_table, right_table, mode)
    try:
        output_df.to_excel(output_path, index=False)
    except Exception as exc:  # pragma: no cover - dipende dai file utente
        raise SystemExit(f"Errore durante il salvataggio del file di output: {exc}") from exc

    print(f"Modalità di confronto: {mode}")
    print(
        f"File confrontati: {os.path.basename(left_table.path)} ({left_table.label}) "
        f"vs {os.path.basename(right_table.path)} ({right_table.label})"
    )
    print(f"Righe confrontate: {len(output_df)}")
    print(f"File generato: {output_path}")


if __name__ == "__main__":
    main()
