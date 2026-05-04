import argparse
import os
import re
from datetime import datetime

import pandas as pd

from confronta_indirizzi import (
    _compare_street_values,
    _compare_values_with_score,
    _load_comune_resolver,
    _safe_string,
    _weighted_address_score,
)
from dividi_indirizzi_helper import dividi_indirizzo_siatel

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
DEFAULT_SIATEL_PATH = os.path.join(OUTPUT_DIR, "result_siatel.xlsx")
DEFAULT_OUTPUT_NAME = "result_erede_prevalente.xlsx"
DEFAULT_PREVALENCE_RULES = {
    "indirizzo_simile_stesso_comune": 100,
    "stesso_comune": 70,
    "piu_giovane": 50,
    "sesso_femminile": 20,
}
IRREPERIBLE_VALUES = {"INDIRIZZO ASSENTE", "IRREPERIBILE"}
CF_RE = re.compile(r"[A-Z0-9]{16}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Incrocia il file idrico con il risultato SIATEL e determina l'erede prevalente "
            "usando il confronto comune/indirizzo del repository corrente."
        )
    )
    parser.add_argument(
        "--idrico",
        required=True,
        help="Percorso del file Excel idrico.",
    )
    parser.add_argument(
        "--siatel",
        default=DEFAULT_SIATEL_PATH,
        help="Percorso del file Excel SIATEL (default: output/result_siatel.xlsx).",
    )
    parser.add_argument(
        "--output-name",
        default=DEFAULT_OUTPUT_NAME,
        help=f"Nome del file di output (default: {DEFAULT_OUTPUT_NAME}).",
    )
    parser.add_argument(
        "--output-dir",
        default=OUTPUT_DIR,
        help="Cartella di destinazione del file generato.",
    )
    parser.add_argument(
        "--comune-map",
        help=(
            "CSV opzionale con colonne 'alias' e 'canonical' per gestire comuni equivalenti. "
            "Usa lo stesso formato degli altri strumenti del repository."
        ),
    )
    return parser.parse_args()


def extract_codici_fiscali(text) -> list[str]:
    return CF_RE.findall(str(text).upper())


def calcola_eta(data_nascita_str) -> int | None:
    value = _safe_string(data_nascita_str)
    if not value:
        return None
    try:
        nascita = datetime.strptime(value, "%d/%m/%Y")
    except ValueError:
        return None
    oggi = datetime.today()
    return oggi.year - nascita.year - ((oggi.month, oggi.day) < (nascita.month, nascita.day))


def ensure_output_name(output_name: str) -> str:
    normalized = output_name.strip() or DEFAULT_OUTPUT_NAME
    if not normalized.lower().endswith(".xlsx"):
        normalized = f"{normalized}.xlsx"
    return normalized


def build_defunto_id(row: pd.Series) -> str:
    if "ID_DEFUNTO" in row.index and _safe_string(row.get("ID_DEFUNTO", "")):
        return _safe_string(row.get("ID_DEFUNTO", ""))
    if "ID" in row.index and _safe_string(row.get("ID", "")):
        return _safe_string(row.get("ID", ""))
    fields = [
        "NOME_DEFUNTO",
        "COGNOME_DEFUNTO",
        "Indirizzo_idrico",
        "Civico_idrico",
        "Comune_idrico",
        "DataNascitaDefunto",
        "DataDecessoDefunto",
    ]
    return "|".join(_safe_string(row.get(column, "")) for column in fields)


def split_address(address: str) -> dict[str, str]:
    return dividi_indirizzo_siatel(address)


def compare_addresses(
    comune_defunto: str,
    indirizzo_defunto: str,
    civico_defunto: str,
    comune_erede: str,
    indirizzo_erede: str,
    resolver,
) -> dict[str, object]:
    address_empty = not _safe_string(indirizzo_erede)
    if address_empty or _safe_string(indirizzo_erede).upper() in IRREPERIBLE_VALUES:
        return {
            "casistica_comune": "EREDE IRREPERIBILE",
            "casistica_indirizzo": "EREDE IRREPERIBILE",
            "flag_comune": "assente",
            "flag_via": "assente",
            "flag_civico": "assente",
            "flag_dettagli": "assente",
            "flag_generale": "irreperibile",
            "score_indirizzo": 0,
            "indirizzo_erede_diviso": "",
            "civico_erede_diviso": "",
            "scala_erede_divisa": "",
            "interno_erede_diviso": "",
            "piano_erede_diviso": "",
            "estensione_erede_divisa": "",
            "esito_split_erede": "",
        }

    splitted = split_address(indirizzo_erede)
    comune_flag, comune_score = _compare_values_with_score(
        comune_defunto,
        comune_erede,
        resolver=resolver,
    )
    via_flag, via_score = _compare_street_values(
        indirizzo_defunto,
        splitted.get("indirizzo", ""),
    )
    civico_flag, civico_score = _compare_values_with_score(
        civico_defunto,
        splitted.get("civico", ""),
    )
    dettagli_erede = " | ".join(
        value
        for value in [
            _safe_string(splitted.get("scala", "")),
            _safe_string(splitted.get("interno", "")),
            _safe_string(splitted.get("piano", "")),
            _safe_string(splitted.get("estensione", "")),
        ]
        if value
    )
    dettagli_flag, dettagli_score = _compare_values_with_score("", dettagli_erede)
    score_indirizzo = _weighted_address_score(
        comune_score,
        via_score,
        civico_score,
        dettagli_score,
    )

    stesso_comune = comune_flag == "uguale"
    indirizzo_simile = via_flag in {"uguale", "uguale_ordinato", "simile"} and score_indirizzo >= 80
    flag_generale = "coincidenza" if stesso_comune and indirizzo_simile else "differenza"

    return {
        "casistica_comune": "STESSO COMUNE" if stesso_comune else "DIFFERENTE COMUNE",
        "casistica_indirizzo": "INDIRIZZO SIMILE" if indirizzo_simile else "INDIRIZZO DIVERSO",
        "flag_comune": comune_flag,
        "flag_via": via_flag,
        "flag_civico": civico_flag,
        "flag_dettagli": dettagli_flag,
        "flag_generale": flag_generale,
        "score_indirizzo": score_indirizzo,
        "indirizzo_erede_diviso": _safe_string(splitted.get("indirizzo", "")),
        "civico_erede_diviso": _safe_string(splitted.get("civico", "")),
        "scala_erede_divisa": _safe_string(splitted.get("scala", "")),
        "interno_erede_diviso": _safe_string(splitted.get("interno", "")),
        "piano_erede_diviso": _safe_string(splitted.get("piano", "")),
        "estensione_erede_divisa": _safe_string(splitted.get("estensione", "")),
        "esito_split_erede": _safe_string(splitted.get("esito", "")),
    }


def calcola_punteggio(row: pd.Series, rules: dict[str, int]) -> int:
    punteggio = 0
    if (
        row.get("casistica_comune") == "STESSO COMUNE"
        and row.get("casistica_indirizzo") == "INDIRIZZO SIMILE"
    ):
        punteggio += rules.get("indirizzo_simile_stesso_comune", 0)
    elif row.get("casistica_comune") == "STESSO COMUNE":
        punteggio += rules.get("stesso_comune", 0)

    eta_erede = row.get("eta_erede")
    if pd.notna(eta_erede):
        punteggio += rules.get("piu_giovane", 0) - int(eta_erede) // 10

    if _safe_string(row.get("SessoTrib", "")).upper() == "F":
        punteggio += rules.get("sesso_femminile", 0)

    return int(punteggio)


def scegli_erede_prevalente(group: pd.DataFrame) -> pd.DataFrame:
    result = group.copy()
    result["erede_prevalente"] = ""
    validi = result[
        result["stato"].eq("trovato")
        & result["casistica_indirizzo"].ne("EREDE IRREPERIBILE")
        & result["DataDecesso"].fillna("").astype(str).str.strip().eq("")
    ]
    if validi.empty:
        result["erede_prevalente"] = "TUTTI EREDI DECEDUTI"
        return result

    max_score = validi["punteggio_prevalenza"].max()
    candidati = validi[validi["punteggio_prevalenza"] == max_score].sort_index()
    result.loc[candidati.index[0], "erede_prevalente"] = "EREDE PREVALENTE"
    return result


def read_siatel_dataframe(path: str) -> pd.DataFrame:
    try:
        workbook = pd.ExcelFile(path)
    except Exception as exc:
        raise SystemExit(f"Errore durante la lettura del file SIATEL '{path}': {exc}") from exc

    sheet_name = "elaborato" if "elaborato" in workbook.sheet_names else workbook.sheet_names[0]
    try:
        dataframe = pd.read_excel(path, sheet_name=sheet_name)
    except Exception as exc:
        raise SystemExit(f"Errore durante la lettura del foglio SIATEL '{sheet_name}': {exc}") from exc

    if dataframe.empty:
        raise SystemExit("Errore: il file SIATEL non contiene righe utili.")
    if "Cf" not in dataframe.columns:
        raise SystemExit("Errore: il file SIATEL non contiene la colonna 'Cf'.")
    dataframe = dataframe.copy()
    dataframe["Cf"] = dataframe["Cf"].astype(str).str.strip().str.upper()
    return dataframe


def main() -> None:
    args = parse_args()
    idrico_path = os.path.abspath(args.idrico)
    siatel_path = os.path.abspath(args.siatel)

    if not os.path.isfile(idrico_path):
        raise SystemExit(f"Errore: il file idrico '{args.idrico}' non esiste.")
    if not os.path.isfile(siatel_path):
        raise SystemExit(f"Errore: il file SIATEL '{args.siatel}' non esiste.")

    try:
        forniture_df = pd.read_excel(idrico_path)
    except Exception as exc:
        raise SystemExit(f"Errore durante la lettura del file idrico '{args.idrico}': {exc}") from exc
    if forniture_df.empty:
        raise SystemExit("Errore: il file idrico non contiene righe utili.")
    if "CODICE_FISCALE_EREDE" not in forniture_df.columns:
        raise SystemExit("Errore: il file idrico non contiene la colonna 'CODICE_FISCALE_EREDE'.")

    siatel_df = read_siatel_dataframe(siatel_path)
    comune_resolver = _load_comune_resolver(args.comune_map)

    risultati: list[dict] = []
    for _, row in forniture_df.iterrows():
        codici_fiscali = extract_codici_fiscali(row.get("CODICE_FISCALE_EREDE", ""))
        numero_eredi = len(codici_fiscali)
        trovato = False

        comune_defunto = _safe_string(row.get("Comune_idrico", ""))
        indirizzo_defunto = _safe_string(row.get("Indirizzo_idrico", ""))
        civico_defunto = _safe_string(row.get("Civico_idrico", ""))
        id_defunto = build_defunto_id(row)

        for codice in codici_fiscali:
            match_df = siatel_df[siatel_df["Cf"] == codice]
            if match_df.empty:
                continue

            trovato = True
            for _, match_row in match_df.iterrows():
                risultato = {**row.to_dict(), **match_row.to_dict()}
                risultato["ID_DEFUNTO"] = id_defunto
                risultato["cf_erede_estratto"] = codice
                risultato["numero_eredi_rilevati"] = numero_eredi
                risultato["casistica_unico_erede"] = (
                    "UNICO EREDE" if numero_eredi == 1 else "PIU' EREDI"
                )
                risultato["eta_erede"] = calcola_eta(match_row.get("DataNascitaTrib"))

                confronto = compare_addresses(
                    comune_defunto,
                    indirizzo_defunto,
                    civico_defunto,
                    _safe_string(match_row.get("Comune", "")),
                    _safe_string(match_row.get("Indirizzo", "")),
                    comune_resolver,
                )
                risultato.update(confronto)
                risultato["stato"] = "trovato"
                risultati.append(risultato)

        if not trovato:
            risultato = row.to_dict()
            risultato["ID_DEFUNTO"] = id_defunto
            risultato["cf_erede_estratto"] = ""
            risultato["numero_eredi_rilevati"] = numero_eredi
            risultato["stato"] = "non trovato"
            risultato["casistica_unico_erede"] = "NON TROVATO"
            risultato["casistica_comune"] = "NON TROVATO"
            risultato["casistica_indirizzo"] = "NON TROVATO"
            risultato["eta_erede"] = None
            risultato["flag_comune"] = "non_trovato"
            risultato["flag_via"] = "non_trovato"
            risultato["flag_civico"] = "non_trovato"
            risultato["flag_dettagli"] = "non_trovato"
            risultato["flag_generale"] = "non_trovato"
            risultato["score_indirizzo"] = 0
            risultato["indirizzo_erede_diviso"] = ""
            risultato["civico_erede_diviso"] = ""
            risultato["scala_erede_divisa"] = ""
            risultato["interno_erede_diviso"] = ""
            risultato["piano_erede_diviso"] = ""
            risultato["estensione_erede_divisa"] = ""
            risultato["esito_split_erede"] = ""
            risultati.append(risultato)

    risultati_df = pd.DataFrame(risultati)
    if risultati_df.empty:
        raise SystemExit("Errore: nessun incrocio utile prodotto.")

    risultati_df["punteggio_prevalenza"] = risultati_df.apply(
        calcola_punteggio, axis=1, rules=DEFAULT_PREVALENCE_RULES
    )

    grouped_results = []
    for _, group in risultati_df.groupby("ID_DEFUNTO", sort=False):
        grouped_results.append(scegli_erede_prevalente(group))
    risultati_df = pd.concat(grouped_results, ignore_index=True)

    output_dir = os.path.abspath(args.output_dir)
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, ensure_output_name(args.output_name))
    try:
        risultati_df.to_excel(output_path, index=False)
    except Exception as exc:
        raise SystemExit(f"Errore durante il salvataggio del file di output: {exc}") from exc

    print(f"File idrico: {idrico_path}")
    print(f"File SIATEL: {siatel_path}")
    print(f"Righe prodotte: {len(risultati_df)}")
    print(f"File generato: {output_path}")


if __name__ == "__main__":
    main()
