import argparse
import os

import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

ELABORATO_HEADERS = [
    "Tipo",
    "CU",
    "Cf",
    "Cognome",
    "Nome",
    "Sesso",
    "DataNascita",
    "CodComune",
    "ComuneNascita",
    "PrNascita",
    "Esito",
    "CfTrib",
    "CognomeTrib",
    "NomeTrib",
    "SessoTrib",
    "DataNascitaTrib",
    "ComuneNascitaTrib",
    "PrNascitaTrib",
    "Comune",
    "Pr",
    "CAP",
    "Indirizzo",
    "FonteInd",
    "DataDecorrenza",
    "FonteDecesso",
    "DataDecesso",
    "PartitaIVA",
    "StatoPIVA",
    "CodAttivita",
    "TipologiaCod",
    "DataInizio",
    "DataCess",
    "ComuneSL",
    "PrSL",
    "CAPSL",
    "IndirizzoSL",
    "FonteSL",
    "DataDecorrenzaSL",
    "CFRapL",
    "CodCarica",
    "DataDecorRapL",
]

ALTRI_DATI_HEADERS = [
    "Tipo",
    "CU",
    "Cf",
    "Cognome",
    "Nome",
    "Sesso",
    "DataNascita",
    "CodComune",
    "ComuneNascita",
    "PrNascita",
    "Esito",
    "CfTrib",
    "NumeroOccorrenze",
    "PartitaIVA",
    "CodAttivita",
    "TipologiaCod",
    "StatoPIVA",
    "DataCess",
    "TipologiaCess",
    "PIVAConfluenza",
    "CodCarica",
    "DataDecorrenza",
    "DataFineCarica",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Converte un file testuale SIATEL a larghezza fissa in un file Excel."
    )
    parser.add_argument("--file", required=True, help="Percorso del file testuale SIATEL.")
    parser.add_argument(
        "--output-name",
        default="result_siatel.xlsx",
        help="Nome del file Excel di output (default: result_siatel.xlsx).",
    )
    parser.add_argument(
        "--output-dir",
        default=OUTPUT_DIR,
        help="Cartella di destinazione per il file generato (default: cartella 'output').",
    )
    return parser.parse_args()


def format_date(date_str: str) -> str:
    value = date_str.strip()
    if not value:
        return ""
    return f"{value[0:2]}/{value[2:4]}/{value[4:8]}"


def format_esito(esito: str) -> str:
    mapping = {
        "0000": "Operazione correttamente eseguita",
        "0001": "Codice fiscale restituito aggiornato",
        "0002": "Codice fiscale non validato",
        "0003": "Codice fiscale base di omocodice",
        "0005": "Su tipo Record 1: Codice fiscale di persona non fisica o partita iva",
        "0006": "Codice fiscale non utilizzabile",
        "0007": "Individuati più soggetti con i dati anagrafici comunicati",
        "0008": "Interrogazione soggetto con Partita IVA",
        "0009": "Codice fiscale formalmente errato",
        "0021": "Operazione non effettuabile",
        "0110": "Dati insufficienti o errati per l'individuazione del soggetto",
        "0111": "Soggetto non presente in Archivio Anagrafico",
        "0112": "Soggetto non presente in Archivio Anagrafico",
        "2002": "Codice fiscale errato",
        "2010": "Indicare i dati anagrafici o, in alternativa, il codice fiscale",
        "2011": "Data di nascita errata",
        "4000": "Errata indicazione del Tipo Record",
        "5100": "Codice fiscale assente o dati anagrafici incompleti",
    }
    return mapping.get(esito, "")


def format_fonte_dom(fonte: str) -> str:
    mapping = {
        "0": "Censimento da mod. AA3",
        "1": "Censimento contribuenti IVA",
        "2": "Censimento da IVA 31/38",
        "3": "Censimento mod. 750/760",
        "4": "Collegamento servizi al contribuente",
        "5": "Censimento mod. AA5",
        "6": "Collegamento IVA",
        "7": "Censimento mod. AA7",
        "8": "Censimento mod. AA8",
        "9": "Censimento mod.AA9",
        "A": "Mod. 740/74",
        "B": "Mod. 740/75",
        "C": "Mod. 101/75",
        "D": "Questura",
        "E": "Mod. AA4",
        "F": "Censimento",
        "G": "Questionari",
        "H": "Mod. 740/UNICO",
        "I": "Mod. 101",
        "J": "Mod. 730",
        "K": "Sportello Unico Immigrazione",
        "O": "Comune di residenza (SISTEMA INA-SAIA)",
        "Q": "Conferma da Self-service",
        "R": "Comune di residenza (INA-SAIA MASSIVO)",
        "S": "Anagrafe comunale (allineamento) impostata a seguito di allineamento batch",
        "T": "CCIAA",
        "U": "Telematico entrate",
        "V": "Comune di residenza impostata a seguito di comunicazioni da inasaia, ftp,Siatel,AP5",
        "W": "Comune di nascita vecchia fonte impostata in seguito a comunicazione di attribuzione effettuata dal comune di nascita",
        "X": "Domicilio fiscale da provvedimento ex art.59 DPR 600",
        "Z": "Comune di residenza impostata in attribuzione codice fiscale da ina-saia, ftp,Siatel, AP5 ",
    }
    return mapping.get(fonte, "")


def format_cod_carica(codice: str) -> str:
    mapping = {
        "0": "CARICA RAPPRESENTANTE SCONOSCIUTA",
        "1": "RAPPRESENTANTE LEGALE",
        "2": "SOCIO AMMINISTRATORE/RAPPR. DI MINORE, INTERDETTO O INABILITATO",
        "3": "CURATORE FALLIMENTARE",
        "4": "COMMISSARIO LIQUIDATORE",
        "5": "COMMISSARIO GIUDIZIARIO",
        "6": "RAPPRESENTANTE FISCALE",
        "7": "EREDE",
        "8": "LIQUIDATORE",
        "9": "BENEFICIARIO (DITTE)",
        "A": "RAPPRESENTANTE FISCALE DI SOGGETTO NON RESIDENTE (ART. 44 COMMA 3, D.L. N. 331/1993)",
        "B": "AMMINISTRATORE di CONDOMINIO",
        "C": "COMMISSARIO LIQUIDATORE di Pubblica Amministrazione",
        "D": "RAPPRESENTANTE di MINORE, INABILITATO o INTERDETTO per soggetti con natura giuridica 44 o 54",
    }
    return mapping.get(codice, "")


def format_tipo_cess(tipo: str) -> str:
    mapping = {
        "1": "CAMBIO DI DOMICILIO FISCALE",
        "2": "MODIFICAZIONE DI SOCIETA' IN DITTA O VICEVERSA",
        "3": "FUSIONE PROPRIA",
        "4": "SUCCESSIONE O DONAZIONE",
        "5": "UNIFICAZIONE",
        "6": "ESERCIZIO DI PIU' ATTIVITA'",
        "7": "FUSIONE PER INCORPORAZIONE",
        "8": "CONFERIMENTO DI AZIENDA",
        "9": "CONFERIMENTO DI ATTIVITA'",
        "A": "CONFERIMENTO (solo per società), CESSIONE E DONAZIONE DI AZIENDA",
        "B": "SUCCESSIONE EREDITARIA",
        "C": "CESSAZIONE",
        "D": "CESSAZIONE EX ART.2 NONIES L. 564/94",
        "E": "CESSAZIONE EX ART.5 D.L. 282/2002",
        "F": "ESTINZIONE PER FUSIONE",
        "G": "Cessazione da sistema centrale per sanatoria(Codice tributo 8110)",
        "H": "Cessazione da sistema centrale per decesso del titolare",
        "M": "Cessazione PER ADESIONE/REVOCA AL NUOVO REGIME DI FRANCHIGIA DITTE",
        "P": "CESSAZIONE AI SOLI FINI IVA",
        "T": "SCISSIONE TOTALE",
        "U": "CESSAZIONE D'UFFICIO",
        "X": "ATTIVITA' SOSPESA",
        "Z": "ISTITUZIONE SECONDO UFFICIO IVA",
    }
    return mapping.get(tipo, "")


def format_fonte_decesso(decesso: str) -> str:
    mapping = {
        "0": "Vivente",
        "1": "Decesso comunicato da Comune on line",
        "2": "Decesso comunicato da Comune FTP",
        "3": "Decesso acquisito da dichiarazione dei redditi",
        "4": "Decesso acquisito da atto di successione",
        "5": "Informazione fornita da INPS (pensioni cancellate)",
        "6": "Informazione fornita da Ministero Economia e Finanze (pensioni cancellate)",
        "7": "Decesso comunicato da ufficio Agenzia Entrate",
        "A": "Decesso certificato da ufficio Agenzia Entrate",
        "B": "Decesso presunto per limiti di eta",
        "C": "Decesso comunicato da servizi AIRE",
        "D": "Decesso Comunicato da INPS (DAL 2011)",
        "O": "Decesso comunicato da COMUNE (Sistema INA-SAIA)",
        "R": "Decesso comunicato da COMUNE (INA-SAIA MASSIVO)",
        "Y": "Decesso recuperato da Anagrafe Comunale",
        "Z": "Fonte sconosciuta",
    }
    return mapping.get(decesso, "")


def format_tipo_cod_attivita(tipo: str) -> str:
    mapping = {
        "0": "ante 1/1/04",
        "1": "ATECOFIN 2004",
        "2": "ATECO 2007",
    }
    return mapping.get(tipo, "")


def format_cod_attivita(cod: str) -> str:
    value = cod.strip()
    if not value or not value.isdigit():
        return value
    if len(value) == 5:
        return f"{value[0:2]}.{value[2:4]}.{value[4]}"
    if len(value) == 6:
        return f"{value[0:2]}.{value[2:4]}.{value[4:6]}"
    return value


def split_row_1(row: str) -> list[str]:
    return [
        row[0:1].strip(),
        row[1:16].strip(),
        row[16:32].strip(),
        row[32:72].strip(),
        row[72:112].strip(),
        row[112:113].strip(),
        format_date(row[113:121]),
        row[121:125].strip(),
        row[125:170].strip(),
        row[170:172].strip(),
        format_esito(row[228:238].strip()),
        row[238:254].strip(),
        row[254:294].strip(),
        row[294:334].strip(),
        row[334:335].strip(),
        format_date(row[335:343]),
        row[343:388].strip(),
        row[388:390].strip(),
        row[390:435].strip(),
        row[435:437].strip(),
        row[437:442].strip(),
        row[442:477].strip(),
        format_fonte_dom(row[477:478].strip()),
        format_date(row[478:486]),
        format_fonte_decesso(row[486:487].strip()),
        format_date(row[487:495]),
        row[495:506].strip(),
        row[506:507].strip(),
        format_cod_attivita(row[507:513]),
        format_tipo_cod_attivita(row[513:514].strip()),
        format_date(row[514:522]),
        format_date(row[522:530]),
        row[530:575].strip(),
        row[575:577].strip(),
        row[577:582].strip(),
        row[582:617].strip(),
        format_fonte_dom(row[617:618].strip()),
        format_date(row[618:626]),
        "",
        "",
        "",
    ]


def split_row_2(row: str) -> list[str]:
    return [
        row[0:1].strip(),
        row[1:16].strip(),
        row[16:32].strip(),
        row[32:72].strip(),
        row[72:112].strip(),
        row[112:113].strip(),
        format_date(row[113:121]),
        row[121:125].strip(),
        row[125:170].strip(),
        row[170:172].strip(),
        format_esito(row[228:238].strip()),
        row[238:249].strip(),
        row[249:399].strip(),
        "",
        "",
        "",
        "",
        "",
        row[399:444].strip(),
        row[444:446].strip(),
        row[446:451].strip(),
        row[451:486].strip(),
        format_fonte_dom(row[486:487].strip()),
        format_date(row[487:495]),
        format_fonte_decesso(row[495:496].strip()),
        format_date(row[496:504]),
        row[504:515].strip(),
        row[515:516].strip(),
        format_cod_attivita(row[516:522]),
        format_tipo_cod_attivita(row[522:523].strip()),
        format_date(row[523:531]),
        format_date(row[531:539]),
        row[539:584].strip(),
        row[584:586].strip(),
        row[586:591].strip(),
        row[591:626].strip(),
        format_fonte_dom(row[626:627].strip()),
        format_date(row[627:635]),
        row[635:651].strip(),
        format_cod_carica(row[651:652].strip()),
        format_date(row[652:660]),
    ]


def split_row_i(row: str) -> list[str]:
    return [
        row[0:1].strip(),
        row[1:16].strip(),
        row[16:32].strip(),
        row[32:72].strip(),
        row[72:112].strip(),
        row[112:113].strip(),
        format_date(row[113:121]),
        row[121:125].strip(),
        row[125:170].strip(),
        row[170:172].strip(),
        format_esito(row[228:238].strip()),
        row[238:254].strip(),
        row[254:256].strip(),
        row[256:267].strip(),
        format_cod_attivita(row[267:273]),
        format_tipo_cod_attivita(row[273:274].strip()),
        row[274:275].strip(),
        format_date(row[275:283]),
        format_tipo_cess(row[283:284].strip()),
        row[284:295].strip(),
        "",
        "",
        "",
    ]


def split_row_r(row: str) -> list[str]:
    return [
        row[0:1].strip(),
        row[1:16].strip(),
        row[16:32].strip(),
        row[32:72].strip(),
        row[72:112].strip(),
        row[112:113].strip(),
        format_date(row[113:121]),
        row[121:125].strip(),
        row[125:170].strip(),
        row[170:172].strip(),
        format_esito(row[228:238].strip()),
        row[238:249].strip(),
        row[249:251].strip(),
        row[251:267].strip(),
        "",
        "",
        "",
        "",
        "",
        "",
        format_cod_carica(row[267:268].strip()),
        format_date(row[268:276]),
        format_date(row[276:284]),
    ]


def split_row_s(row: str) -> list[str]:
    return [
        row[0:1].strip(),
        row[1:16].strip(),
        row[16:32].strip(),
        row[32:72].strip(),
        row[72:112].strip(),
        row[112:113].strip(),
        format_date(row[113:121]),
        row[121:125].strip(),
        row[125:170].strip(),
        row[170:172].strip(),
        format_esito(row[228:238].strip()),
        row[238:254].strip(),
        row[254:256].strip(),
        row[256:272].strip(),
        "",
        "",
        "",
        "",
        "",
        "",
        format_cod_carica(row[272:273].strip()),
        format_date(row[273:281]),
        format_date(row[281:289]),
    ]


def split_siatel_row(row: str) -> tuple[str, list[str]] | None:
    record_type = row[0:1]
    if record_type == "1":
        return ("elaborato", split_row_1(row))
    if record_type == "2":
        return ("elaborato", split_row_2(row))
    if record_type == "I":
        return ("altri_dati", split_row_i(row))
    if record_type == "R":
        return ("altri_dati", split_row_r(row))
    if record_type == "S":
        return ("altri_dati", split_row_s(row))
    return None


def read_lines(file_path: str) -> list[str]:
    encodings = ("utf-8-sig", "utf-8", "latin-1")
    last_error = None
    for encoding in encodings:
        try:
            with open(file_path, "r", encoding=encoding) as handle:
                return handle.readlines()
        except UnicodeDecodeError as exc:
            last_error = exc
    raise SystemExit(f"Errore durante la lettura del file '{file_path}': {last_error}")


def ensure_output_name(output_name: str) -> str:
    normalized = output_name.strip() or "result_siatel.xlsx"
    if not normalized.lower().endswith(".xlsx"):
        normalized = f"{normalized}.xlsx"
    return normalized


def main() -> None:
    args = parse_args()
    input_path = os.path.abspath(args.file)
    if not os.path.isfile(input_path):
        raise SystemExit(f"Errore: il file '{args.file}' non esiste.")

    lines = read_lines(input_path)
    if not lines:
        raise SystemExit("Errore: il file selezionato e vuoto.")

    elaborato_rows: list[list[str]] = []
    altri_dati_rows: list[list[str]] = []

    for line in lines:
        parsed = split_siatel_row(line.rstrip("\r\n"))
        if not parsed:
            continue
        sheet_name, row = parsed
        if sheet_name == "elaborato":
            elaborato_rows.append(row)
        else:
            altri_dati_rows.append(row)

    if not elaborato_rows and not altri_dati_rows:
        raise SystemExit("Errore: nessun record SIATEL valido trovato nel file.")

    output_dir = os.path.abspath(args.output_dir)
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, ensure_output_name(args.output_name))

    with pd.ExcelWriter(output_path) as writer:
        if elaborato_rows:
            pd.DataFrame(elaborato_rows, columns=ELABORATO_HEADERS).to_excel(
                writer, sheet_name="elaborato", index=False
            )
        if altri_dati_rows:
            pd.DataFrame(altri_dati_rows, columns=ALTRI_DATI_HEADERS).to_excel(
                writer, sheet_name="altri_dati", index=False
            )

    print(f"File generato: {output_path}")
    if elaborato_rows:
        print(f" - Sheet elaborato: {len(elaborato_rows)} righe")
    if altri_dati_rows:
        print(f" - Sheet altri_dati: {len(altri_dati_rows)} righe")


if __name__ == "__main__":
    main()
