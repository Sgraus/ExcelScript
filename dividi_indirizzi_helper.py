import re
from typing import Dict, Tuple

MONTH_PATTERN = (
    r"GENNAIO|FEBBRAIO|MARZO|APRILE|MAGGIO|GIUGNO|"
    r"LUGLIO|AGOSTO|SETTEMBRE|OTTOBRE|NOVEMBRE|DICEMBRE"
)

ROAD_ACRONYM_RE = re.compile(
    r"\bS\.?\s*S\b|\bS\.?\s*P\b|\bS\.?\s*R\b|"
    r"\bSS\b|\bSP\b|\bSR\b|\bS\.S\.|\bS\.P\.|\bS\.R\.",
    re.UNICODE,
)
KM_RE = re.compile(r"\bKM\b|\bKM\.?\d", re.UNICODE)
N_MARKER_RE = re.compile(r"\bN(?:\.|[°º])?\s+\d+", re.UNICODE)
SNC_RE = re.compile(r"\bSNC\b|\bS\.?N\b", re.UNICODE)
N_EXPLICIT_RE = re.compile(
    r"\bN(?:\.|[°º])?\s+(\d+[A-Z]?)(?:\s*[-\/ ]\s*([A-Z0-9]+))?", re.UNICODE
)
SNC_OR_SN_RE = SNC_RE
GENERIC_NUMBER_RE = re.compile(
    r"\b(\d+[A-Z]?)(?:\s*[-\/ ]\s*([A-Z0-9]+))?", re.UNICODE
)
CIVICO_HEAD_RE = re.compile(r"^(\d+[A-Z]?)(.*)$", re.UNICODE)
SUFFIX_RE = re.compile(r"^\s*[-\/]\s*([A-Z0-9]+)(.*)$", re.UNICODE)
TRAILING_N_RE = re.compile(r"\b(?:N\.?|N[°º])\b\s*$", re.UNICODE)
SEPARATORS_RE = re.compile(r"[,\.;]+")
SPACES_RE = re.compile(r"\s+")

SCALA_PATTERN = re.compile(r"\b(SC(?:ALA)?\.?\s+)([^\s,;]+)\b", re.UNICODE)
SCALA_SHORT_PATTERN = re.compile(r"\bSC\b\s*([A-Z0-9\-]+)", re.UNICODE)
INTERNO_PATTERN = re.compile(r"\b(INT(?:\.|ERNO)?\s+)([^\s,;]+)\b", re.UNICODE)
PIANO_PATTERN = re.compile(r"\b(PI(?:\.|ANO)?|P\.)\s*([^\s,;]+)\b", re.UNICODE)


def _clean_spaces(text: str) -> str:
    return SPACES_RE.sub(" ", text).strip()


def dividi_indirizzo_siatel(indirizzo) -> Dict[str, str]:
    result: Dict[str, str] = {
        "indirizzo": "",
        "civico": "",
        "scala": "",
        "interno": "",
        "piano": "",
        "estensione": "",
        "esito": "separazione fallita",
    }

    if indirizzo is None:
        return result

    orig = str(indirizzo).strip()
    if orig == "":
        return result

    upper = orig.upper()
    norm = SEPARATORS_RE.sub(" ", upper)
    norm = _clean_spaces(norm)

    has_road_acr = bool(ROAD_ACRONYM_RE.search(norm))
    has_km = bool(KM_RE.search(norm))
    has_n_marker = bool(N_MARKER_RE.search(norm))
    has_snc = bool(SNC_RE.search(norm))

    if has_road_acr and has_km and not has_n_marker and not has_snc:
        result["indirizzo"] = orig
        return result

    first_match_pos = None
    first_civico = None

    match = N_EXPLICIT_RE.search(norm)
    if match:
        first_match_pos = match.start(1)
        first_civico = match.group(1).strip()

    if first_match_pos is None:
        match = SNC_OR_SN_RE.search(norm)
        if match:
            first_match_pos = match.start()
            first_civico = "SNC"

    if first_match_pos is None:
        month_match = re.search(rf"\b(\d{{1,2}})\s+({MONTH_PATTERN})\b", norm, re.UNICODE)
        if month_match:
            after_month_pos = month_match.end()
            tail = norm[after_month_pos:].strip()
            if tail:
                tail_match = GENERIC_NUMBER_RE.search(tail)
                if tail_match:
                    first_match_pos = after_month_pos + tail_match.start()
                    first_civico = tail_match.group(1).strip()

    if first_match_pos is None:
        match = GENERIC_NUMBER_RE.search(norm)
        if match:
            first_match_pos = match.start()
            first_civico = match.group(1).strip()

    if first_match_pos is None:
        result["indirizzo"] = orig
        return result

    before = norm[:first_match_pos].strip()
    after = norm[first_match_pos:].strip()
    before = TRAILING_N_RE.sub("", before).strip()

    if before == "":
        result["indirizzo"] = orig
        return result

    result["indirizzo"] = before

    civico_base = ""
    rest = ""

    if first_civico == "SNC":
        civico_base = "SNC"
        rest = re.sub(r"^SNC\b\s*", "", after).strip()
    else:
        civico_match = CIVICO_HEAD_RE.match(after)
        if civico_match:
            civico_base = civico_match.group(1).strip()
            rest = civico_match.group(2).strip()
        else:
            result["civico"] = after.strip()
            result["esito"] = "separazione parziale"
            return result

        if rest:
            suffix_match = SUFFIX_RE.match(rest)
            if suffix_match:
                result["estensione"] = suffix_match.group(1).strip()
                rest = suffix_match.group(2).strip()

    leftover = rest

    def consume(pattern: re.Pattern, source: str) -> Tuple[str, str]:
        consume_match = pattern.search(source)
        if not consume_match:
            return "", source
        value = consume_match.group(2).strip()
        new_source = pattern.sub("", source, count=1).strip()
        return value, new_source

    scala, leftover = consume(SCALA_PATTERN, leftover)
    if scala == "":
        scala_match = SCALA_SHORT_PATTERN.search(leftover)
        if scala_match:
            scala = scala_match.group(1).strip()
            leftover = SCALA_SHORT_PATTERN.sub("", leftover, count=1).strip()

    interno, leftover = consume(INTERNO_PATTERN, leftover)
    piano, leftover = consume(PIANO_PATTERN, leftover)

    leftover = _clean_spaces(SEPARATORS_RE.sub(" ", leftover))
    if leftover != "":
        result["estensione"] = (
            f"{result['estensione']} {leftover}".strip()
            if result["estensione"]
            else leftover
        )

    result["civico"] = civico_base
    if scala:
        result["scala"] = scala
    if interno:
        result["interno"] = interno
    if piano:
        result["piano"] = piano

    clean_parsed = result["indirizzo"] != "" and result["civico"] != ""
    if clean_parsed:
        has_unstructured = result["estensione"] != ""
        if first_civico == "SNC":
            result["esito"] = "separazione parziale"
        elif has_unstructured:
            result["esito"] = "separazione parziale"
        else:
            result["esito"] = "correttamente separato"
    else:
        if result["indirizzo"] == "":
            result["indirizzo"] = orig
        result["esito"] = "separazione fallita"

    return result


def dividi_indirizzo_siatel_compatto(indirizzo) -> Dict[str, str]:
    base = dividi_indirizzo_siatel(indirizzo)

    pieces = []
    scala = base.get("scala", "").strip()
    if scala:
        pieces.append(f"SC {scala}")

    interno = base.get("interno", "").strip()
    if interno:
        pieces.append(f"INT {interno}")

    piano = base.get("piano", "").strip()
    if piano:
        pieces.append(f"PI {piano}")

    estensione = base.get("estensione", "").strip()
    if estensione:
        pieces.append(estensione)

    specifica = " ".join(pieces).strip()

    return {
        "indirizzo_diviso": str(base.get("indirizzo", "")).strip(),
        "civico_diviso": str(base.get("civico", "")).strip(),
        "specifica_civico_diviso": specifica,
        "esito": str(base.get("esito", "separazione fallita")).strip()
        or "separazione fallita",
    }


__all__ = ["dividi_indirizzo_siatel", "dividi_indirizzo_siatel_compatto"]
