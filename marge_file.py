import argparse
import os
import pandas as pd

# Imposta la cartella contenente i file Excel
cartella = os.path.join(os.path.dirname(__file__), 'input')


def find_excel_files(folder: str):
    """
    Restituisce la lista dei file Excel validi (xlsx/xls) nella cartella,
    escludendo file temporanei (prefisso '~$').
    """
    allowed_ext = {'.xlsx', '.xls'}
    if not os.path.isdir(folder):
        raise SystemExit(f"Errore: la cartella di input non esiste: {folder}")
    files = [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
    excel_files = [
        os.path.join(folder, f)
        for f in files
        if os.path.splitext(f)[1].lower() in allowed_ext and not f.startswith('~$')
    ]
    return excel_files


def read_excel_or_fail(path: str) -> pd.DataFrame:
    """
    Legge un file Excel e controlla la presenza della colonna 'match'.
    """
    try:
        df = pd.read_excel(path)
    except Exception as e:
        raise SystemExit(f"Errore durante la lettura di '{os.path.basename(path)}': {e}")

    if 'match' not in df.columns:
        raise SystemExit(f"Errore: la colonna 'match' manca nel file: {os.path.basename(path)}")

    # Normalizza la colonna 'match' a stringa, rimuovendo spazi
    df['match'] = df['match'].astype(str).str.strip()
    return df


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Esegue un merge tra due file Excel presenti nella cartella 'input'."
    )
    parser.add_argument(
        '--primary',
        help=(
            "Nome del file da usare come lato completo del merge. "
            "Se non specificato viene scelto automaticamente il file con meno righe."
        )
    )
    return parser.parse_args()


def main():
    args = parse_args()
    primary_choice = args.primary.strip() if args.primary else None

    excel_files = find_excel_files(cartella)

    if len(excel_files) > 2:
        raise SystemExit(f"Errore: trovati {len(excel_files)} file Excel nella cartella 'input'. Devono essere al massimo 2.")
    if len(excel_files) < 2:
        raise SystemExit("Errore: per eseguire il merge servono esattamente 2 file Excel nella cartella 'input'.")

    # Leggi entrambi i file e valida la colonna 'match'
    data = []
    for path in excel_files:
        data.append((path, read_excel_or_fail(path)))

    def normalize(name: str) -> str:
        return os.path.normcase(name.strip())

    def matches_choice(path: str) -> bool:
        filename = os.path.basename(path)
        base_no_ext = os.path.splitext(filename)[0]
        candidate = normalize(primary_choice) if primary_choice else ""
        return candidate in {normalize(filename), normalize(base_no_ext)}

    base_entry = None
    other_entry = None

    if primary_choice:
        for entry in data:
            if matches_choice(entry[0]):
                base_entry = entry
                break
        if base_entry is None:
            raise SystemExit(
                f"Errore: il file principale '{primary_choice}' non è tra i file presenti nella cartella 'input'."
            )
        other_entry = next((e for e in data if e is not base_entry), None)
        if other_entry is None:
            raise SystemExit("Errore interno: non è stato trovato il secondo file per il merge.")
    else:
        data_sorted = sorted(data, key=lambda item: len(item[1]))
        base_entry, other_entry = data_sorted[0], data_sorted[1]

    base_path, base_df = base_entry
    other_path, other_df = other_entry
    base_name = os.path.splitext(os.path.basename(base_path))[0]
    other_name = os.path.splitext(os.path.basename(other_path))[0]

    # Merge LEFT usando come "sinistra" la tabella più piccola
    merged = base_df.merge(
        other_df,
        on='match',
        how='left',
        suffixes=(f'_{base_name}', f'_{other_name}')
    )

    # Porta 'match' come prima colonna
    cols = ['match'] + [c for c in merged.columns if c != 'match']
    merged = merged.loc[:, cols]

    # Scrivi l'output
    out_dir = os.path.join(os.path.dirname(__file__), 'output')
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f'merge_{base_name}_{other_name}.xlsx')
    try:
        merged.to_excel(out_path, index=False)
    except Exception as e:
        raise SystemExit(f"Errore durante il salvataggio dell'output: {e}")

    if primary_choice:
        print(f"File principale selezionato manualmente: {os.path.basename(base_path)}")
    print(f"Merge completato. File base: {base_name} (righe: {len(base_df)}). Altro file: {other_name} (righe: {len(other_df)}).")
    print(f"Totale righe in output: {len(merged)}")
    print(f"File salvato in: {out_path}")


if __name__ == '__main__':
    main()
