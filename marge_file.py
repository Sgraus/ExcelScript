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


def main():
    excel_files = find_excel_files(cartella)

    if len(excel_files) > 2:
        raise SystemExit(f"Errore: trovati {len(excel_files)} file Excel nella cartella 'input'. Devono essere al massimo 2.")
    if len(excel_files) < 2:
        raise SystemExit("Errore: per eseguire il merge servono esattamente 2 file Excel nella cartella 'input'.")

    # Leggi entrambi i file e valida la colonna 'match'
    df_a = read_excel_or_fail(excel_files[0])
    df_b = read_excel_or_fail(excel_files[1])

    # Scegli come base il file con meno righe
    if len(df_a) <= len(df_b):
        base_df, other_df = df_a, df_b
        base_name = os.path.splitext(os.path.basename(excel_files[0]))[0]
        other_name = os.path.splitext(os.path.basename(excel_files[1]))[0]
    else:
        base_df, other_df = df_b, df_a
        base_name = os.path.splitext(os.path.basename(excel_files[1]))[0]
        other_name = os.path.splitext(os.path.basename(excel_files[0]))[0]

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

    print(f"Merge completato. File base: {base_name} (righe: {len(base_df)}). Altro file: {other_name} (righe: {len(other_df)}).")
    print(f"Totale righe in output: {len(merged)}")
    print(f"File salvato in: {out_path}")


if __name__ == '__main__':
    main()

