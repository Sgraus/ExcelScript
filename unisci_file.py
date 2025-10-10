import os
import pandas as pd

# Imposta la cartella contenente i file Excel
cartella = os.path.join(os.path.dirname(__file__), 'input')

# Trova tutti i file Excel nella cartella
file_excel = [f for f in os.listdir(cartella) if f.endswith('.xlsx') or f.endswith('.xls')]

# Lista per i DataFrame
df_list = []

for file in file_excel:
    percorso = os.path.join(cartella, file)
    df = pd.read_excel(percorso)
    df_list.append(df)

# Unisci tutti i DataFrame
df_unito = pd.concat(df_list, ignore_index=True)

# Salva il risultato
df_unito.to_excel(os.path.join(cartella, 'unione.xlsx'), index=False)