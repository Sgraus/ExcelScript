import os
import pandas as pd

# Define paths
base_dir = os.path.dirname(os.path.abspath(__file__))
input_dir = os.path.join(base_dir, "input")
output_dir = os.path.join(base_dir, "output")

# Create output directory if it doesn't exist
os.makedirs(output_dir, exist_ok=True)

# Find first supported file in input folder
supported_exts = {".xlsx", ".xlsm", ".xls", ".csv"}
input_files = [f for f in os.listdir(input_dir) if os.path.splitext(f)[1].lower() in supported_exts]
if not input_files:
    raise FileNotFoundError("Nessun file supportato trovato nella cartella 'input'. Formati accettati: .xlsx, .xlsm, .xls, .csv")
input_file = os.path.join(input_dir, input_files[0])

# Load file (assume first row contains the header)
ext = os.path.splitext(input_file)[1].lower()
if ext == ".csv":
    # Auto-detect delimiter (comma/semicolon/tab) using Python engine
    df = pd.read_csv(input_file, sep=None, engine="python")
else:
    # Excel family (.xlsx/.xlsm/.xls)
    # pandas will choose the right engine (openpyxl/xlrd) if installed
    df = pd.read_excel(input_file, header=0)

if df.shape[0] == 0:
    raise ValueError("Il file è vuoto.")

# Keep columns as read; do NOT drop the first data row.
columns = list(df.columns)
data = df.reset_index(drop=True)

# Split into chunks of 100 rows
chunk_size = 100
file_name = os.path.basename(input_file)

if len(data) == 0:
    # No data rows, but still create an empty file with just the header
    output_file = os.path.join(output_dir, f"1 {file_name}")
    empty_df = pd.DataFrame(columns=columns)
    empty_df.to_excel(output_file, index=False)
    print(f"Creato (vuoto, solo intestazione): {output_file}")
else:
    for i in range(0, len(data), chunk_size):
        df_chunk = data.iloc[i:i + chunk_size].copy()
        df_chunk.columns = columns  # ensure header names persist
        part_num = i // chunk_size + 1
        output_file = os.path.join(output_dir, f"{part_num} {file_name}")
        df_chunk.to_excel(output_file, index=False)
        print(f"Creato: {output_file}")