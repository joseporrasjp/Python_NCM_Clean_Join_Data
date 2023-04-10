import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Crear una ventana para seleccionar el archivo
root = tk.Tk()
root.withdraw()

# Seleccionar el archivo de Excel
file_path = filedialog.askopenfilename(title="Selecciona un archivo de Excel", filetypes=[("Excel files", "*.xlsx")])

# Cargar el archivo de Excel
df_dict = pd.read_excel(file_path, sheet_name=None)

# Unir todas las hojas en un solo DataFrame
df_merged = pd.concat(df_dict.values(), ignore_index=True)

# Dejar solo el encabezado principal de la primera hoja
main_header = df_dict[list(df_dict.keys())[0]].iloc[0]
df_merged.columns = main_header

# Quitar los encabezados de las hojas secundarias
df_merged = df_merged.iloc[1:]

# Guardar el archivo resultante con el nombre original más "Clean_Data_C5_2023"
clean_file_path = file_path[:-5] + "_Clean_Data_C5_2023.xlsx"
df_merged.to_excel(clean_file_path, index=False)

print(f"El archivo {clean_file_path} se ha guardado automáticamente.")
