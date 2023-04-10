import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Seleccionar archivo de Excel con GUI
Tk().withdraw() # ocultar ventana principal de tkinter
file_path = askopenfilename()

# Leer archivo Excel
dfs = pd.read_excel(file_path, sheet_name=None)

# Unir todas las hojas del DataFrame en un solo DataFrame
merged_df = pd.concat(dfs.values())

# Eliminar encabezados duplicados en la hoja principal
merged_df = merged_df.loc[:, ~merged_df.columns.duplicated()]


# Obtener el nombre del archivo original sin la extensión
output_file_name = file_path.split("/")[-1].split(".")[0]

# Guardar el archivo con el nombre original y agregar "_Clean_Data_C5_2023.xlsx"
output_file_name += "_Clean_Data_C5_2023.xlsx"
merged_df.to_excel(output_file_name, index=False)

print(f"Archivo guardado como {output_file_name}")
