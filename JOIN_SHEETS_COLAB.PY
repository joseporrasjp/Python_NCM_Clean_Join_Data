import pandas as pd
from google.colab import files

# Seleccionar archivo de Excel
uploaded = files.upload()

# Leer archivo Excel
for file_name in uploaded.keys():
    # Leer todas las hojas del archivo Excel en un diccionario de DataFrames
    dfs = pd.read_excel(file_name, sheet_name=None)

# Unir todas las hojas del DataFrame en un solo DataFrame
merged_df = pd.concat(dfs.values())

# Eliminar encabezados duplicados en la hoja principal
merged_df = merged_df.loc[:, ~merged_df.columns.duplicated()]


# Obtener el nombre del archivo original sin la extensión
output_file_name = file_name.split(".")[0]

# Guardar el archivo con el nombre original y agregar "_Clean_Data_C5_2023.xlsx"
output_file_name += "_Clean_Data_C5_2023.xlsx"
merged_df.to_excel(output_file_name, index=False)

# Descargar el archivo generado
files.download(output_file_name)
