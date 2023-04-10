import pandas as pd

# Cargar el archivo de Excel
file_path = "path/to/excel/file.xlsx"
df = pd.read_excel(file_path, sheet_name=None)

# Unir todas las hojas en un solo DataFrame
df_merged = pd.concat(df.values(), ignore_index=True)

# Quitar los encabezados de las hojas secundarias
df_merged = df_merged[df_merged.columns[~df_merged.columns.duplicated()]]

# Guardar el archivo resultante con el nombre original más "Clean_Data_C5_2023"
clean_file_path = file_path[:-5] + "_Clean_Data_C5_2023.xlsx"
df_merged.to_excel(clean_file_path, index=False)

print(f"El archivo {clean_file_path} se ha guardado automáticamente.")