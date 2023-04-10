import pandas as pd
import tkinter as tk
from tkinter import filedialog

def main():
    # Seleccionar archivo de Excel
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    
    # Leer todas las hojas del archivo
    sheets = pd.read_excel(file_path, sheet_name=None)
    
    # Unir todas las hojas en un solo DataFrame
    df = pd.concat(sheets, ignore_index=True)
    
    # Eliminar el índice que se crea por defecto
    df.reset_index(drop=True, inplace=True)
    
    # Asignar nuevos nombres a las columnas
    new_column_names = ['CODIGO_ZONA', 'CAMPANIA', 'PAIS', 'FECHA_FACTURACION', 'CODIGO_CONSEJERA',
                        'NIVEL_ESTRELLA', 'LINE_NUMBER', 'DESCRIPCION', 'PRECIO', 'CANTIDAD',
                        'VALOR_TOTAL', 'COMISION', 'VALOR_NETO', 'TIPO_REGISTO']
    df.columns = new_column_names
    
    # Validar que existan las columnas requeridas
    if not all(column_name in df.columns for column_name in new_column_names):
        missing_columns = set(new_column_names) - set(df.columns)
        df = df.dropna(subset=missing_columns, how='all')
    
    # Eliminar filas duplicadas de encabezado
    df = df.loc[df.ne(df.columns).any(axis=1)]
    
    # Guardar el archivo limpio con un nuevo nombre
    output_file_path = file_path.replace('.xlsx', '_Clean_Data_C5_2023.xlsx')
    df.to_excel(output_file_path, index=False)
    
    print(f"Archivo limpio guardado en: {output_file_path}")
    
if __name__ == '__main__':
    main()
