import tkinter as tk
from tkinter import filedialog
import pandas as pd
import datetime

# Función que se ejecuta al hacer clic en el botón "Seleccionar archivo"
def select_file():
    file_path = filedialog.askopenfilename()
    file_path_var.set(file_path)

# Función que se ejecuta al hacer clic en el botón "Validar"
def validate():
    # Lee el archivo seleccionado
    file_path = file_path_var.get()
    df = pd.read_excel(file_path)

    # Realiza la validación
    column1 = column1_var.get()
    column2 = column2_var.get()
    df['Validación'] = ['MATCH' if df[column1][i] == df[column2][i] else 'NO MATCH' for i in range(len(df))]

    # Crea un objeto ExcelWriter para guardar el resultado
    date_str = datetime.datetime.now().strftime('%Y-%m-%d')
    writer = pd.ExcelWriter('VALIDATION_COMPLETE'+date_str+'.xlsx', engine='openpyxl')

    # Copia el archivo original al nuevo archivo
    df.to_excel(writer, index=False)

    # Crea un objeto de formato de celda para colorear las celdas de la columna "Validación"
    red_fill = openpyxl.styles.PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    green_fill = openpyxl.styles.PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')

    # Selecciona la hoja de cálculo del nuevo archivo y aplica el formato de celda
    worksheet = writer.sheets['Sheet1']
    for i in range(len(df)):
        if df['Validación'][i] == 'MATCH':
            worksheet.cell(row=i+1, column=len(df.columns)+1).value = df['Validación'][i]
            worksheet.cell(row=i+1, column=len(df.columns)+1).fill = green_fill
        else:
            worksheet.cell(row=i+1, column=len(df.columns)+1).value = df['Validación'][i]
            worksheet.cell(row=i+1, column=len(df.columns)+1).fill = red_fill

    # Guarda el archivo
    writer.save()

# Crea la ventana principal de la GUI
root = tk.Tk()
root.title("Validación de archivo de Excel")

# Crea los widgets de la GUI
select_file_button = tk.Button(root, text="Seleccionar archivo", command=select_file)
file_path_var = tk.StringVar()
file_path_entry = tk.Entry(root, textvariable=file_path_var)
column1_label = tk.Label(root, text="Nombre de la columna 1:")
column1_var = tk.StringVar()
column1_entry = tk.Entry(root, textvariable=column1_var)
column2_label = tk.Label(root, text="Nombre de la columna 2:")
column2_var = tk.StringVar()
column2_entry = tk.Entry(root, textvariable=column2_var)
validate_button = tk
