import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook

# Creamos la ventana de selección de archivo
root = tk.Tk()
root.withdraw()

# Abrimos el diálogo para seleccionar el archivo
file_path = filedialog.askopenfilename(title="Seleccione un archivo de Excel",
                                       filetypes=[("Archivos de Excel", "*.xlsx")])

# Cargamos el libro de trabajo de Excel
workbook = load_workbook(filename=file_path)

# Obtenemos el nombre del archivo sin la extensión
file_name = file_path.split("/")[-1].split(".")[0]

# Creamos un nuevo libro de trabajo para almacenar los datos limpios
clean_workbook = load_workbook(filename=file_path)
clean_workbook.remove(clean_workbook.active)  # eliminamos la hoja de inicio predeterminada
clean_sheet = clean_workbook.create_sheet(title=file_name)

# Unimos todas las hojas en una sola hoja en el libro de trabajo limpio
for sheet_name in workbook.sheetnames:
    sheet = workbook[sheet_name]
    for row in sheet.iter_rows(min_row=1, values_only=True):
        clean_sheet.append(row)

# Guardamos el libro de trabajo limpio con un nuevo nombre
clean_file_name = f"{file_name}_Clean_Data_C5_2023.xlsx"
clean_workbook.save(filename=clean_file_name)

print(f"El archivo {clean_file_name} ha sido guardado correctamente.")
