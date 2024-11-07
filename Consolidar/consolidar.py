import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import openpyxl
import os
import numpy as np
from pathlib import Path
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl import load_workbook

class ExcelConsolidator(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Consolidador de archivos Excel")
        self.geometry("600x400")
        
        # Crear las pestañas
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill='both', expand=True)
        
        # Pestaña para consolidar carpeta
        self.tab_consolidar_carpeta = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_consolidar_carpeta, text="Consolidar carpeta")
        self.create_consolidar_carpeta_tab()
        
        # Pestaña para consolidar PH y NPH
        self.tab_consolidar_ph_nph = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_consolidar_ph_nph, text="Consolidar PH y NPH")
        self.create_consolidar_ph_nph_tab()
    
    def create_consolidar_carpeta_tab(self):
        # Variables
        self.folder_path = tk.StringVar()
        
        # Instrucciones y entrada de carpeta
        instructions = tk.Label(self.tab_consolidar_carpeta, text="Seleccione la carpeta con los archivos Excel a consolidar")
        instructions.pack(pady=10)
        
        folder_frame = tk.Frame(self.tab_consolidar_carpeta)
        folder_frame.pack(pady=10)
        
        self.folder_entry = tk.Entry(folder_frame, textvariable=self.folder_path, width=50)
        self.folder_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        browse_btn = tk.Button(folder_frame, text="Buscar", command=self.browse_folder)
        browse_btn.pack(side=tk.LEFT)
        
        # Botón para consolidar
        consolidate_btn = tk.Button(self.tab_consolidar_carpeta, text="Consolidar carpeta", command=self.consolidate_files)
        consolidate_btn.pack(pady=20)
    
    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        self.folder_path.set(folder_selected)

    def format_number(self, value):
        """
        Convierte números grandes a texto para evitar notación científica
        """
        if isinstance(value, (int, float, np.integer, np.floating)):
            if abs(value) >= 1e10:  # Para números mayores o iguales a 10 billones
                return str(int(value))  # Convertir a texto sin decimales
        return value

    def read_excel_preserve_numbers(self, file_path, sheet_name):
        """
        Lee un archivo Excel preservando los números largos como texto
        """
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheet = wb[sheet_name]
        
        data = []
        headers = []
        
        # Leer encabezados
        for cell in sheet[1]:
            headers.append(cell.value)
            
        # Leer datos
        for row in sheet.iter_rows(min_row=2):
            row_data = []
            for cell in row:
                value = cell.value
                # Formatear el valor si es necesario
                formatted_value = self.format_number(value)
                row_data.append(formatted_value)
            data.append(row_data)
            
        # Crear DataFrame
        df = pd.DataFrame(data, columns=headers)
        return df

    def make_unique_columns(self, df):
        """
        Asegura que los nombres de las columnas sean únicos.
        Si hay columnas duplicadas, les añade un sufijo incremental.
        """
        cols = pd.Series(df.columns)
        for dup in cols[cols.duplicated()].unique():
            cols[cols == dup] = [dup + f"_{i}" if i != 0 else dup for i in range(sum(cols == dup))]
        df.columns = cols
        return df
    
    def consolidate_files(self):
        if not self.folder_path.get():
            messagebox.showerror("Error", "Por favor seleccione una carpeta")
            return

        try:
            folder = Path(self.folder_path.get())
            excel_files = list(folder.glob("*.xlsx"))

            if not excel_files:
                messagebox.showerror("Error", "No se encontraron archivos Excel en la carpeta seleccionada")
                return

            first_file = openpyxl.load_workbook(excel_files[0], data_only=True)
            sheet_names = first_file.sheetnames

            consolidated_wb = openpyxl.Workbook()
            consolidated_wb.remove(consolidated_wb.active)

            for sheet_name in sheet_names:
                # Excluir la hoja "Listas" de la consolidación
                if sheet_name == "Listas":
                    # Copiar la hoja "Listas" tal como está desde el primer archivo
                    data = self.read_excel_preserve_numbers(excel_files[0], sheet_name)
                    consolidated_sheet = consolidated_wb.create_sheet(title=sheet_name)
                    consolidated_sheet.append(data.columns.tolist())  # Agregar encabezados

                    # Agregar datos de la hoja original
                    for row in data.itertuples(index=False, name=None):
                        consolidated_sheet.append(row)
                    
                    continue  # Pasar a la siguiente hoja sin consolidar

                # Proceso normal de consolidación para las demás hojas
                consolidated_sheet = consolidated_wb.create_sheet(title=sheet_name)
                dfs = []
                data_validations = []

                for excel_file in excel_files:
                    df = self.read_excel_preserve_numbers(excel_file, sheet_name)

                    # Asegurar que los nombres de las columnas sean únicos
                    df = self.make_unique_columns(df)

                    # Resetear el índice para evitar duplicados
                    df.reset_index(drop=True, inplace=True)

                    # Eliminar filas duplicadas
                    df = df.drop_duplicates().reset_index(drop=True)

                    if 'Npn' in df.columns:
                        df['NPN_TERRENO'] = df['Npn'].astype(str).str[:21]

                    dfs.append(df)

                    # Cargar el libro fuente para copiar validaciones de datos
                    source_wb = openpyxl.load_workbook(excel_file, data_only=False)
                    source_sheet = source_wb[sheet_name]

                    # Copiar validaciones de datos, evitando duplicados
                    for dv in source_sheet.data_validations.dataValidation:
                        if dv not in data_validations:
                            data_validations.append(dv)

                # Concatenar todos los DataFrames asegurando índices únicos
                consolidated_df = pd.concat(dfs, ignore_index=True, sort=False)

                # Aplicar formato de número para evitar notación científica
                for column in consolidated_df.columns:
                    consolidated_df[column] = consolidated_df[column].apply(self.format_number)

                # Guardar en un archivo Excel temporal
                with pd.ExcelWriter("temp_consolidated.xlsx", engine='openpyxl') as writer:
                    consolidated_df.to_excel(writer, sheet_name=sheet_name, index=False)

                # Cargar desde el archivo temporal y transferir a la hoja consolidada
                temp_wb = openpyxl.load_workbook("temp_consolidated.xlsx")
                temp_sheet = temp_wb[sheet_name]

                for row in temp_sheet.iter_rows():
                    for cell in row:
                        consolidated_sheet.cell(
                            row=cell.row,
                            column=cell.column,
                            value=cell.value
                        )

                # Aplicar validaciones de datos a la hoja consolidada
                for dv in data_validations:
                    consolidated_dv = DataValidation(
                        type=dv.type,
                        formula1=dv.formula1,
                        formula2=dv.formula2,
                        allow_blank=dv.allow_blank,
                        showDropDown=dv.showDropDown,
                        showErrorMessage=dv.showErrorMessage,
                        errorTitle=dv.errorTitle,
                        error=dv.error,
                        promptTitle=dv.promptTitle,
                        prompt=dv.prompt
                    )

                    # Aquí corregimos cómo manejamos los rangos de las validaciones
                    for range in dv.ranges:
                        start_cell = range.min_col, range.min_row  # Min columna y fila
                        end_cell = range.max_col, range.max_row    # Max columna y fila

                        # Convertimos las celdas de min y max en las referencias de celda correspondientes
                        start_cell_str = openpyxl.utils.get_column_letter(start_cell[0]) + str(start_cell[1])
                        end_cell_str = openpyxl.utils.get_column_letter(end_cell[0]) + str(end_cell[1])

                        # Asegurarnos de que no generamos un rango donde la fila de inicio es mayor que la fila de fin
                        if start_cell[1] <= consolidated_df.shape[0]:
                            new_range = f"{start_cell_str}:{openpyxl.utils.get_column_letter(end_cell[0])}{consolidated_df.shape[0] + 1}"
                        else:
                            new_range = f"{start_cell_str}:{openpyxl.utils.get_column_letter(end_cell[0])}{start_cell[1]}"

                        consolidated_dv.add(new_range)  # Agregar el nuevo rango a la validación

                    # Aplicar la validación de datos al rango
                    consolidated_sheet.add_data_validation(consolidated_dv)

                # Ajustar el ancho de las columnas basado en la longitud máxima de las celdas
                for column in temp_sheet.columns:
                    max_length = max(len(str(cell.value)) if cell.value else 0 for cell in column)
                    adjusted_width = (max_length + 2)
                    consolidated_sheet.column_dimensions[column[0].column_letter].width = adjusted_width

            # Guardar el libro consolidado final
            output_path = folder / "Consolidado_con_validaciones.xlsx"
            consolidated_wb.save(output_path)

            # Eliminar el archivo temporal
            if os.path.exists("temp_consolidated.xlsx"):
                os.remove("temp_consolidated.xlsx")

            messagebox.showinfo(
                "Éxito",
                f"Archivos consolidados exitosamente con validaciones.\nArchivo guardado como: {output_path}"
            )

        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error durante la consolidación:\n{str(e)}")
            
    def create_consolidar_ph_nph_tab(self):
        # Variables
        self.ruta_archivo_1 = ""
        self.ruta_archivo_2 = ""
        self.ruta_guardado = ""
        
        # Etiquetas y cuadros de entrada
        tk.Label(self.tab_consolidar_ph_nph, text="Archivo NPH:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.entry_archivo_1 = tk.Entry(self.tab_consolidar_ph_nph, width=50)
        self.entry_archivo_1.grid(row=0, column=1, padx=10, pady=10)
        tk.Button(self.tab_consolidar_ph_nph, text="Seleccionar", command=self.seleccionar_archivo_1).grid(row=0, column=2, padx=10, pady=10)

        tk.Label(self.tab_consolidar_ph_nph, text="Archivo PH:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
        self.entry_archivo_2 = tk.Entry(self.tab_consolidar_ph_nph, width=50)
        self.entry_archivo_2.grid(row=1, column=1, padx=10, pady=10)
        tk.Button(self.tab_consolidar_ph_nph, text="Seleccionar", command=self.seleccionar_archivo_2).grid(row=1, column=2, padx=10, pady=10)

        tk.Label(self.tab_consolidar_ph_nph, text="Guardar como:").grid(row=2, column=0, padx=10, pady=10, sticky="e")
        self.entry_guardado = tk.Entry(self.tab_consolidar_ph_nph, width=50)
        self.entry_guardado.grid(row=2, column=1, padx=10, pady=10)
        tk.Button(self.tab_consolidar_ph_nph, text="Seleccionar", command=self.seleccionar_guardado).grid(row=2, column=2, padx=10, pady=10)
        
        # Botón para ejecutar la consolidación
        tk.Button(self.tab_consolidar_ph_nph, text="Consolidar archivos", command=self.consolidar_archivos).grid(row=3, column=1, padx=10, pady=20)
    
    def seleccionar_archivo_1(self):
        self.ruta_archivo_1 = filedialog.askopenfilename(title="Selecciona el primer archivo de Excel", filetypes=[("Excel files", "*.xlsx")])
        self.entry_archivo_1.delete(0, tk.END)
        self.entry_archivo_1.insert(0, self.ruta_archivo_1)
    
    def seleccionar_archivo_2(self):
        self.ruta_archivo_2 = filedialog.askopenfilename(title="Selecciona el segundo archivo de Excel", filetypes=[("Excel files", "*.xlsx")])
        self.entry_archivo_2.delete(0, tk.END)
        self.entry_archivo_2.insert(0, self.ruta_archivo_2)

    def seleccionar_guardado(self):
        self.ruta_guardado = filedialog.asksaveasfilename(defaultextension=".xlsx", title="Guardar archivo consolidado como", filetypes=[("Excel files", "*.xlsx")])
        self.entry_guardado.delete(0, tk.END)
        self.entry_guardado.insert(0, self.ruta_guardado)
        

    def combinar_hojas(self, lista_hojas, data_archivo_1, data_archivo_2):
        df_consolidado = pd.DataFrame()
        for hoja in lista_hojas:
            df_consolidado = pd.concat([df_consolidado, data_archivo_1.get(hoja, pd.DataFrame()), data_archivo_2.get(hoja, pd.DataFrame())], axis=0, ignore_index=True)
        return df_consolidado

    
    def consolidar_archivos(self):
        if not self.ruta_archivo_1 or not self.ruta_archivo_2 or not self.ruta_guardado:
            messagebox.showwarning("Selección incompleta", "Por favor selecciona los dos archivos de Excel y la ubicación para guardar el archivo consolidado.")
            return

        # Definir las correspondencias entre hojas
        correspondencias = {
            'Fichas': ['Fichas', 'Ficha', 'FichasPrediales'],  
            'Propietarios': ['Propietarios'],
            'Construcciones': ['Construcciones', 'ConstruccionesFicha'],
            'CalificacionesConstrucciones': ['CalificacionesConstrucciones', 'CalificacionesConsFicha'],
            'ConstruccionesGenerales': ['ConstruccionesGenerales', 'ConstruccionGeneralFicha'],
            'Colindantes': ['Colindantes', 'ColindantesFicha'],
            'ZonasHomogeneas': ['ZonasHomogeneas'],
            'Cartografia': ['Cartografia'],
            'InformacionGrafica': ['InformacionGrafica'],
            'Listas': ['Listas']
        }

        # Leer todas las hojas de ambos archivos
        data_archivo_1 = pd.read_excel(self.ruta_archivo_1, sheet_name=None, dtype=str)
        data_archivo_2 = pd.read_excel(self.ruta_archivo_2, sheet_name=None, dtype=str)

        # Renombrar columnas en el archivo 2 para mantener la consistencia con el archivo 1
        if "Ficha" in data_archivo_2:
            data_archivo_2["Ficha"].rename(columns={
                'NumCedCatastral': 'NumCedulaCatastral',
            }, inplace=True)

        if "FichasPrediales" in data_archivo_2:
            data_archivo_2["FichasPrediales"].rename(columns={
                'DestinoEconomico':'DestinoEcconomico'
            }, inplace=True)

        consolidado = {}

        # Consolidar las hojas correspondientes
        for hoja_destino, hojas_fuente in correspondencias.items():
            data_frames = []
            for nombre_hoja in hojas_fuente:
                if nombre_hoja in data_archivo_1:
                    data_frames.append(data_archivo_1[nombre_hoja])
                if nombre_hoja in data_archivo_2:
                    data_frames.append(data_archivo_2[nombre_hoja])

            # Concatenar todos los DataFrames de la hoja actual
            if data_frames:
                df_consolidado = pd.concat(data_frames, ignore_index=True)
                consolidado[hoja_destino] = df_consolidado

        # Agregar las hojas del archivo 2 que no tienen correspondencia en las especificaciones
        hojas_no_correspondidas = set(data_archivo_2.keys()).difference([item for sublist in correspondencias.values() for item in sublist])
        for hoja in hojas_no_correspondidas:
            consolidado[hoja] = data_archivo_2[hoja]

        # Guardar el archivo consolidado en la ubicación especificada
        ruta_consolidada = self.ruta_guardado
        with pd.ExcelWriter(ruta_consolidada, engine='openpyxl') as writer:
            for nombre_hoja, df in consolidado.items():
                df.to_excel(writer, sheet_name=nombre_hoja, index=False)

        # Cargar el archivo consolidado con openpyxl para agregar las listas desplegables
        try:
            wb_consolidado = load_workbook(ruta_consolidada)

            # Copiar validaciones de datos de los archivos originales
            for ruta_archivo in [self.ruta_archivo_1, self.ruta_archivo_2]:
                wb_original = load_workbook(ruta_archivo)
                for nombre_hoja in wb_original.sheetnames:
                    if nombre_hoja in wb_consolidado.sheetnames:
                        hoja_original = wb_original[nombre_hoja]
                        hoja_consolidada = wb_consolidado[nombre_hoja]

                        # Copiar las validaciones de datos
                        if hoja_original.data_validations:
                            for dv in hoja_original.data_validations.dataValidation:
                                new_dv = DataValidation(
                                    type=dv.type, formula1=dv.formula1, formula2=dv.formula2,
                                    showDropDown=dv.showDropDown, allowBlank=dv.allowBlank
                                )
                                new_dv.sqref = dv.sqref
                                hoja_consolidada.add_data_validation(new_dv)

            # Guardar el archivo consolidado final con validaciones
            wb_consolidado.save(ruta_consolidada)

        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar el archivo: {e}")
            return

        messagebox.showinfo("Consolidación completada", f"El archivo consolidado se ha guardado en {ruta_consolidada}")
if __name__ == "__main__":
    app = ExcelConsolidator()
    app.mainloop()
