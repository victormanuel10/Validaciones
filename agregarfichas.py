'''

# -*- coding: utf-8 -*-
import arcpy
import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog

class AgregarFichas:
    def __init__(self, parent):
        self.parent = parent

        # Variables para almacenar rutas
        self.gdb_path = tk.StringVar()
        self.excel_file_path = tk.StringVar()

        # Crear widgets para la pestaña
        self.crear_widgets()

    def crear_widgets(self):
        # Botón para seleccionar la GDB
        tk.Button(self.parent, text="Seleccionar GDB", command=self.select_gdb).grid(row=0, column=0, pady=10)

        # Mostrar la ruta de la GDB seleccionada
        tk.Label(self.parent, text="GDB seleccionada:").grid(row=0, column=1, pady=10)
        tk.Label(self.parent, textvariable=self.gdb_path).grid(row=0, column=2, pady=10)

        # Botón para seleccionar y cargar el archivo de Excel
        tk.Button(self.parent, text="Seleccionar Excel", command=self.select_excel).grid(row=1, column=0, pady=10)

        # Mostrar la ruta del archivo Excel seleccionado
        tk.Label(self.parent, text="Archivo Excel seleccionado:").grid(row=1, column=1, pady=10)
        tk.Label(self.parent, textvariable=self.excel_file_path).grid(row=1, column=2, pady=10)

        # Botón para ejecutar todo el proceso
        tk.Button(self.parent, text="Ejecutar", command=self.process_all).grid(row=2, column=0, pady=10)

    def select_gdb(self):
        """Abrir un cuadro de diálogo para seleccionar una geodatabase."""
        gdb = filedialog.askdirectory(title="Seleccionar GDB")
        if gdb and os.path.isdir(gdb) and gdb.endswith(".gdb"):
            self.gdb_path.set(gdb)
        else:
            messagebox.showerror("Error", "Seleccione una geodatabase válida")

    def select_excel(self):
        """Seleccionar el archivo Excel para su posterior procesamiento."""
        excel_file = filedialog.askopenfilename(title="Seleccionar archivo Excel",
                                                  filetypes=[("Archivos Excel", "*.xlsx")])
        if excel_file:
            self.excel_file_path.set(excel_file)
        else:
            messagebox.showerror("Error", "Seleccione un archivo Excel válido.")

    def process_excel(self):
        """Procesar el archivo Excel y realizar la actualización en el feature class."""
        try:
            gdb = self.gdb_path.get()
            excel_file = self.excel_file_path.get()

            if not gdb:
                messagebox.showerror("Error", "Seleccione una geodatabase primero.")
                return

            if not excel_file:
                messagebox.showerror("Error", "Seleccione un archivo Excel primero.")
                return

            # Leer el Excel, específicamente el libro "Fichas"
            df = pd.read_excel(excel_file, sheet_name='Fichas')

            # Verificar que el archivo contiene las columnas necesarias
            if 'Npn' not in df.columns or 'NroFicha' not in df.columns:
                messagebox.showerror("Error", "El archivo Excel no contiene las columnas 'Npn' y 'NroFicha'.")
                return

            # Definir la ruta del feature class
            fc_path = os.path.join(gdb, "r_lc_terreno")
            if not arcpy.Exists(fc_path):
                messagebox.showerror("Error", "El feature class 'r_lc_terreno' no existe en la geodatabase.")
                return

            # Crear un diccionario del Excel con Npn como clave y Nficha como valor
            npn_to_ficha = dict(zip(df['Npn'].astype(str), df['NroFicha'].astype(str)))

            # Iniciar una sesión de edición
            with arcpy.da.Editor(gdb) as edit_session:
                # Iterar sobre las filas del feature class r_lc_terreno
                with arcpy.da.UpdateCursor(fc_path, ['terreno_codigo', 'NroFicha']) as cursor:
                    for row in cursor:
                        terreno_codigo = str(row[0])
                        if terreno_codigo in npn_to_ficha:
                            row[1] = npn_to_ficha[terreno_codigo]
                            cursor.updateRow(row)

            messagebox.showinfo("Éxito", "Datos del archivo Excel importados y actualizados correctamente.")

        except Exception as e:
            messagebox.showerror("Error", "Error al procesar el archivo Excel: " + str(e))

    def add_fields(self):
        """Agregar los campos a las feature classes especificadas."""
        try:
            gdb = self.gdb_path.get()
            if not gdb:
                messagebox.showerror("Error", "Seleccione una geodatabase primero.")
                return

            feature_classes = [
                "u_lc_terreno", "r_lc_terreno", "u_lc_construccion", "r_lc_construccion", "u_lc_unidadconstruccion",
                "r_lc_unidadconstruccion"
            ]

            common_fields = [
                ('NroFicha', 'Text', 50),
            ]

            additional_fields = [
                ('Npnresumen', 'Text', 30)
            ]

            arcpy.env.workspace = gdb

            for fc in feature_classes:
                fc_path = os.path.join(gdb, fc)

                if not arcpy.Exists(fc_path):
                    messagebox.showerror("Error", "La feature class " + fc + " no existe en la geodatabase.")
                    continue

                for field_name, field_type, field_length in common_fields:
                    existing_fields = [f.name for f in arcpy.ListFields(fc_path)]
                    if field_name not in existing_fields:
                        arcpy.AddField_management(fc_path, field_name, field_type, field_length=field_length)

                if fc in ["u_lc_unidadconstruccion", "r_lc_unidadconstruccion"]:
                    for field_name, field_type, field_length in additional_fields:
                        existing_fields = [f.name for f in arcpy.ListFields(fc_path)]
                        if field_name not in existing_fields:
                            arcpy.AddField_management(fc_path, field_name, field_type, field_length=field_length)

                    arcpy.CalculateField_management(
                        in_table=fc_path,
                        field="Npnresumen",
                        expression="!codigo_unidad_construccion![:24] + '00' + !codigo_unidad_construccion![26:]",
                        expression_type="PYTHON"
                    )

            messagebox.showinfo("Éxito", "Campos agregados correctamente a las feature classes.")

        except Exception as e:
            messagebox.showerror("Error", "Error al agregar campos: " + str(e))

    def process_all(self):
        """Función que ejecuta la selección de GDB, agrega campos e importa Excel en orden."""
        if not self.gdb_path.get():
            messagebox.showerror("Error", "Debe seleccionar una geodatabase antes de continuar.")
            return

        if not self.excel_file_path.get():
            messagebox.showerror("Error", "Debe seleccionar un archivo Excel antes de continuar.")
            return

        self.add_fields()
        self.process_excel()

'''
