import pandas as pd
from tkinter import messagebox
class FiltroFichas:

    def __init__(self, archivo_entry):
        self.archivo_entry = archivo_entry
        self.df_filtrado = None
        self.df_filtrado_rph_parcelacion = None

    def filtrar_datos_fichas(self):
        archivo_excel = self.archivo_entry.get()

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return None

        try:
            # Leer la hoja Fichas
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            # Filtrar según los valores requeridos en la columna CaracteristicaPredio
            valores_validos = ['1|NPH (0)', '12|INFORMAL (2)', '13|BIEN DE USO PUBLICO (3)']
            df_fichas['CaracteristicaPredio'] = df_fichas['CaracteristicaPredio'].astype(str).str.strip()

            # Aplicar el filtro
            self.df_filtrado = df_fichas[df_fichas['CaracteristicaPredio'].isin(valores_validos)]

            if self.df_filtrado.empty:
                messagebox.showinfo("Información", "No se encontraron registros con los valores específicos de CaracteristicaPredio.")
                return None

            return self.df_filtrado

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return None

    def filtrar_datos_rph_parcelacion(self):
        archivo_excel = self.archivo_entry.get()

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return None

        try:
            # Leer la hoja Fichas
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            # Filtrar según los valores '2|RPH' y '3|Parcelacion'
            valores_rph_parcelacion = ['2|RPH', '3|Parcelacion']
            df_fichas['CaracteristicaPredio'] = df_fichas['CaracteristicaPredio'].astype(str).str.strip()

            # Aplicar el filtro
            self.df_filtrado_rph_parcelacion = df_fichas[df_fichas['CaracteristicaPredio'].isin(valores_rph_parcelacion)]

            if self.df_filtrado_rph_parcelacion.empty:
                messagebox.showinfo("Información", "No se encontraron registros con CaracteristicaPredio igual a '2|RPH' o '3|Parcelacion'.")
                return None

            return self.df_filtrado_rph_parcelacion

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return None

    def obtener_fichas_filtradas(self):
        if self.df_filtrado is not None:
            return self.df_filtrado
        else:
            return self.filtrar_datos_fichas()

    def obtener_fichas_rph_parcelacion(self):
        if self.df_filtrado_rph_parcelacion is not None:
            return self.df_filtrado_rph_parcelacion
        else:
            return self.filtrar_datos_rph_parcelacion()
