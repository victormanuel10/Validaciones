import pandas as pd
from tkinter import messagebox

class FichasRPH:
    def __init__(self, archivo_entry):
        # archivo_entry can be either a string (file path) or tkinter.Entry
        self.archivo_entry = archivo_entry
        self.resultados_generales = []

    def obtener_archivo(self):
        """ Helper function to get the file path from either a tkinter.Entry or a string. """
        if isinstance(self.archivo_entry, str):
            return self.archivo_entry
        elif hasattr(self.archivo_entry, 'get'):
            return self.archivo_entry.get()
        else:
            return None

    def validar_coeficiente_copropiedad(self):
        # Get the file path
        archivo_excel = self.obtener_archivo()

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja FICHAS
            df_fichas = pd.read_excel(archivo_excel, sheet_name='FichasPrediales')
            
            # Crear una columna con los primeros 19 dígitos de NumCedulaCatastral
            df_fichas['CedulaCatastral'] = df_fichas['NumCedulaCatastral'].astype(str).str[:19]
            
            # Agrupar por los primeros 19 dígitos de NumCedulaCatastral y sumar CoeficienteCopropiedad
            suma_coeficientes = df_fichas.groupby('CedulaCatastral')['CoeficienteCopropiedad'].sum().reset_index()
            
            # Filtrar los casos donde la suma no sea 100
            errores = suma_coeficientes[suma_coeficientes['CoeficienteCopropiedad'] != 100]

            resultados = []

            # Crear resultados para los errores encontrados
            for index, row in errores.iterrows():
                resultado = {
                    'CedulaCatastral': row['CedulaCatastral'],
                    'Suma CoeficienteCopropiedad': row['CoeficienteCopropiedad'],
                    'Observacion': 'La suma de CoeficienteCopropiedad no es 100',
                    'Nombre Hoja': 'FichasPrediales'
                }
                resultados.append(resultado)

            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_CoeficienteCopropiedad.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros.")
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []