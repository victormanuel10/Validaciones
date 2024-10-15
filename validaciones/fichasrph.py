import pandas as pd
from tkinter import messagebox
from NPHORPH.fichasvalidador import FiltroFichas

class FichasRPH:
    def __init__(self, archivo_entry):
        self.archivo_entry = archivo_entry
        self.resultados_generales = []
        self.filtro_fichas=FiltroFichas(archivo_entry)

    def validar_coeficiente_copropiedad(self):
        archivo_excel = self.archivo_entry.get()
        
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return []
        df_fichas_filtradas = self.filtro_fichas.obtener_fichas_rph_parcelacion()
        if df_fichas_filtradas is None:
            return []
        try:
            # Leer la hoja FICHAS (o la hoja donde esté NumCedulaCatastral y CoeficienteCopropiedad)
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
