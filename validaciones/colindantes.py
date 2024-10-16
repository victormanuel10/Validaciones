import pandas as pd
from tkinter import messagebox
from datetime import datetime

class Colindantes:
    
    def __init__(self, archivo_entry):
        self.archivo_entry = archivo_entry
        
    
    def validar_orientaciones_colindantes(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Colindantes'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
            
        try:
            # Leer el archivo Excel y la hoja Colindantes
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
            
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")
            
            # Normalizar los valores de la columna 'Orientacion' para evitar problemas con mayúsculas o espacios
            df['Orientacion'] = df['Orientacion'].str.strip().str.upper()
            
            # Agrupar por NroFicha y revisar si cada uno tiene al menos las orientaciones "ESTE", "NORTE", "SUR", "OESTE"
            orientaciones_requeridas = {"ESTE", "NORTE", "SUR", "OESTE"}
            resultados = []
            fichas = df.groupby('NroFicha')
            
            for nro_ficha, grupo in fichas:
                # Obtener las orientaciones únicas en el grupo
                orientaciones_presentes = set(grupo['Orientacion'].unique())
                
                # Verificar si faltan orientaciones
                orientaciones_faltantes = orientaciones_requeridas - orientaciones_presentes
                
                if orientaciones_faltantes:
                    resultado = {
                        'NroFicha': nro_ficha,
                        'Observacion': f"Faltan orientaciones: {', '.join(orientaciones_faltantes)}",
                        'Nombre Hoja':nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Error en NroFicha {nro_ficha}: {resultado['Observacion']}")
            
            # Si se encuentran errores, se guardan en un archivo Excel
            if resultados:
                '''
                df_resultado = pd.DataFrame(resultados)
                output_file = 'ERRORES_ORIENTACIONES_COLINDANTES.xlsx'
                df_resultado.to_excel(output_file, sheet_name='ErroresOrientaciones', index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} errores.")
                '''
            else:
                messagebox.showinfo("Sin errores", "Todos los NroFicha tienen las orientaciones 'ESTE', 'NORTE', 'SUR', y 'OESTE'.")
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")