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
                        'Radicado':grupo['Radicado'],
                        'Nombre Hoja':nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Error en NroFicha {nro_ficha}: {resultado['Observacion']}")
            '''
            # Si se encuentran errores, se guardan en un archivo Excel
            if resultados:
                
                df_resultado = pd.DataFrame(resultados)
                output_file = 'ERRORES_ORIENTACIONES_COLINDANTES.xlsx'
                df_resultado.to_excel(output_file, sheet_name='ErroresOrientaciones', index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} errores.")
                
            else:
                messagebox.showinfo("Sin errores", "Todos los NroFicha tienen las orientaciones 'ESTE', 'NORTE', 'SUR', y 'OESTE'.")
            '''    
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
    def validar_orientaciones_rph(self):
        archivo_excel = self.archivo_entry.get()
        hoja_colindantes = 'Colindantes'
        hoja_fichas = 'Fichas'

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return

        try:
            # Leer las hojas del archivo Excel
            df_colindantes = pd.read_excel(archivo_excel, sheet_name=hoja_colindantes)
            df_fichas = pd.read_excel(archivo_excel, sheet_name=hoja_fichas)

            print(f"Leyendo archivo: {archivo_excel}")
            print(f"Dimensiones de Colindantes: {df_colindantes.shape}, Fichas: {df_fichas.shape}")

            # Normalizar las columnas necesarias
            df_colindantes['Orientacion'] = df_colindantes['Orientacion'].fillna('').str.strip().str.upper()
            df_colindantes['NroFicha'] = df_colindantes['NroFicha'].fillna('').astype(str).str.strip()
            df_fichas['NroFicha'] = df_fichas['NroFicha'].fillna('').astype(str).str.strip()
            df_fichas['Npn'] = df_fichas['Npn'].fillna('').astype(str).str.strip()

            # Filtro en la hoja Fichas
            df_fichas['Ultimos_4'] = df_fichas['Npn'].str[-4:].apply(lambda x: sum(int(d) for d in x if d.isdigit()))
            fichas_validas = df_fichas[
                (df_fichas['Npn'].str[21:22] == '9') & 
                (df_fichas['Ultimos_4'] != 0)
            ]['NroFicha'].unique()

            print(f"NroFicha válidas desde Fichas: {fichas_validas}")

            # Verifica si hay fichas válidas
            if len(fichas_validas) == 0:
                print("No se encontraron fichas válidas en la hoja Fichas.")
                messagebox.showinfo("Sin datos", "No se encontraron fichas válidas para validar.")
                return []

            # Filtrar las NroFicha de Colindantes
            df_colindantes_filtradas = df_colindantes[df_colindantes['NroFicha'].isin(fichas_validas)]
            print(f"Dimensiones de Colindantes filtradas: {df_colindantes_filtradas.shape}")

            # Orientaciones requeridas
            orientaciones_requeridas = {"ESTE", "NORTE", "SUR", "OESTE", "ZENIT", "NADIR"}
            resultados = []

            # Agrupar por NroFicha y verificar orientaciones
            fichas = df_colindantes_filtradas.groupby('NroFicha')
            for nro_ficha, grupo in fichas:
                orientaciones_presentes = set(grupo['Orientacion'].unique())
                orientaciones_faltantes = orientaciones_requeridas - orientaciones_presentes

                if orientaciones_faltantes:
                    resultado = {
                        'NroFicha': nro_ficha,
                        'Observacion': f"Faltan orientaciones: {', '.join(orientaciones_faltantes)} en Rph",
                        'Radicado':grupo['Radicado'],
                        'Nombre Hoja': hoja_colindantes
                    }
                    resultados.append(resultado)
                    print(f"Error en NroFicha {nro_ficha}: {resultado['Observacion']}")
            '''
            
            # Guardar resultados si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'ERRORES_ORIENTACIONES_COLINDANTES.xlsx'
                df_resultado.to_excel(output_file, sheet_name='ErroresOrientaciones', index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Errores encontrados", f"Se encontraron errores. Archivo guardado en '{output_file}'.")
            else:
                messagebox.showinfo("Sin errores", "Todas las NroFicha cumplen con las orientaciones requeridas.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []