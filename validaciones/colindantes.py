import pandas as pd
from tkinter import messagebox
from datetime import datetime

class Colindantes:
    
    def __init__(self, archivo_entry):
        self.archivo_entry = archivo_entry
    def validar_orientaciones_colindantes(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Colindantes'
        hoja_fichas = 'Fichas'  # Hoja donde se encuentra la columna 'Npn'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
            
        try:
            # Leer el archivo Excel y las hojas necesarias
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
            df_fichas = pd.read_excel(archivo_excel, sheet_name=hoja_fichas)

            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Normalizar los valores de la columna 'Orientacion'
            df['Orientacion'] = df['Orientacion'].str.strip().str.upper()

            # Validar la existencia de las columnas necesarias
            if 'NroFicha' not in df.columns or 'NroFicha' not in df_fichas.columns:
                messagebox.showerror("Error", "La columna 'NroFicha' no existe en las hojas necesarias.")
                return
            
            df = pd.merge(df, df_fichas[['NroFicha', 'Npn']], on='NroFicha', how='left')

            orientacion_map = {
                "SURESTE": {"SUR", "ESTE"},
                "NORESTE": {"NORTE", "ESTE"},
                "SUROESTE": {"SUR", "OESTE"},
                "NOROESTE": {"NORTE", "OESTE"}
            }

            orientaciones_requeridas = {"ESTE", "NORTE", "SUR", "OESTE"}
            resultados = []
            fichas = df.groupby('NroFicha')
            
            for nro_ficha, grupo in fichas:
                # Obtener todas las orientaciones únicas en el grupo
                orientaciones_presentes = set()
                for orientacion in grupo['Orientacion'].unique():
                    if orientacion in orientacion_map:
                        # Añadir componentes de orientaciones compuestas
                        orientaciones_presentes.update(orientacion_map[orientacion])
                    else:
                        # Añadir orientación simple
                        orientaciones_presentes.add(orientacion)
                
                # Verificar si faltan orientaciones
                orientaciones_faltantes = orientaciones_requeridas - orientaciones_presentes
                
                if orientaciones_faltantes:
                    radicados = ', '.join(grupo['Radicado'].dropna().astype(str).unique())
                    resultado = {
                        'NroFicha': nro_ficha,
                        'Observacion': f"Faltan orientaciones: {', '.join(orientaciones_faltantes)}",
                        'Radicado': radicados,
                        'Nombre Hoja': nombre_hoja,
                        'Npn': grupo['Npn'].iloc[0]  # Agregar el valor de 'Npn'
                    }
                    resultados.append(resultado)
                    print(f"Error en NroFicha {nro_ficha}: {resultado['Observacion']}")

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
            ][['NroFicha', 'Npn']]

            if fichas_validas.empty:
                print("No se encontraron fichas válidas en la hoja Fichas.")
                messagebox.showinfo("Sin datos", "No se encontraron fichas válidas para validar.")
                return []

            # Filtrar las NroFicha de Colindantes
            df_colindantes_filtradas = pd.merge(
                df_colindantes,
                fichas_validas,
                on='NroFicha',
                how='inner'
            )

            # Orientaciones requeridas
            orientaciones_requeridas = {"ESTE", "NORTE", "SUR", "OESTE", "ZENIT", "NADIR"}

            # Mapa de orientaciones compuestas
            orientacion_map = {
                "SURESTE": {"SUR", "ESTE"},
                "NORESTE": {"NORTE", "ESTE"},
                "SUROESTE": {"SUR", "OESTE"},
                "NOROESTE": {"NORTE", "OESTE"}
            }

            resultados = []

            # Agrupar por NroFicha y verificar orientaciones
            fichas = df_colindantes_filtradas.groupby('NroFicha')
            for nro_ficha, grupo in fichas:
                orientaciones_presentes = set()
                for orientacion in grupo['Orientacion'].unique():
                    if orientacion in orientacion_map:
                        orientaciones_presentes.update(orientacion_map[orientacion])
                    else:
                        orientaciones_presentes.add(orientacion)

                orientaciones_faltantes = orientaciones_requeridas - orientaciones_presentes

                if orientaciones_faltantes:
                    radicados = ', '.join(grupo['Radicado'].dropna().astype(str).unique())
                    npn = grupo['Npn'].iloc[0]
                    resultado = {
                        'NroFicha': nro_ficha,
                        'Npn': npn,
                        'Observacion': f"Faltan orientaciones: {', '.join(orientaciones_faltantes)} en Rph",
                        'Radicado': radicados,
                        'Nombre Hoja': hoja_colindantes
                    }
                    resultados.append(resultado)
                    print(f"Error en NroFicha {nro_ficha}: {resultado['Observacion']}")

            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []