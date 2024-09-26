import pandas as pd
from tkinter import messagebox
from datetime import datetime

class ZonasHomogeneas:
    def __init__(self, archivo_entry):
        self.archivo_entry = archivo_entry
        
    def validar_tipo_zonas_homogeneas(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'ZonasHomogeneas'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel y la hoja ZonasHomogeneas
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
            
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")
            
            resultados = []
            fichas = df.groupby('NroFicha')
            
            for nro_ficha, grupo in fichas:
                tiene_fisica = 'Fisica' in grupo['Tipo'].values
                tiene_geoeconomica = 'Geoeconomica' in grupo['Tipo'].values
                
                # Si falta alguno de los tipos, se agrega a los resultados
                if not (tiene_fisica and tiene_geoeconomica):
                    observacion = []
                    if not tiene_fisica:
                        observacion.append("Falta tipo 'Fisica'")
                    if not tiene_geoeconomica:
                        observacion.append("Falta tipo 'Geoeconomica'")
                    
                    resultado = {
                        'NroFicha': nro_ficha,
                        'Observacion': ', '.join(observacion),
                        'Nombre Hoja':nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Error en NroFicha {nro_ficha}: {resultado['Observacion']}")
            
            # Si se encuentran errores, se guardan en un archivo Excel
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'ERRORES_ZONAS_HOMOGENEAS.xlsx'
                df_resultado.to_excel(output_file, sheet_name='ErroresZonasHomogeneas', index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} errores.")
            
            else:
                messagebox.showinfo("Sin errores", "Todos los NroFicha tienen registros de 'fisica' y 'geoeconomica'.")
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")