import pandas as pd
from tkinter import messagebox
from datetime import datetime

class Construcciones:
    def __init__(self, archivo_entry):
        self.archivo_entry = archivo_entry
        
        
    def validar_construcciones_No_convencionales(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Construcciones'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_construcciones")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            
            resultados = []

            for index, row in df.iterrows():
                conv = row['ConvencionalNoConvencional']
                calificacion = row['calificacionNoConvencional']

                if conv == 'No Convencional' and pd.isna(calificacion) or calificacion=='':
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'ConvencionalNoConvencional': row['ConvencionalNoConvencional'],
                        'calificacionNoConvencional': row['calificacionNoConvencional'],
                        'Observacion': 'Calificación no convencional es nula para Noconvencional'
                    }
                    resultados.append(resultado)
                    
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
                
                
            if resultados:
                df_resultado = pd.DataFrame(resultados)

                # Guardar el resultado en un nuevo archivo Excel
                output_file = 'CONSTRUCCIONES_VALIDACION.xlsx'
                sheet_name = 'CONSTRUCCIONES_VALIDACION'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")

                messagebox.showinfo("Éxito", f"Proceso completado. {len(resultados)} registros encontrados.")
                
            else:
                messagebox.showinfo("Información", "No se encontraron registros que cumplan con la condición.")
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")