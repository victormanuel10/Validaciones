import pandas as pd
from tkinter import messagebox
from datetime import datetime

class CalificaionesConstrucciones:
    def __init__(self, archivo_entry):
        self.archivo_entry = archivo_entry
        
    def validar_banios(self):
        
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'CalificacionesConstrucciones' 
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return 
        
        try:
            
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: Validar_baños")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            
            resultados = []

            for index, row in df.iterrows():
                Tamaniobanio = row['TamanioBanio']
                EnchapesBanio = row['EnchapesBanio']
                MobiliarioBanio = row['MobiliarioBanio']
                ConservacionBanio = row['ConservacionBanio']
                
                if Tamaniobanio == '311|SIN BAÑO' and (pd.notna(EnchapesBanio) or pd.notna(MobiliarioBanio) or pd.notna(ConservacionBanio)):
                    resultado = {
                        'Secuencia':row['Secuencia'],
                        'Tamaño baño': row['TamanioBanio'],
                        'Observacion': 'No puede tener EnchapesBanio, MobiliarioBanio, ConservacionBanio ',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
                
                
            if resultados:
                df_resultado = pd.DataFrame(resultados)

                '''
                output_file = 'Validar_Baños.xlsx'
                sheet_name = 'Validar Baños'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")

                '''
                
                
                
                
                
                messagebox.showinfo( f"Tamaño baño. {len(resultados)} registros.")
                
            else:
                messagebox.showinfo("Información", "No se encontraron registros que cumplan con la condición.")
            return resultados      
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
    
        
    