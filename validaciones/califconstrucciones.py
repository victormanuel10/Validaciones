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
                
                
                
                
                
                #messagebox.showinfo( f"Tamaño baño. {len(resultados)} registros.")
                
            else:
                messagebox.showinfo("Información", "No se encontraron registros que cumplan con la condición.")
            return resultados      
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
    
    
    def Validar_armazon(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'CalificacionesConstrucciones'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: Validar_armazon")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            resultados = []

            for index, row in df.iterrows():
                Armazon = row['Armazon']
                Muro = row['Muro']

                # Lista de muros válidos para la combinación con Armazon
                muros_validos = ['122|BAHAREQUE,ADOBE, TAPIA', '121|MATERIALES DE DESECHOS,ESTERILLA', '123|MADERA']

                # Condición corregida para validar el Armazón y los Muros
                if Armazon == '111|MADERA, TAPIA' and Muro not in muros_validos:
                    resultado = {
                        'Secuencia': row['Secuencia'],
                        'Armazon': row['Armazon'],
                        'Muro': row['Muro'],
                        'Observacion': 'Muro invalido para armazon',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

            if resultados:
                df_resultado = pd.DataFrame(resultados)
                '''
                
                
                output_file = 'MuroInvalido.xlsx'
                sheet_name = 'MuroInvalido'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")

                messagebox.showinfo("Éxito", f"Proceso completado. Muro Invalido con {len(resultados)} registros.")
                '''
            else:
                messagebox.showinfo("Información", "No se encontraron registros que cumplan con la condición.")

            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    
    def Validar_Cubierta(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'CalificacionesConstrucciones'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: Validar_armazon")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            resultados = []

            for index, row in df.iterrows():
                Cubierta = row['Cubierta']
                Muro = row['Muro']

                # Lista de muros válidos para la combinación con Armazon
                muros_validos = ['125|BLOQUE,LADRILLO,MADERA FINA']

                # Condición corregida para validar el Armazón y los Muros
                if Cubierta == '133|ENTREPISO' and Muro not in muros_validos:
                    resultado = {
                        'Secuencia': row['Secuencia'],
                        'Cubierta': row['ArmCubiertaazon'],
                        'Muro': row['Muro'],
                        'Observacion': 'Muro invalido para armazon',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

            if resultados:
                df_resultado = pd.DataFrame(resultados)
                '''
                output_file = 'MuroInvalido.xlsx'
                sheet_name = 'MuroInvalido'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")

                messagebox.showinfo("Éxito", f"Proceso completado. Muro Invalido con {len(resultados)} registros.")
                
                '''
                
            else:
                messagebox.showinfo("Información", "No se encontraron registros que cumplan con la condición.")

            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    def Conservacion_banio_bueno(self):
        
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'CalificacionesConstrucciones'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: Validar_armazon")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            resultados = []

            for index, row in df.iterrows():
                ConservacionBanio = row['ConservacionBanio']
                EnchapesBanio = row['EnchapesBanio']

                # Lista de muros válidos para la combinación con Armazon
                
                # Condición corregida para validar el Armazón y los Muros
                if  ConservacionBanio == '343|BUENO' and EnchapesBanio=='322|PAÑETE, BALDOSA COMÚN DE CEMENTO':
                    resultado = {
                        'Secuencia': row['Secuencia'],
                        'ConservacionBanio':row['ConservacionBanio'],
                        'EnchapesBanio': row['EnchapesBanio'],
                        'Observacion': 'La conservacion del baño es incorrecta para enchape',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

            if resultados:
                df_resultado = pd.DataFrame(resultados)
                '''
                output_file = 'BanioBueno.xlsx'
                sheet_name = 'BanioBueno'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")

                messagebox.showinfo("Éxito", f"Proceso completado. Muro Invalido con {len(resultados)} registros.")
                
                '''
                
                
            else:
                messagebox.showinfo("Información", "No se encontraron registros que cumplan con la condición.")

            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    
    def conservacion_cubierta_bueno(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja_calificaciones = 'CalificacionesConstrucciones'
        nombre_hoja_construcciones = 'Construcciones'

        if not archivo_excel or not nombre_hoja_calificaciones or not nombre_hoja_construcciones:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica los nombres de las hojas.")
            return

        try:
            
            df_calificaciones = pd.read_excel(archivo_excel, sheet_name=nombre_hoja_calificaciones)
            df_construcciones = pd.read_excel(archivo_excel, sheet_name=nombre_hoja_construcciones)

            print(f"Leyendo archivo: {archivo_excel}")
            print(f"Dimensiones del DataFrame de Calificaciones: {df_calificaciones.shape}")
            print(f"Dimensiones del DataFrame de Construcciones: {df_construcciones.shape}")

            resultados = []

            for index, row in df_calificaciones.iterrows():
                conservacion = row['Conservacion']
                cubierta = row['Cubierta']

                if conservacion == '143|BUENO' and cubierta == '132|ZINC,TEJA DE BARRO':
                    secuencia = row['Secuencia']

                    
                    construccion_row = df_construcciones[df_construcciones['secuencia'] == secuencia]

                    if not construccion_row.empty and construccion_row.iloc[0]['EdadConstruccion'] >= 20:
                        resultado = {
                            'Secuencia': secuencia,
                            'Conservacion': conservacion,
                            'Cubierta': cubierta,
                            'EdadConstruccion': construccion_row.iloc[0]['EdadConstruccion'],
                            'Observacion': 'La edad de la construcción es mayor o igual a 20 años',
                            'Nombre Hoja': nombre_hoja_calificaciones
                        }
                        resultados.append(resultado)
                        print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

            if resultados:
                '''
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_Construcciones.xlsx'
                sheet_name = 'Errores'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Proceso completado. Se encontraron {len(resultados)} registros con errores.")
            
                '''
            else:
                messagebox.showinfo("Información", "No se encontraron registros que cumplan con las condiciones.")

            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
