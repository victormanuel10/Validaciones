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
                        'secuencia':row['secuencia'],
                        'Tipo Contruccion':row['TipoConstruccion'],
                        'ConvencionalNoConvencional': row['ConvencionalNoConvencional'],
                        'calificacionNoConvencional': row['calificacionNoConvencional'],
                        'Observacion': 'Calificación no convencional es nula para Noconvencional',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
                
                
            if resultados:
                df_resultado = pd.DataFrame(resultados)

                '''
                output_file = 'CONSTRUCCIONES_VALIDACION.xlsx'
                sheet_name = 'CONSTRUCCIONES_VALIDACION'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")

                
                
                messagebox.showinfo("Éxito", f"No convencionales. {len(resultados)} registros.")
                '''
            else:
                messagebox.showinfo("Información", "No se encontraron registros No convencionales.")
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
    def tipo_construccion_noconvencionales(self):
        
        archivo_excel=self.archivo_entry.get()
        nombre_hoja='Construcciones'
        
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
                TipoConstruccion = row['TipoConstruccion']
                NoConvensional = row['ConvencionalNoConvencional']

                if NoConvensional == 'No Convencional' and TipoConstruccion != '' :
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'secuencia':row['secuencia'],
                        'Tipo Contruccion':row['TipoConstruccion'],
                        'ConvencionalNoConvencional': row['ConvencionalNoConvencional'],
                        'calificacionNoConvencional': row['calificacionNoConvencional'],
                        'Observacion': 'TipoConstruccion debe ser vacio si es No convencional',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
                
                
            if resultados:
                df_resultado = pd.DataFrame(resultados)

                '''
                output_file = 'TipoConstruccion.xlsx'
                sheet_name = 'TipoConstruccion'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")

                
                
                messagebox.showinfo("Éxito", f"TipoConstruccion no vacio en No convencionales. {len(resultados)} registros.")
                '''
            else:
                messagebox.showinfo("Información", "No se encontraron registros tipo construccion.")
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
    
    
    def areaconstruida_mayora1000(self):
        
        archivo_excel=self.archivo_entry.get()
        nombre_hoja='Construcciones'
        
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
                AreaConstruida = row['AreaConstruida']
                

                if AreaConstruida >= 1000 :
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'secuencia':row['secuencia'],
                        'AreaConstruida': row['AreaConstruida'],
                        'Observacion': 'El area construida es mayor a 1000 mts',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
                
                
            if resultados:
                df_resultado = pd.DataFrame(resultados)

                '''
                output_file = 'Areamayor1000.xlsx'
                sheet_name = 'Areamayor1000'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"AreaMayor a 1000. {len(resultados)} registros.")
                '''
            else:
                messagebox.showinfo("Información", "No se encontraron registros Areamayor1000.")
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
    def validar_secuencia_construcciones_vs_generales(self):
        archivo_excel = self.archivo_entry.get()

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return
        
        try:
            # Leer las hojas Construcciones y ConstruccionesGenerales
            df_construcciones = pd.read_excel(archivo_excel, sheet_name='Construcciones')
            df_generales = pd.read_excel(archivo_excel, sheet_name='ConstruccionesGenerales')

            print(f"Leyendo archivo: {archivo_excel}")
            print(f"Dimensiones de Construcciones: {df_construcciones.shape}")
            print(f"Dimensiones de ConstruccionesGenerales: {df_generales.shape}")

            # Extraer las columnas de Secuencia
            secuencia_construcciones = df_construcciones['secuencia'].dropna().unique()
            secuencia_generales = df_generales['Secuencia'].dropna().unique()

            # Encontrar secuencias en Construcciones que no están en ConstruccionesGenerales
            secuencias_faltantes = set(secuencia_construcciones) - set(secuencia_generales)

            resultados = []
            for secuencia in secuencias_faltantes:
                resultado = {
                    'secuencia': secuencia,
                    'Observacion': 'Secuencia esta en Construcciones pero no está en ConstruccionesGenerales',
                    'Nombre Hoja': 'Construcciones'
                }
                resultados.append(resultado)
                print(f"Secuencia faltante: {resultado}")

            # Si se encuentran errores, guardar los resultados en un archivo Excel
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                '''
                output_file = 'ERRORES_SECUENCIAS_CONSTRUCCIONES.xlsx'
                df_resultado.to_excel(output_file, sheet_name='ErroresSecuencias', index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} errores.")
                '''
            else:
                messagebox.showinfo("Sin errores", "Todas las secuencias en Construcciones están presentes en ConstruccionesGenerales.")
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
    
    def validar_edad_construccion(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Construcciones'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                edad_construccion = row.get('EdadConstruccion', None)

                # Verificar si 'EdadConstruccion' es <= 0
                if edad_construccion is not None and edad_construccion <= 0:
                    resultado = {
                        'NroFicha': row.get('NroFicha', 'Sin NroFicha'),
                        'Secuencia': row.get('Secuencia', 'Sin Secuencia'),
                        'EdadConstruccion': edad_construccion,
                        'Observacion': 'Edad de construcción inválida (<= 0)',
                    }
                    resultados.append(resultado)
                    print(f"Fila {index}: Agregado a resultados: {resultado}")

            print(f"Total de errores encontrados: {len(resultados)}")

            # Agregar resultados a la lista general
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    def validar_porcentaje_construido(self):
        """
        Verifica que en la hoja 'Construcciones' no haya valores en la columna 'PorcentajeConstruido' 
        que sean iguales o menores a 0. Si los hay, genera un error por cada registro que cumple la condición.
        """
        archivo_excel = self.archivo_entry.get()
        
        
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'Construcciones'
            df_construcciones = pd.read_excel(archivo_excel, sheet_name='Construcciones')
            
            # Filtrar registros donde 'PorcentajeConstruido' es menor o igual a 0
            errores = df_construcciones[df_construcciones['PorcentajeConstruido'] <= 0]
            
            resultados = []

            # Generar lista de errores
            for _, row in errores.iterrows():
                resultado = {
                    'NroFicha': row['NroFicha'],
                    'PorcentajeConstruido': row['PorcentajeConstruido'],
                    'Observacion': 'El valor de PorcentajeConstruido es igual o menor a 0',
                    'Nombre Hoja': 'Construcciones'
                }
                resultados.append(resultado)

            # Guardar los resultados en un archivo Excel si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_PorcentajeConstruido_Construcciones.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros con PorcentajeConstruido igual o menor a 0.")
            else:
                messagebox.showinfo("Sin errores", "No se encontraron valores de PorcentajeConstruido igual o menor a 0 en la hoja 'Construcciones'.")

            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
    def validar_puntos_construcciones(self):
        """
        Verifica que en la hoja 'Construcciones' no haya valores en la columna 'Puntos' 
        que sean menores a 1 o estén vacíos (null). Si los hay, genera un error por cada 
        registro que cumpla alguna de estas condiciones.
        """
        archivo_excel = self.archivo_entry.get()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'Construcciones'
            df_construcciones = pd.read_excel(archivo_excel, sheet_name='Construcciones')
            
            # Filtrar registros donde 'Puntos' es menor a 1 o es null
            errores = df_construcciones[(df_construcciones['Puntos'] < 1) | (df_construcciones['Puntos'].isnull())]
            
            resultados = []

            # Generar lista de errores
            for _, row in errores.iterrows():
                resultado = {
                    'NroFicha': row['NroFicha'],
                    'Puntos': row['Puntos'],
                    'Observacion': 'El valor de Puntos es menor a 1 o está vacío',
                    'Nombre Hoja': 'Construcciones'
                }
                resultados.append(resultado)

            # Guardar los resultados en un archivo Excel si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_Puntos_Construcciones.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros con Puntos menor a 1 o vacío.")
            else:
                messagebox.showinfo("Sin errores", "No se encontraron valores de Puntos menores a 1 o vacíos en la hoja 'Construcciones'.")

            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []