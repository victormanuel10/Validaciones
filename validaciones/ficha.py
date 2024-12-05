# -*- coding: utf-8 -*-
import pandas as pd
from tkinter import messagebox

class Ficha:
    def __init__(self, archivo_entry):
        self.archivo_entry = archivo_entry
        
        
    def terreno_cero(self):
        
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        

        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
            
            print("Leyendo archivo: {}, Hoja: {}".format(archivo_excel, nombre_hoja))
            print(f"funcion: terreno_cero")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                valor_b = row['Npn']
                valor_p = row['AreaTotalTerreno']

                # Verificar si valor_b no es nulo o vacío, y si tiene al menos 22 caracteres
                if pd.notna(valor_b) and len(str(valor_b)) > 21:
                    valor_b_str = str(valor_b)  # Convertir el valor a cadena
                    print(f"Fila {index}: Valor B = '{valor_b_str}', condicion: {valor_b_str[21]}, Valor P = '{valor_p}'")

                    # Verificar las condiciones
                    if valor_b_str[21] == '0' and (valor_p == '0' or valor_p == 0):
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'Observacion': 'Terreno en ceros para ficha que no es mejora',
                            'Npn':row['Npn'],
                            'DestinoEconomico': row['DestinoEcconomico'],
                            'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                            'AreaTotalConstruida':row['AreaTotalConstruida'],
                            'CaracteristicaPredio':row['CaracteristicaPredio'],
                            'AreaTotalTerreno':row['AreaTotalTerreno'],
                            'DireccionReal':row['DireccionReal'],
                            'ModoAdquisicion':row['ModoAdquisicion'],
                            'Tomo':row['Tomo'],
                            'PredioLcTipo':row['PredioLcTipo'],
                            'NumCedulaCatastral':row['NumCedulaCatastral'],
                            'AreaTotalLote':row['AreaTotalLote'],
                            'AreaLoteComun':row['AreaLoteComun'],
                            'AreaLotePrivada':row['AreaLotePrivada'],
                            'Radicado':row['Radicado'],
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
                else:
                    print(f"Fila {index}: NumCedulaCatastral no tiene suficientes caracteres o es nulo.")

            print(f"Total de resultados encontrados: {len(resultados)}")

            # Crear un nuevo DataFrame con los resultados
            
            df_resultado = pd.DataFrame(resultados)
            
            #messagebox.showinfo("Éxito", f"Proceso completado Terreno cero. con {len(resultados)} registros.")
            return resultados

        except Exception as e:
                print(f"Error: {str(e)}")
                messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    
    def terreno_null(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return []  # Devolver lista vacía si no se encuentra el archivo
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: terreno_null")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                valor_a = str(row['Npn'])  # Convertir a cadena por si acaso
                valor_b = row['AreaTotalTerreno']

            print(f"Fila {index}: Valor A = '{valor_a}'")

            # Verificar que valor_a tenga al menos 22 caracteres
            if len(valor_a) > 21:
                print(f"CARACTER22B : {valor_a[21]}")

                if valor_a[21] == '2' and pd.isna(valor_b):
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'Observacion': 'Terreno nulo para condición de predio',
                        'Npn':row['Npn'],
                        'DestinoEconomico': row['DestinoEcconomico'],
                        'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                        'AreaTotalConstruida':row['AreaTotalConstruida'],
                        'CaracteristicaPredio':row['CaracteristicaPredio'],
                        'AreaTotalTerreno':row['AreaTotalTerreno'],
                        'DireccionReal':row['DireccionReal'],
                        'ModoAdquisicion':row['ModoAdquisicion'],
                        'Tomo':row['Tomo'],
                        'PredioLcTipo':row['PredioLcTipo'],
                        'NumCedulaCatastral':row['NumCedulaCatastral'],
                        'AreaTotalLote':row['AreaTotalLote'],
                        'AreaLoteComun':row['AreaLoteComun'],
                        'AreaLotePrivada':row['AreaLotePrivada'],    
                        'Condicion de predio': valor_a[21],
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
            else:
                print(f"Fila {index}: Valor A tiene menos de 22 caracteres, se omite.")
                    
            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados) if resultados else pd.DataFrame(columns=[
                'NroFicha', 'NumCedulaCatastral', 'Condicion de predio', 'AreaTotalTerreno', 'Observacion'])

            # Guardar el resultado en un nuevo archivo Excel
            '''
            
            output_file = 'TERRENO_NULL.xlsx'
            sheet_name = 'TERRENO_NULL'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")
            messagebox.showinfo("Éxito",
                                f"Proceso completado Terreno null. con {len(resultados)} registros.")
         
            '''
            
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []  # Devolver lista vacía en caso de error
            
    def informal_matricula(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"Función: informal_matricula")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Verificar que las columnas requeridas existen
            columnas_requeridas = ['MatriculaInmobiliaria', 'ModoAdquisicion', 'NroFicha']
            if not all(col in df.columns for col in columnas_requeridas):
                messagebox.showerror("Error", f"Faltan columnas requeridas en la hoja '{nombre_hoja}': {columnas_requeridas}")
                return

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                valor_a = row.get('MatriculaInmobiliaria', '')
                valor_b = row.get('ModoAdquisicion', '')

                print(f"Fila {index}: MatriculaInmobiliaria = '{valor_a}', ModoAdquisicion = '{valor_b}'")

                # Verificar las condiciones: valor_b es '2|POSESIÓN' y valor_a NO está vacío
                if valor_b == '2|POSESIÓN' and (valor_a != '' and pd.notna(valor_a)):
                    resultado = {
                        'NroFicha': row.get('NroFicha', ''),
                        'Observacion': 'Modo de adquisición posesión con matrícula',
                        'Npn':row['Npn'],
                        'DestinoEconomico': row['DestinoEcconomico'],
                        'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                        'AreaTotalConstruida':row['AreaTotalConstruida'],
                        'CaracteristicaPredio':row['CaracteristicaPredio'],
                        'AreaTotalTerreno':row['AreaTotalTerreno'],
                        'DireccionReal':row['DireccionReal'],
                        'ModoAdquisicion':row['ModoAdquisicion'],
                        'Tomo':row['Tomo'],
                        'PredioLcTipo':row['PredioLcTipo'],
                        'NumCedulaCatastral':row['NumCedulaCatastral'],
                        'AreaTotalLote':row['AreaTotalLote'],
                        'AreaLoteComun':row['AreaLoteComun'],
                        'AreaLotePrivada':row['AreaLotePrivada'],
                        'Radicado':row['Radicado'],    
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                
                if valor_b == '5|OCUPACIÓN' and (valor_a != '' and pd.notna(valor_a)):
                    resultado = {
                        'NroFicha': row.get('NroFicha', ''),                        
                        'Observacion': 'Modo de adquisición Ocupacion con matrícula',
                        'Npn':row['Npn'],
                        'DestinoEconomico': row['DestinoEcconomico'],
                        'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                        'AreaTotalConstruida':row['AreaTotalConstruida'],
                        'CaracteristicaPredio':row['CaracteristicaPredio'],
                        'AreaTotalTerreno':row['AreaTotalTerreno'],
                        'DireccionReal':row['DireccionReal'],
                        'ModoAdquisicion':row['ModoAdquisicion'],
                        'Tomo':row['Tomo'],
                        'PredioLcTipo':row['PredioLcTipo'],
                        'NumCedulaCatastral':row['NumCedulaCatastral'],
                        'AreaTotalLote':row['AreaTotalLote'],
                        'AreaLoteComun':row['AreaLoteComun'],
                        'AreaLotePrivada':row['AreaLotePrivada'],    
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

            print(f"Total de resultados encontrados: {len(resultados)}")

            # Crear un nuevo DataFrame con los resultados
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

                # Si deseas guardar los resultados en un archivo Excel, descomenta el siguiente bloque
                '''
                output_file = 'INFORMA_MATRICULA.xlsx'
                sheet_name = 'INFORMAL_MATRICULA'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
                '''
            else:
                print("No se encontraron registros que cumplan con las condiciones.")
                messagebox.showinfo("Información", "No se encontraron registros que cumplan con las condiciones.")

            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    def matricula_mejora(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: matricula_mejora")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                valor_a = str(row['Npn']) if pd.notna(row['Npn']) else ''
                valor_b = row['MatriculaInmobiliaria']

                print(f"Fila {index}: Valor A = '{valor_a}'")
                print(f"Fila {index}: Valor A = '{valor_b}'")
                # Verificar si valor_a tiene al menos 22 caracteres
                if len(valor_a) >= 22:
                    print(f"CARACTER22B : {valor_a[21]}")

                    # Verificar las condiciones
                    if valor_a[21] == '2' and pd.notna(valor_b) or valor_b=='':
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'Observacion': 'Condición de predio 2 con matrícula',
                            'Npn':row['Npn'],
                            'DestinoEconomico': row['DestinoEcconomico'],
                            'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                            'AreaTotalConstruida':row['AreaTotalConstruida'],
                            'CaracteristicaPredio':row['CaracteristicaPredio'],
                            'AreaTotalTerreno':row['AreaTotalTerreno'],
                            'DireccionReal':row['DireccionReal'],
                            'ModoAdquisicion':row['ModoAdquisicion'],
                            'Tomo':row['Tomo'],
                            'PredioLcTipo':row['PredioLcTipo'],
                            'NumCedulaCatastral':row['NumCedulaCatastral'],
                            'AreaTotalLote':row['AreaTotalLote'],
                            'AreaLoteComun':row['AreaLoteComun'],
                            'AreaLotePrivada':row['AreaLotePrivada'],
                            'Condicion de predio': valor_a[21],
                            'Radicado':row['Radicado'],
                            'Nombre Hoja': nombre_hoja
                            
                        }
                        
                        resultados.append(resultado)
                        print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
                else:
                    print(f"Fila {index} tiene una NumCedulaCatastral demasiado corta: '{valor_a}'")

            print(f"Total de resultados encontrados: {len(resultados)}")

            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)
            '''
            # Guardar el resultado en un nuevo archivo Excel
            output_file = 'MATRICULA_MEJORA.xlsx'
            sheet_name = 'MATRICULA_MEJORA'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            messagebox.showinfo("Éxito",
                                f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
            
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
      
            
    def circulo_mejora(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: circulo_mejora")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iloc[0:].iterrows():
                valor_a = row['Npn']
                valor_b = row['circulo']

            print(f"Fila {index}: Valor A = '{valor_a}'")

            # Verificar que 'valor_a' tiene al menos 22 caracteres
            if len(valor_a) > 21:
                print(f"CARACTER22B : {valor_a[21]}")

                # Verificar las condiciones
                if valor_a[21] == '2' and pd.notna(valor_b):
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'Observacion': 'Informalidad con matrícula',
                        'Npn':row['Npn'],
                        'Npn':row['Npn'],
                        'DestinoEconomico': row['DestinoEcconomico'],
                        'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                        'AreaTotalConstruida':row['AreaTotalConstruida'],
                        'CaracteristicaPredio':row['CaracteristicaPredio'],
                        'AreaTotalTerreno':row['AreaTotalTerreno'],
                        'DireccionReal':row['DireccionReal'],
                        'ModoAdquisicion':row['ModoAdquisicion'],
                        'Tomo':row['Tomo'],
                        'PredioLcTipo':row['PredioLcTipo'],
                        'NumCedulaCatastral':row['NumCedulaCatastral'],
                        'AreaTotalLote':row['AreaTotalLote'],
                        'AreaLoteComun':row['AreaLoteComun'],
                        'AreaLotePrivada':row['AreaLotePrivada'],
                        'Condicion de predio': valor_a[21],
                        'Radicado':row['Radicado'],
                        'circulo': row['circulo'],
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
            else:
                print(f"El valor de 'NumCedulaCatastral' en la fila {index} no tiene suficientes caracteres.")

            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)
            '''
            # Guardar el resultado en un nuevo archivo Excel
            output_file = 'CIRCULO_MEJORA.xlsx'
            sheet_name = 'CIRCULO_MEJORA'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            messagebox.showinfo("Éxito",
                                f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
            
            '''
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    def tomo_mejora(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: tomo_mejora")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame, comenzando desde la fila 2
            for index, row in df.iloc[0:].iterrows():
                valor_a = row['Npn']
                
                # Verificar que 'valor_a' no es nan y tiene al menos 22 caracteres antes de continuar
                if pd.notna(valor_a) and len(str(valor_a)) > 21:
                    valor_a = str(valor_a)  # Convertir el valor a string

                    print(f"Fila {index}: Valor A = '{valor_a}'")

                    # Verificar las condiciones
                    if valor_a[21] == '2' and pd.notna(row['Tomo']) and float(row['Tomo']) != 0:
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'Observacion': 'Informalidad con Tomo',
                            'Npn':row['Npn'],
                            'DestinoEconomico': row['DestinoEcconomico'],
                            'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                            'AreaTotalConstruida':row['AreaTotalConstruida'],
                            'CaracteristicaPredio':row['CaracteristicaPredio'],
                            'AreaTotalTerreno':row['AreaTotalTerreno'],
                            'DireccionReal':row['DireccionReal'],
                            'ModoAdquisicion':row['ModoAdquisicion'],
                            'Tomo':row['Tomo'],
                            'PredioLcTipo':row['PredioLcTipo'],
                            'NumCedulaCatastral':row['NumCedulaCatastral'],
                            'AreaTotalLote':row['AreaTotalLote'],
                            'AreaLoteComun':row['AreaLoteComun'],
                            'AreaLotePrivada':row['AreaLotePrivada'],
                            'Condicion de predio': valor_a[21],
                            'Radicado':row['Radicado'],
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                else:
                    print(f"El valor de 'Npn' en la fila {index} es inválido o no tiene suficientes caracteres.")

            print(f"Total de resultados encontrados: {len(resultados)}")
            
            '''
            
            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)
            
            # Guardar el resultado en un nuevo archivo Excel
            output_file = 'TOMO_MEJORA.xlsx'
            sheet_name = 'TOMO_MEJORA'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            messagebox.showinfo("Éxito",
                                f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
            '''
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
    
    def modo_adquisicion_informal(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: modo_adquisicion_informal")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame, comenzando desde la fila 2
            for index, row in df.iloc[0:].iterrows():
                valor_a = row['Npn']
                valor_b = row['ModoAdquisicion']

                print(f"Fila {index}: Valor A = '{valor_a}'")

                # Convertir valor_a a cadena si no lo es
                if not isinstance(valor_a, str):
                    valor_a = str(valor_a)

                # Verificar que 'valor_a' tiene al menos 22 caracteres antes de acceder al índice 21
                if len(valor_a) > 21:
                    # Verificar las condiciones
                    if valor_a[21] == '2' and (valor_b not in ['5|OCUPACIÓN', '2|POSESIÓN']):
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'Observacion': 'Condición de predio 2 con modo de adquisición diferente a posesión u ocupacion',
                            'Npn': row['Npn'],
                            'DestinoEconomico': row['DestinoEcconomico'],
                            'MatriculaInmobiliaria': row['MatriculaInmobiliaria'],
                            'AreaTotalConstruida': row['AreaTotalConstruida'],
                            'CaracteristicaPredio': row['CaracteristicaPredio'],
                            'AreaTotalTerreno': row['AreaTotalTerreno'],
                            'DireccionReal':row['DireccionReal'],
                            'ModoAdquisicion': row['ModoAdquisicion'],
                            'Tomo': row['Tomo'],
                            'PredioLcTipo': row['PredioLcTipo'],
                            'NumCedulaCatastral': row['NumCedulaCatastral'],
                            'AreaTotalLote': row['AreaTotalLote'],
                            'AreaLoteComun': row['AreaLoteComun'],
                            'AreaLotePrivada': row['AreaLotePrivada'],
                            'Condicion de predio': valor_a[21],
                            'Radicado':row['Radicado'],
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
                else:
                    print(f"El valor de 'Npn' en la fila {index} no tiene suficientes caracteres.")

            print(f"Total de resultados encontrados: {len(resultados)}")

            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)
            '''
            # Guardar el resultado en un nuevo archivo Excel
            output_file = 'MODO_ADQUISICION_INFORMAL.xlsx'
            sheet_name = 'MODO_ADQUISICION_INFORMAL'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            messagebox.showinfo("Éxito",
                                f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
            
            '''
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}") 
            
    def validar_modo_adquisicion_caracteristica(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_modo_adquisicion_caracteristica")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                modo_adquisicion = row.get('ModoAdquisicion', None)
                caracteristica_predio = row.get('CaracteristicaPredio', None)
                PredioLcTipo=row.get('PredioLcTipo',None)
                # Verificar que ModoAdquisicion sea igual a '2|POSESIÓN'
                if modo_adquisicion in ['2|POSESIÓN', '5|OCUPACIÓN']:
                    # Si CaracteristicaPredio es diferente de '12|INFORMAL (2)', generar error
                    if caracteristica_predio != '12|INFORMAL (2)' and PredioLcTipo!='Predio.Publico.Presunto_Baldio':
                        resultado = {
                            'NroFicha': row.get('NroFicha', 'Sin valor'),
                            'Observacion': 'Caracteristica incorrecta para Modo de Aquisición Ocupación o Posesión',
                            'Npn':row['Npn'],
                            'DestinoEconomico': row['DestinoEcconomico'],
                            'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                            'AreaTotalConstruida':row['AreaTotalConstruida'],
                            'CaracteristicaPredio':row['CaracteristicaPredio'],
                            'AreaTotalTerreno':row['AreaTotalTerreno'],
                            'DireccionReal':row['DireccionReal'],
                            'ModoAdquisicion':row['ModoAdquisicion'],
                            'Tomo':row['Tomo'],
                            'PredioLcTipo':row['PredioLcTipo'],
                            'NumCedulaCatastral':row['NumCedulaCatastral'],
                            'AreaTotalLote':row['AreaTotalLote'],
                            'AreaLoteComun':row['AreaLoteComun'],
                            'AreaLotePrivada':row['AreaLotePrivada'],
                            'Radicado':row['Radicado'],
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        print(f"Fila {index} cumple las condiciones de error. Agregado: {resultado}")

            print(f"Total de resultados encontrados: {len(resultados)}")

            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)
            '''
            # Guardar el resultado en un nuevo archivo Excel
            output_file = 'VALIDACION_MODO_CARACTERISTICA.xlsx'
            sheet_name = 'VALIDACION_MODO_CARACTERISTICA'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            messagebox.showinfo("Éxito",
                                f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
            '''
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
           
            
    def ficha_repetida(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: ficha_repetida")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Verificar si hay duplicados en la columna 'NroFicha'
            duplicados = df[df.duplicated(subset='NroFicha', keep=False)]  # Detectar duplicados
            
            print(f"Total de registros duplicados encontrados: {duplicados.shape[0]}")
            resultados = []
            if not duplicados.empty:
                for index, row in duplicados.iterrows():
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'Observacion': 'NroFicha duplicado',
                        'Npn':row['Npn'],
                        'DestinoEconomico': row['DestinoEcconomico'],
                       'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                        'AreaTotalConstruida':row['AreaTotalConstruida'],
                        'CaracteristicaPredio':row['CaracteristicaPredio'],
                        'AreaTotalTerreno':row['AreaTotalTerreno'],
                        'ModoAdquisicion':row['ModoAdquisicion'],
                        'Tomo':row['Tomo'],
                        'PredioLcTipo':row['PredioLcTipo'],
                        'NumCedulaCatastral':row['NumCedulaCatastral'],
                        'AreaTotalLote':row['AreaTotalLote'],
                        'AreaLoteComun':row['AreaLoteComun'],
                        'AreaLotePrivada':row['AreaLotePrivada'],
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': 'Fichas'
                    }
                    resultados.append(resultado)
                    print(f"Fila {index}: NroFicha duplicado encontrado: {resultado}")
                '''
                
                # Guardar resultados en un archivo Excel si existen duplicados
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_NroFicha_Duplicados_Fichas.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Se encontraron {len(resultados)} registros duplicados en 'NroFicha'.")
                else:
                messagebox.showinfo("Sin duplicados", "No se encontraron registros duplicados en 'NroFicha'.")
                '''
            return duplicados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
        
    def rural_destino_invalido(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
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

            # Iterar sobre las filas del DataFrame, comenzando desde la fila 2
            for index, row in df.iloc[0:].iterrows():
                valor_b = str(row['NumCedulaCatastral'])
                valor_p = row['DestinoEcconomico']

                print(f"Fila {index}: Valor B = '{valor_b}',condicion:{valor_b[6]}, Valor P = '{valor_p}'")

                # Verificar las condiciones
                if valor_b[6] == '0' and (valor_p == '12|LOTE URBANIZADO NO CONSTRUIDO' or valor_p == '13|LOTE URBANIZABLE NO URBANIZADO' or valor_p == '14|LOTE NO URBANIZABLE'):
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        
                        'Observacion': 'En sector rural no es valido destinaciones 12,13 y 14',
                        'Npn':row['Npn'],
                        'DestinoEconomico': row['DestinoEcconomico'],
                        'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                        'AreaTotalConstruida':row['AreaTotalConstruida'],
                        'CaracteristicaPredio':row['CaracteristicaPredio'],
                        'AreaTotalTerreno':row['AreaTotalTerreno'],
                        'DireccionReal':row['DireccionReal'],
                        'ModoAdquisicion':row['ModoAdquisicion'],
                        'Tomo':row['Tomo'],
                        'PredioLcTipo':row['PredioLcTipo'],
                        'NumCedulaCatastral':row['NumCedulaCatastral'],
                        'AreaTotalLote':row['AreaTotalLote'],
                        'AreaLoteComun':row['AreaLoteComun'],
                        'AreaLotePrivada':row['AreaLotePrivada'],
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

            print(f"Total de resultados encontrados: {len(resultados)}")


            df_resultado = pd.DataFrame(resultados)
            '''
            output_file = 'rural_destino_invalido.xlsx'
            sheet_name = 'rural_destino_invalido'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            '''
            
            messagebox.showinfo("Éxito",
                                f"Proceso completado Rural destino invalido.' con {len(resultados)} registros.")
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
    def destino_economico_mayorcero(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        
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
                destino_economico = row['DestinoEcconomico']
                area_total_construida = row['AreaTotalConstruida']

                if destino_economico in ['12|Lote_Urbanizado_No_Construido', 
                                        '13|Lote_Urbanizable_No_Urbanizado', 
                                        '14|Lote_No_Urbanizable',] and area_total_construida > 0:
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'Observacion': 'Destino económico 12, 13 y 14 no debe tener área construida mayor a cero',
                        'Npn':row['Npn'],
                        'DestinoEconomico': row['DestinoEcconomico'],
                        'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                        'AreaTotalConstruida':row['AreaTotalConstruida'],
                        'CaracteristicaPredio':row['CaracteristicaPredio'],
                        'AreaTotalTerreno':row['AreaTotalTerreno'],
                        'DireccionReal':row['DireccionReal'],
                        'ModoAdquisicion':row['ModoAdquisicion'],
                        'Tomo':row['Tomo'],
                        'PredioLcTipo':row['PredioLcTipo'],
                        'NumCedulaCatastral':row['NumCedulaCatastral'],
                        'AreaTotalLote':row['AreaTotalLote'],
                        'AreaLoteComun':row['AreaLoteComun'],
                        'AreaLotePrivada':row['AreaLotePrivada'],
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': nombre_hoja
                        
                    }
                    resultados.append(resultado)
                    print(f"Fila {index}: Agregado a resultados: {resultado}")

            print(f"Total de errores encontrados: {len(resultados)}")
            
            if resultados:
                # Crear un nuevo DataFrame con los resultados
                df_resultado = pd.DataFrame(resultados)
                '''
                output_file = 'ERRORES_DESTINO_ECONOMICO.xlsx'
                sheet_name = 'ErroresDestinoEconomico'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Proceso completado Destino Economico 12 13 14.con {len(resultados)} errores.")
            
                '''
                
            else:
                print("No se encontraron errores.")
                messagebox.showinfo("Información", "No se encontraron registros con errores.")
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
    def areaterrenocero(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                areatotalterreno = row['AreaTotalTerreno']
                area_total_construida = row['AreaTotalConstruida']
                Npn = str(row.get('Npn', '')).strip()

                # Validar que el 22º dígito de Npn no sea 8 ni 9
                if not (len(Npn) >= 22 and Npn[21] in ['8', '9']):
                    if areatotalterreno == '' or areatotalterreno == 0 or pd.isna(areatotalterreno):
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            
                            'Observacion': 'Área de terreno invalida para característica diferente a RPH o Condominio',
                            'Npn':row['Npn'],
                            'DestinoEconomico': row['DestinoEcconomico'],
                            'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                            'AreaTotalConstruida':row['AreaTotalConstruida'],
                            'CaracteristicaPredio':row['CaracteristicaPredio'],
                            'AreaTotalTerreno':row['AreaTotalTerreno'],
                            'DireccionReal':row['DireccionReal'],
                            'ModoAdquisicion':row['ModoAdquisicion'],
                            'Tomo':row['Tomo'],
                            'PredioLcTipo':row['PredioLcTipo'],
                            'NumCedulaCatastral':row['NumCedulaCatastral'],
                            'AreaTotalLote':row['AreaTotalLote'],
                            'AreaLoteComun':row['AreaLoteComun'],
                            'AreaLotePrivada':row['AreaLotePrivada'],
                            'Radicado':row['Radicado'],
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                    if areatotalterreno>0 and areatotalterreno<4:
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'AreaTotalTerreno': areatotalterreno,
                            'AreaTotalConstruida': area_total_construida,
                            'Observacion': 'Area terreno menor a 4 (aviso)',
                            'Npn': row['Npn'],
                            'Radicado':row['Radicado'],
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        print(f"Fila {index}: Agregado a resultados: {resultado}")

            print(f"Total de errores encontrados: {len(resultados)}")
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
    
    
    def validar_area_construida_fichas_construcciones(self):
        archivo_excel = self.archivo_entry.get()
        hoja_fichas = 'Fichas'
        hoja_construcciones = 'Construcciones'

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return

        try:
            # Leer las hojas Fichas y Construcciones
            df_fichas = pd.read_excel(archivo_excel, sheet_name=hoja_fichas)
            df_construcciones = pd.read_excel(archivo_excel, sheet_name=hoja_construcciones)

            print(f"Leyendo archivo: {archivo_excel}")
            print(f"Dimensiones de Fichas: {df_fichas.shape}")
            print(f"Dimensiones de Construcciones: {df_construcciones.shape}")

            # Validar columnas necesarias
            columnas_necesarias_fichas = ['NroFicha', 'AreaTotalConstruida', 'Npn']
            columnas_necesarias_construcciones = ['NroFicha']

            for columna in columnas_necesarias_fichas:
                if columna not in df_fichas.columns:
                    messagebox.showerror("Error", f"La columna '{columna}' no existe en la hoja {hoja_fichas}.")
                    return
            for columna in columnas_necesarias_construcciones:
                if columna not in df_construcciones.columns:
                    messagebox.showerror("Error", f"La columna '{columna}' no existe en la hoja {hoja_construcciones}.")
                    return

            # Obtener los NroFicha que están en la hoja Construcciones
            fichas_con_construcciones = df_construcciones['NroFicha'].unique()

            # Filtrar las filas en Fichas con esas NroFichas
            fichas_validar = df_fichas[df_fichas['NroFicha'].isin(fichas_con_construcciones)]

            # Excluir registros donde el 22.º carácter de Npn sea igual a 9 o 8
            fichas_validar = fichas_validar[
                ~fichas_validar['Npn'].str[21].isin(['8', '9'])
            ]

            # Filtrar las filas donde AreaTotalConstruida <= 0 o es nulo
            errores = fichas_validar[
                (fichas_validar['AreaTotalConstruida'] <= 0) |
                (fichas_validar['AreaTotalConstruida'].isna())
            ]

            # Resultados a mostrar
            resultados = []
            for _, row in errores.iterrows():
                resultado = {
                    'NroFicha': row['NroFicha'],
                    'Observacion': 'Área Total Construida en fichas debe ser mayor a 0 para las fichas con construcciones',
                    'Npn': row['Npn'],
                    'DestinoEconomico': row['DestinoEcconomico'],
                    'MatriculaInmobiliaria': row['MatriculaInmobiliaria'],
                    'AreaTotalConstruida': row['AreaTotalConstruida'],
                    'CaracteristicaPredio': row['CaracteristicaPredio'],
                    'AreaTotalTerreno': row['AreaTotalTerreno'],
                    'DireccionReal':row['DireccionReal'],
                    'ModoAdquisicion': row['ModoAdquisicion'],
                    'Tomo': row['Tomo'],
                    'PredioLcTipo': row['PredioLcTipo'],
                    'NumCedulaCatastral': row['NumCedulaCatastral'],
                    'AreaTotalLote': row['AreaTotalLote'],
                    'AreaLoteComun': row['AreaLoteComun'],
                    'AreaLotePrivada': row['AreaLotePrivada'],
                    'Radicado':row['Radicado'],
                    'Nombre Hoja': 'Fichas'
                }
                resultados.append(resultado)
                print(f"Error encontrado: {resultado}")

            # Retornar los resultados
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
    
    '''    
    def validar_area_construida(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            # Leer el archivo Excel
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Normalizar la columna 'DestinoEcconomico' (convertir a cadena y eliminar inconsistencias)
            df['DestinoEcconomico'] = df['DestinoEcconomico'].astype(str).str.strip().str.upper()

            # Lista de valores de DestinoEconomico a excluir (convertidos a mayúsculas)
            excluir_destinos = [
                "12|LOTE_URBANIZADO_NO_CONSTRUIDO",
                "13|LOTE_URBANIZABLE_NO_URBANIZADO",
                "14|LOTE_NO_URBANIZABLE",
                "0|NA",
                "24|AGRICOLA",
                "61|AGROFORESTAL",
                "30|FORESTAL",
                "60|ACUICOLA",
                "63|INFRAESTRUCTURA_HIDRAULICA",
                "19|USO_PUBLICO",
                ""
            ]
            excluir_destinos = [destino.upper() for destino in excluir_destinos]

            # Filtrar el DataFrame eliminando los registros con los destinos a excluir
            df_filtrado = df[~df['DestinoEcconomico'].isin(excluir_destinos)]

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame filtrado
            for index, row in df_filtrado.iterrows():
                area_total_construida = row['AreaTotalConstruida']
                Npn = str(row.get('Npn', '')).strip()

                # Validar condición
                if not (len(Npn) >= 22 and Npn[21] in ['8', '9']):
                    # Verificar si el área es nula o menor o igual a cero
                    if pd.isna(area_total_construida) or area_total_construida <= 0:
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'Observacion': f'Área Total Construida es cero o null para destino económico: {row["DestinoEcconomico"]}',
                            'Npn': row['Npn'],
                            'DestinoEconomico': row['DestinoEcconomico'],
                            'MatriculaInmobiliaria': row['MatriculaInmobiliaria'],
                            'AreaTotalConstruida': row['AreaTotalConstruida'],
                            'CaracteristicaPredio': row['CaracteristicaPredio'],
                            'AreaTotalTerreno': row['AreaTotalTerreno'],
                            'ModoAdquisicion': row['ModoAdquisicion'],
                            'Tomo': row['Tomo'],
                            'PredioLcTipo': row['PredioLcTipo'],
                            'NumCedulaCatastral': row['NumCedulaCatastral'],
                            'AreaTotalLote': row['AreaTotalLote'],
                            'AreaLoteComun': row['AreaLoteComun'],
                            'AreaLotePrivada': row['AreaLotePrivada'],
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        print(f"Fila {index}: Agregado a resultados: {resultado}")

            print(f"Total de errores encontrados: {len(resultados)}")
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
    '''        
            
    def predios_con_direcciones_invalidas(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            resultados = []

            # Lista de palabras no permitidas
            palabras_no_permitidas = ['ZONA', 'BLOQUE', 'Bloque', 'EDIFICIO', 'Edificio', 'LOS', 'BARRIO', 'Barrio', 
                                    'VIA', 'Via', 'Lote', 'LOTE', 'CALLE', 'calle', 'AVENIDA', 'avenida', 
                                    'CRA', 'Cra', 'KL', 'CARRERA', 'Carrera', 'Diagonal','S.N','S.N.','SN','S.D','s.d']

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                MatriculaInmobiliaria = str(row['NumCedulaCatastral']).strip()
                print(f"Fila {index}: NumCedulaCatastral = '{MatriculaInmobiliaria}'")
                if len(MatriculaInmobiliaria) >= 4 and MatriculaInmobiliaria[3] == '1':
                        DireccionReal = row['DireccionReal']
                        # Validar si la dirección está vacía
                        if not DireccionReal or pd.isna(DireccionReal):
                            observacion = 'Predio sin dirección'
                        # Validar si los primeros 8 caracteres contienen palabras no permitidas
                        elif any(palabra in DireccionReal[:8] for palabra in palabras_no_permitidas):
                            observacion = 'Contiene palabras no permitidas en dirección'
                        else:
                            continue

                        # Agregar a resultados
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            
                            'Observacion': observacion,
                            'Npn':row['Npn'],
                            'DestinoEconomico': row['DestinoEcconomico'],
                            'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                            'AreaTotalConstruida':row['AreaTotalTerreno'],
                            'CaracteristicaPredio':row['CaracteristicaPredio'],
                            'AreaTotalTerreno':row['AreaTotalTerreno'],
                            'Direccion': DireccionReal if not pd.isna(DireccionReal) else '',
                            'ModoAdquisicion':row['ModoAdquisicion'],
                            'Tomo':row['Tomo'],
                            'PredioLcTipo':row['PredioLcTipo'],
                            'NumCedulaCatastral':row['NumCedulaCatastral'],
                            'AreaTotalLote':row['AreaTotalLote'],
                            'AreaLoteComun':row['AreaLoteComun'],
                            'AreaLotePrivada':row['AreaLotePrivada'],
                            'Radicado':row['Radicado'],
                            
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        print(f"Fila {index}: Agregado a resultados: {resultado}")

                print(f"Total de errores encontrados: {len(resultados)}")
            '''
            
            if resultados:
                # Crear un nuevo DataFrame con los resultados
                df_resultado = pd.DataFrame(resultados)
                
                output_file = 'Direcciones_invalidas.xlsx'
                sheet_name = 'Direcciones'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Se encontraron {len(resultados)} registros con errores.")
            else:
                print("No se encontraron direcciones inválidas.")
                messagebox.showinfo("Información", "No se encontraron registros con errores en las direcciones.")
            '''
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
        
    def validar_nrofichas_propietarios(self):
        archivo_excel = self.archivo_entry.get()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return []
        
        try:
            # Leer las hojas Propietarios y Fichas
            df_propietarios = pd.read_excel(archivo_excel, sheet_name='Propietarios')
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            # Asegurarse de que las columnas 'NroFicha' sean de tipo string y sin espacios en blanco
            df_propietarios['NroFicha'] = df_propietarios['NroFicha'].astype(str).str.strip()
            df_fichas['NroFicha'] = df_fichas['NroFicha'].fillna('').astype(str).str.strip()

            df_propietarios['NroFicha'] = pd.to_numeric(df_propietarios['NroFicha'], errors='coerce').fillna(0).astype(int)
            df_fichas['NroFicha'] = pd.to_numeric(df_fichas['NroFicha'], errors='coerce').fillna(0).astype(int)
            
            # Obtener los valores únicos y no nulos de 'NroFicha'
            nro_ficha_propietarios = set(df_propietarios['NroFicha'].dropna().unique())
            nro_ficha_fichas = set(df_fichas['NroFicha'].dropna().unique())

            # Verificar si hay NroFicha en Propietarios que no están en Fichas
            fichas_faltantes_en_fichas = nro_ficha_propietarios - nro_ficha_fichas

            resultados = []

            for nro_ficha in fichas_faltantes_en_fichas:
                resultado = {
                    'NroFicha': nro_ficha,
                    'Observacion': 'NroFicha en Propietarios no está en Fichas',

                    'Nombre Hoja': 'Propietarios'
                }
                resultados.append(resultado)
            
            '''
            
            if resultados:
                # Crear un nuevo DataFrame con los resultados
                df_resultado = pd.DataFrame(resultados)
                
                output_file = 'Fichas Falta.xlsx'
                sheet_name = 'Direcciones'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Se encontraron {len(resultados)} registros con errores.")    
            else:
                messagebox.showinfo("Información", "No faltan fichas en Fichas desde Propietarios.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
    def validar_fichas_en_propietarios(self):
        archivo_excel = self.archivo_entry.get()

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return []

        try:
            # Leer las hojas Propietarios y Fichas
            df_propietarios = pd.read_excel(archivo_excel, sheet_name='Propietarios')
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            # Asegurarse de que las columnas 'NroFicha' sean tipo entero y sin espacios en blanco
            df_propietarios['NroFicha'] = df_propietarios['NroFicha'].astype(str).str.strip()
            df_fichas['NroFicha'] = df_fichas['NroFicha'].fillna('').astype(str).str.strip()

            # Convertir las columnas 'NroFicha' a entero, asegurándose de que no haya valores flotantes
            df_propietarios['NroFicha'] = pd.to_numeric(df_propietarios['NroFicha'], errors='coerce').fillna(0).astype(int)
            df_fichas['NroFicha'] = pd.to_numeric(df_fichas['NroFicha'], errors='coerce').fillna(0).astype(int)

            # Manejar la columna 'Npn', asegurarse de que sea string y rellenar valores nulos
            if 'Npn' in df_fichas.columns:
                df_fichas['Npn'] = df_fichas['Npn'].fillna('').astype(str).str.strip()

                # Mostrar el número de registros antes de filtrar
                print(f"Total registros en Fichas antes de filtrar: {len(df_fichas)}")

                # Filtrar registros donde el dígito 22 de 'Npn' sea 9 y termine con tres ceros
                condicion_excepcion = (
                    (df_fichas['Npn'].str[21:22] == '9') |
                    (df_fichas['Npn'].str[21:22] == '8') &
                    (df_fichas['Npn'].str[-3:] == '000') |
                    (df_fichas['Npn'].str[21:22] == '2')
                )
                df_fichas_excluidos = df_fichas[condicion_excepcion]

                # Mostrar los registros que fueron excluidos
                print(f"Registros excluidos (que cumplen con la condición de Npn):")
                print(df_fichas_excluidos[['NroFicha', 'Npn']])

                # Filtrar fuera los registros que cumplen la condición de exclusión
                df_fichas = df_fichas[~condicion_excepcion]

                # Mostrar el número de registros después de filtrar
                print(f"Total registros en Fichas después de filtrar: {len(df_fichas)}")
                print(df_fichas)
                print(df_propietarios)
                
            df_fichas['NroFicha'] = pd.to_numeric(df_fichas['NroFicha'], errors='coerce')
            df_propietarios['NroFicha'] = pd.to_numeric(df_propietarios['NroFicha'], errors='coerce')
            # Obtener los valores únicos de NroFicha de ambas hojas
            nro_fichas_fichas = set(df_fichas['NroFicha'].dropna().unique())
            nro_fichas_propietarios = set(df_propietarios['NroFicha'].dropna().unique())

            # Encontrar fichas que están en Fichas pero no en Propietarios
            fichas_faltantes = nro_fichas_fichas - nro_fichas_propietarios

            resultados = []
            for nro_ficha in fichas_faltantes:
                resultado = {
                    'NroFicha': nro_ficha,
                    'Observacion': 'NroFicha está en Fichas pero no en Propietarios',
                    'Nombre Hoja': 'Propietarios'
                }
                resultados.append(resultado)
            '''
            
            # Mostrar resultados
            if resultados:
                # Crear un DataFrame con los resultados
                df_resultado = pd.DataFrame(resultados)
                
                # Guardar en un archivo Excel
                output_file = 'FichasNoEnPropietarios.xlsx'
                df_resultado.to_excel(output_file, sheet_name='Errores', index=False)
                messagebox.showinfo("Éxito", f"Se encontraron {len(resultados)} fichas faltantes. Archivo guardado: {output_file}")
                print(f"Archivo guardado: {output_file}")
            else:
                messagebox.showinfo("Sin errores", "Todas las fichas están presentes en la hoja Propietarios.")
                print("No se encontraron fichas faltantes.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
            
    def porcentaje_litigiocero(self):
        
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        
        try:
        
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
               
                

                if row['PorcentajeLitigio'] != 0 and pd.notna(row['PorcentajeLitigio']):
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'Observacion': 'PorcentajeLitigio diferente de cero',
                        'Npn':row['Npn'],
                        'DestinoEconomico': row['DestinoEcconomico'],
                        'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                        'AreaTotalConstruida':row['AreaTotalConstruida'],
                        'CaracteristicaPredio':row['CaracteristicaPredio'],
                        'AreaTotalTerreno':row['AreaTotalTerreno'],
                        'DireccionReal':row['DireccionReal'],
                        'ModoAdquisicion':row['ModoAdquisicion'],
                        'Tomo':row['Tomo'],
                        'PredioLcTipo':row['PredioLcTipo'],
                        'NumCedulaCatastral':row['NumCedulaCatastral'],
                        'AreaTotalLote':row['AreaTotalLote'],
                        'AreaLoteComun':row['AreaLoteComun'],
                        'AreaLotePrivada':row['AreaLotePrivada'],
                        'PorcentajeLitigio':row['PorcentajeLitigio'],
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Fila {index}: Agregado a resultados: {resultado}")

            print(f"Total de errores encontrados: {len(resultados)}")
            '''
            
            if resultados:
                # Crear un nuevo DataFrame con los resultados
                
                
                df_resultado = pd.DataFrame(resultados)
                
                output_file = 'PorcentajeLitigio.xlsx'
                sheet_name = 'PorcentajeLitigio'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                messagebox.showinfo("Éxito", f"PorcentajeLitigio diferente cero {len(resultados)} errores.")
                print(f"Archivo guardado: {output_file}")
            else:
                print("No se encontraron errores.")
                messagebox.showinfo("Información", "No se encontraron registros PorcentajeLitigiocero.")
            '''
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
    
        
    def validar_npn(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los errores
            errores = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                npn = row['Npn']  
                
                if isinstance(npn, str) and len(npn) >= 22:
                    if npn[21] == '0':
                        resto_caracteres = npn[22:]
                        if any(char != '0' for char in resto_caracteres):
                            # Si no son todos '0', agregar a los errores
                            error = {
                                'NroFicha': row['NroFicha'],
                                
                                'Observacion': 'Condición de predio NPH con número de piso, o edificio, o unidad',
                                'Npn':row['Npn'],
                                'DestinoEconomico': row['DestinoEcconomico'],
                                'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                                'AreaTotalConstruida':row['AreaTotalConstruida'],
                                'CaracteristicaPredio':row['CaracteristicaPredio'],
                                'AreaTotalTerreno':row['AreaTotalTerreno'],
                                'ModoAdquisicion':row['ModoAdquisicion'],
                                'Tomo':row['Tomo'],
                                'PredioLcTipo':row['PredioLcTipo'],
                                'NumCedulaCatastral':row['NumCedulaCatastral'],
                                'AreaTotalLote':row['AreaTotalLote'],
                                'AreaLoteComun':row['AreaLoteComun'],
                                'AreaLotePrivada':row['AreaLotePrivada'],
                                'Radicado':row['Radicado'],
                                'Nombre Hoja': nombre_hoja
                            }
                            errores.append(error)
                            print(f"Error encontrado en la fila {index}: {error}")
                else:
                    print(f"El valor de 'Npn' en la fila {index} no tiene suficientes caracteres.")

            print(f"Total de errores encontrados: {len(errores)}")

            # Si hay errores, guardarlos en un archivo Excel
            '''
            
            if errores:
                df_errores = pd.DataFrame(errores)
                output_file = 'ERRORES_NPN.xlsx'
                sheet_name = 'ErroresNpn'
                df_errores.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Errores", f"Se encontraron {len(errores)} errores y se guardaron en '{output_file}'.")
            else:
                print("No se encontraron errores.")
                messagebox.showinfo("Información", "No se encontraron registros con errores.")
            '''
            return errores

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    def validar_npn14a17(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_npn")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                npn = str(row['Npn'])  # Convertir a cadena para asegurar el acceso por índice

                # Validar que el Npn tenga al menos 17 caracteres antes de acceder a las posiciones 14-17
                if len(npn) >= 17:
                    # Extraer los caracteres de las posiciones 14, 15, 16 y 17
                    subcadena_npn = npn[13:17]

                    # Verificar si las posiciones 14-17 son "0000"
                    if subcadena_npn == "0000":
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'Observacion': 'NPN contiene 0000 en las posiciones 14-17',
                            'Npn':row['Npn'],
                            'DestinoEconomico': row['DestinoEcconomico'],
                            'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                            'AreaTotalConstruida':row['AreaTotalConstruida'],
                            'CaracteristicaPredio':row['CaracteristicaPredio'],
                            'AreaTotalTerreno':row['AreaTotalTerreno'],
                            'ModoAdquisicion':row['ModoAdquisicion'],
                            'Tomo':row['Tomo'],
                            'PredioLcTipo':row['PredioLcTipo'],
                            'NumCedulaCatastral':row['NumCedulaCatastral'],
                            'AreaTotalLote':row['AreaTotalLote'],
                            'AreaLoteComun':row['AreaLoteComun'],
                            'AreaLotePrivada':row['AreaLotePrivada'],
                            'Radicado':row['Radicado'],
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        print(f"Fila {index}: Agregado a resultados: {resultado}")
                else:
                    print(f"El valor de 'Npn' en la fila {index} no tiene suficientes caracteres.")
            
            print(f"Total de errores encontrados: {len(resultados)}")
            '''
            
            if resultados:
                # Crear un nuevo DataFrame con los resultados
                df_resultado = pd.DataFrame(resultados)
                
                # Guardar el resultado en un archivo Excel
                output_file = 'ERRORES_NPN.xlsx'
                sheet_name = 'ErroresNPN'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Validación completada con {len(resultados)} errores.")
                
            else:
                print("No se encontraron errores.")
                messagebox.showinfo("Información", "No se encontraron registros con errores.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    
    def ultimo_digito(self):
        """
        En la hoja 'Fichas', valida que si el dígito en la posición 22 de 'Npn' es '0',
        entonces la suma de los últimos 4 dígitos debe ser igual a 0. Si la suma es mayor a 0, genera un error.
        """
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido y especifica la hoja.")
            return []

        try:
            # Leer la hoja específica 'Fichas'
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja) 
            
            # Convertir 'Npn' a cadena y asegurar que tenga al menos 24 caracteres
            df['Npn'] = df['Npn'].astype(str).str.zfill(24)
            
            resultados = []

            # Iterar sobre las filas para validar la condición
            for index, row in df.iterrows():
                npn = row['Npn']
                
                # Verificar si el dígito en la posición 22 es '0'
                if npn[21] == '0':
                    # Sumar los últimos 4 dígitos
                    suma_ultimos_cuatro = sum(int(d) for d in npn[-4:])
                    
                    # Generar un error si la suma es mayor a 0
                    if suma_ultimos_cuatro > 0:
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'Observacion': 'Predio NPH con número de unidad predial ',
                            'Npn':row['Npn'],
                            'DestinoEconomico': row['DestinoEcconomico'],
                            'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                            'AreaTotalConstruida':row['AreaTotalConstruida'],
                            'CaracteristicaPredio':row['CaracteristicaPredio'],
                            'AreaTotalTerreno':row['AreaTotalTerreno'],
                            'ModoAdquisicion':row['ModoAdquisicion'],
                            'Tomo':row['Tomo'],
                            'PredioLcTipo':row['PredioLcTipo'],
                            'NumCedulaCatastral':row['NumCedulaCatastral'],
                            'AreaTotalLote':row['AreaTotalLote'],
                            'AreaLoteComun':row['AreaLoteComun'],
                            'AreaLotePrivada':row['AreaLotePrivada'],
                            'Radicado':row['Radicado'],
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        print(f"Error agregado: {resultado}")

            # Guardar resultados en archivo si existen errores
            '''
            
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_Npn_Cero_Suma_Digitos_Fichas.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros que cumplen con la condición.")
            else:
                messagebox.showinfo("Sin errores", "No se encontraron registros que cumplan con la condición de Npn con dígito 22 igual a '0' y suma de últimos 4 dígitos mayor a 0.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
                
    
    def validar_destino_economico_y_longitud_cedula(self):
        """
        Verifica que en la hoja 'FichasPrediales':
        1. Si el cuarto dígito de 'NumCedulaCatastral' es '2' y el 'DestinoEconomico' es uno de los valores especificados, se genera un error.
        2. Valida que todos los valores en 'NumCedulaCatastral' tengan exactamente 28 dígitos.
        """
        archivo_excel = self.archivo_entry.get()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'FichasPrediales'
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            # Lista de valores de DestinoEconomico para verificar
            destinos_invalidos = [
                "12|LOTE URBANIZADO NO CONSTRUIDO",
                "13|LOTE URBANIZABLE NO URBANIZADO",
                "14|LOTE NO URBANIZABLE"
            ]
            
            resultados = []

            # Rellenar valores nulos en 'NumCedulaCatastral' y 'DestinoEconomico' con cadenas vacías
            df_fichas['NumCedulaCatastral'] = df_fichas['NumCedulaCatastral'].fillna('').astype(str)
            df_fichas['DestinoEconomico'] = df_fichas['DestinoEcconomico'].fillna('').astype(str)

            # Validar longitud de 'NumCedulaCatastral' y DestinoEconomico
            for index, row in df_fichas.iterrows():
                num_cedula_catastral = row['NumCedulaCatastral'].strip()  # Convertir a cadena y quitar espacios

                # Validar longitud de 28 dígitos
                if len(num_cedula_catastral) != 28:
                    resultados.append({
                        'NroFicha': row['NroFicha'],
                        
                        'Observacion': 'NumCedulaCatastral no tiene 28 dígitos',
                        'Npn':row['Npn'],
                        'DestinoEconomico': row['DestinoEcconomico'],
                        'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                        'AreaTotalConstruida':row['AreaTotalConstruida'],
                        'CaracteristicaPredio':row['CaracteristicaPredio'],
                        'AreaTotalTerreno':row['AreaTotalTerreno'],
                        'ModoAdquisicion':row['ModoAdquisicion'],
                        'Tomo':row['Tomo'],
                        'PredioLcTipo':row['PredioLcTipo'],
                        'NumCedulaCatastral':row['NumCedulaCatastral'],
                        'AreaTotalLote':row['AreaTotalLote'],
                        'AreaLoteComun':row['AreaLoteComun'],
                        'AreaLotePrivada':row['AreaLotePrivada'],
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': 'Fichas'
                    })

                # Validar DestinoEconomico si el cuarto dígito de NumCedulaCatastral es '2'
                destino_economico = row['DestinoEconomico'].strip()
                if len(num_cedula_catastral) >= 4 and num_cedula_catastral[3] == '2' and destino_economico in destinos_invalidos:
                    resultados.append({
                        'NroFicha': row['NroFicha'],
                        'NumCedulaCatastral': num_cedula_catastral,
                        'DestinoEconomico': destino_economico,
                        'Observacion': 'Destino Económico no válido para ficha Rural',
                        'Nombre Hoja': 'Fichas'
                    })
            '''
            
            # Guardar los resultados en un archivo Excel si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_NumCedulaCatastral_y_DestinoEconomico_FichasPrediales.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros que cumplen con las condiciones.")
            else:
                messagebox.showinfo("Sin errores", "No se encontraron registros que cumplan con las condiciones en 'FichasPrediales'.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
    
    def validar_direccion_referencia_y_nombre(self):
        """
        Verifica que en la hoja 'FichasPrediales', las columnas 'DireccionReferencia' y 'DireccionNombre' no contengan valores nulos.
        """
        archivo_excel = self.archivo_entry.get()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'FichasPrediales'
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            resultados = []

            # Validar si hay valores nulos en las columnas 'DireccionReferencia' y 'DireccionNombre'
            for index, row in df_fichas.iterrows():
                if pd.isnull(row['DireccionReferencia']):
                    resultados.append({
                        'NroFicha': row['NroFicha'],
                        'DireccionReferencia': row['DireccionReferencia'],
                        'DireccionNombre':row['DireccionNombre'],
                        'Observacion': 'DireccionReferencia no está diligenciada',
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': 'FichasPrediales'
                    })

                if pd.isnull(row['DireccionNombre']):
                    resultados.append({
                        'NroFicha': row['NroFicha'],
                        'DireccionReferencia': row['DireccionReferencia'],
                        'DireccionNombre':row['DireccionNombre'],
                        'Observacion': 'DireccionNombre no está diligenciada',
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': 'FichasPrediales'
                    })
            '''
            
            # Guardar los resultados en un archivo Excel si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_DireccionReferencia_y_DireccionNombre_FichasPrediales.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros que cumplen con las condiciones.")
            else:
                messagebox.showinfo("Sin errores", "No se encontraron valores nulos en 'DireccionReferencia' o 'DireccionNombre' en 'FichasPrediales'.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        

    
        
    def Validar_Longitud_NPN(self):
        archivo_excel = self.archivo_entry.get()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'FichasPrediales'
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

           
            
            resultados = []

            # Validar longitud de 'NumCedulaCatastral'
            for index, row in df_fichas.iterrows():
                Npn = str(row['Npn']).strip()  # Convertir a cadena

                # Validar longitud de 28 dígitos
                if len(Npn) != 30:
                    resultados.append({
                        'NroFicha': row['NroFicha'],
                        'Npn': Npn,
                        'Observacion': 'Npn no tiene 30 dígitos',
                        'Nombre Hoja': 'Fichas'
                    })

                # Validar DestinoEconomico si el cuarto dígito de NumCedulaCatastral es '2'
                
            '''
            
            # Guardar los resultados en un archivo Excel si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_NumCedulaCatastral_y_DestinoEconomico_FichasPrediales.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros que cumplen con las condiciones.")
            else:
                messagebox.showinfo("Sin errores", "No se encontraron registros que cumplan con las condiciones en 'FichasPrediales'.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []   
    
    def validar_npn_sin_cuatro_ceros(self):
        """
        Verifica en la hoja 'Fichas' que las posiciones 14, 15, 16 y 17 del campo 'Npn' no sean '0000'.
        Si alguna fila contiene '0000' en esas posiciones, genera un error.
        """
        archivo_excel = self.archivo_entry.get()  # Obtener la ruta del archivo
        nombre_hoja = 'Fichas'

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'Fichas'
            df_fichas = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
            resultados = []

            # Validar que las posiciones 14-17 de 'Npn' no sean '0000'
            for index, row in df_fichas.iterrows():
                npn = str(row['Npn']).zfill(30)  # Convertir a cadena y rellenar para asegurar longitud
                if npn[13:17] == '0000':  # Comprobar posiciones 14, 15, 16 y 17
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'Observacion': 'NPN contiene "0000" en posiciones 14-17',
                        'Npn':row['Npn'],
                        'DestinoEconomico': row['DestinoEcconomico'],
                        'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                        'AreaTotalConstruida':row['AreaTotalConstruida'],
                        'CaracteristicaPredio':row['CaracteristicaPredio'],
                        'AreaTotalTerreno':row['AreaTotalTerreno'],
                        'ModoAdquisicion':row['ModoAdquisicion'],
                        'Tomo':row['Tomo'],
                        'PredioLcTipo':row['PredioLcTipo'],
                        'NumCedulaCatastral':row['NumCedulaCatastral'],
                        'AreaTotalLote':row['AreaTotalLote'],
                        'AreaLoteComun':row['AreaLoteComun'],
                        'AreaLotePrivada':row['AreaLotePrivada'],
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Error encontrado en la fila {index}: {resultado}")
            '''
            
            # Guardar los resultados en un archivo si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_Npn_Posiciones_14_17_Fichas.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros.")
            else:
                messagebox.showinfo("Sin errores", "No se encontraron registros con '0000' en posiciones 14-17 en 'Npn'.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
    
    def validar_Predios_Uso_Publico(self):
        """
        Valida en la hoja 'Fichas' que:
        1. Cuando el dígito 22 de 'Npn' es '3', 'CaracteristicaPredio' debe ser uno de los valores permitidos.
        2. Cuando el dígito 22 de 'Npn' es '2', los últimos 8 dígitos (posiciones 22 a 30) de 'Npn' deben ser '30000000'.
        """
        archivo_excel = self.archivo_entry.get()  # Obtener la ruta del archivo
        nombre_hoja = 'Fichas'

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'Fichas'
            df_fichas = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
            resultados = []

            # Definir las características permitidas cuando el dígito 22 de 'Npn' es '3'
            caracteristicas_permitidas = ['13|BIEN DE USO PUBLICO (3)', '6|EMBALSE', '11|VIA (4)','1|NPH (0)','9|SEPARADORES Y Z.V','8|LOTE']

            # Validar cada fila
            for index, row in df_fichas.iterrows():
                npn = str(row['Npn']).zfill(30)  # Convertir a cadena y rellenar para asegurar longitud
                caracteristica = str(row['CaracteristicaPredio'])

                # Validación 1: Cuando el dígito 22 es '3'
                if npn[21] == '3' and caracteristica not in caracteristicas_permitidas:
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'Observacion': 'CaracteristicaPredio inválida para Condicion 3',
                        'Npn':row['Npn'],
                        'DestinoEconomico': row['DestinoEcconomico'],
                        'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                        'AreaTotalConstruida':row['AreaTotalConstruida'],
                        'CaracteristicaPredio':row['CaracteristicaPredio'],
                        'AreaTotalTerreno':row['AreaTotalTerreno'],
                        'ModoAdquisicion':row['ModoAdquisicion'],
                        'Tomo':row['Tomo'],
                        'PredioLcTipo':row['PredioLcTipo'],
                        'NumCedulaCatastral':row['NumCedulaCatastral'],
                        'AreaTotalLote':row['AreaTotalLote'],
                        'AreaLoteComun':row['AreaLoteComun'],
                        'AreaLotePrivada':row['AreaLotePrivada'],
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Error encontrado en la fila {index}: {resultado}")

                # Validación 2: Cuando el dígito 22 es '3'
                if npn[21] == '3' and npn[21:30] != '300000000':
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'Npn': row['Npn'],
                        'Observacion': 'NPN debe terminar en 300000000 cuando el dígito 22 es 3',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Error encontrado en la fila {index}: {resultado}")
            '''
            
            # Guardar los resultados en un archivo si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_Npn_Caracteristica_Fichas.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros.")
            else:
                messagebox.showinfo("Sin errores", "No se encontraron errores en las validaciones de CaracteristicaPredio y Npn.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
    
    
    def validar_duplicados_npn(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_duplicados_npn")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Validar si la columna 'Npn' existe
            if 'Npn' not in df.columns:
                messagebox.showerror("Error", "La columna 'Npn' no existe en la hoja especificada.")
                return

            # Identificar duplicados en la columna 'Npn'
            duplicados = df[df.duplicated(subset=['Npn'], keep=False)]

            # Resultados a mostrar
            resultados = []
            for _, row in duplicados.iterrows():
                resultado = {
                    'NroFicha': row.get('NroFicha', 'No especificado'),
                    'Observacion': 'NPN duplicado',
                    'Npn':row['Npn'],
                    'DestinoEconomico': row['DestinoEcconomico'],
                    'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                    'AreaTotalConstruida':row['AreaTotalConstruida'],
                    'CaracteristicaPredio':row['CaracteristicaPredio'],
                    'AreaTotalTerreno':row['AreaTotalTerreno'],
                    'ModoAdquisicion':row['ModoAdquisicion'],
                    'Tomo':row['Tomo'],
                    'PredioLcTipo':row['PredioLcTipo'],
                    'NumCedulaCatastral':row['NumCedulaCatastral'],
                    'AreaTotalLote':row['AreaTotalLote'],
                    'AreaLoteComun':row['AreaLoteComun'],
                    'AreaLotePrivada':row['AreaLotePrivada'],
                    'Radicado':row['Radicado'],
                    'Nombre Hoja': nombre_hoja
                }
                resultados.append(resultado)
                print(f"Registro duplicado encontrado: {resultado}")
            '''
            # Verificar si hay duplicados y guardar resultados
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Duplicados_Npn.xlsx'
                sheet_name = 'Duplicados'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Se encontraron {len(resultados)} registros duplicados en la columna 'Npn'.")
            else:
                messagebox.showinfo("Información", "No se encontraron registros duplicados en la columna 'Npn'.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
    def validar_matricula_inmobiliaria_PredioLc_Modo_Adquisicion(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_matricula_inmobiliaria")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Validar si las columnas existen
            if 'PredioLcTipo' not in df.columns or 'ModoAdquisicion' not in df.columns or 'MatriculaInmobiliaria' not in df.columns:
                messagebox.showerror("Error", "Una o más columnas necesarias no existen en la hoja especificada.")
                return

            # Filtrar registros donde PredioLcTipo es 'Predio.Privado.Privado' y ModoAdquisicion es '1|DOMINIO (TRADICION)' y MatriculaInmobiliaria está vacío
            df['MatriculaInmobiliaria'] = df['MatriculaInmobiliaria'].fillna('')  # Rellenar valores nulos con cadena vacía
            registros_invalidos = df[
                (df['PredioLcTipo'] == 'Predio.Privado.Privado') &
                (df['ModoAdquisicion'] == '1|DOMINIO (TRADICION)') &
                (df['MatriculaInmobiliaria'] == '')
            ]

            # Resultados a mostrar
            resultados = []
            for _, row in registros_invalidos.iterrows():
                resultado = {
                    'NroFicha': row.get('NroFicha', 'No especificado'),
                    
                    'Observacion': 'Matricula vacía en predio privado y derecho dominio',
                    'Npn':row['Npn'],
                    'DestinoEconomico': row['DestinoEcconomico'],
                    'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                    'AreaTotalConstruida':row['AreaTotalConstruida'],
                    'CaracteristicaPredio':row['CaracteristicaPredio'],
                    'AreaTotalTerreno':row['AreaTotalTerreno'],
                    'ModoAdquisicion':row['ModoAdquisicion'],
                    'Tomo':row['Tomo'],
                    'PredioLcTipo':row['PredioLcTipo'],
                    'NumCedulaCatastral':row['NumCedulaCatastral'],
                    'AreaTotalLote':row['AreaTotalLote'],
                    'AreaLoteComun':row['AreaLoteComun'],
                    'AreaLotePrivada':row['AreaLotePrivada'],
                    'Radicado':row['Radicado'],
                    'Nombre Hoja': nombre_hoja
                }
                resultados.append(resultado)
                print(f"Condición de error encontrada: {resultado}")
            '''
            # Verificar si hay errores y guardar los resultados
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_MatriculaInmobiliaria.xlsx'
                sheet_name = 'ErroresMatriculaInmobiliaria'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Se encontraron {len(resultados)} registros con MatriculaInmobiliaria vacía.")
            else:
                messagebox.showinfo("Información", "No se encontraron registros con MatriculaInmobiliaria vacía cuando no debería serlo.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
    def validar_matricula_numerica(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_matricula_numerica")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Validar cada fila en la columna MatriculaInmobiliaria
            for _, row in df.iterrows():
                matricula = row.get('MatriculaInmobiliaria', '')
                
                if pd.notna(matricula):  # Validar que no sea NaN
                    if isinstance(matricula, float):  
                        # Convertir a entero si es un número decimal
                        matricula = str(int(matricula))
                    else:
                        # Convertir a cadena para verificar
                        matricula = str(matricula).strip()
                
                    # Verificar si no es completamente numérico
                    if not matricula.isdigit():  
                        resultado = {
                            'NroFicha': row.get('NroFicha', ''),
                            'Observacion': 'El campo MatriculaInmobiliaria contiene valores no numéricos',
                            'Npn':row['Npn'],
                            'DestinoEconomico': row['DestinoEcconomico'],
                            'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                            'AreaTotalConstruida':row['AreaTotalConstruida'],
                            'CaracteristicaPredio':row['CaracteristicaPredio'],
                            'AreaTotalTerreno':row['AreaTotalTerreno'],
                            'ModoAdquisicion':row['ModoAdquisicion'],
                            'Tomo':row['Tomo'],
                            'PredioLcTipo':row['PredioLcTipo'],
                            'NumCedulaCatastral':row['NumCedulaCatastral'],
                            'AreaTotalLote':row['AreaTotalLote'],
                            'AreaLoteComun':row['AreaLoteComun'],
                            'AreaLotePrivada':row['AreaLotePrivada'],
                            'Radicado':row['Radicado'],
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        print(f"Error encontrado: {resultado}")
            '''
            
            # Manejar resultados
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_Matricula_Numerica.xlsx'
                sheet_name = 'Errores Matricula'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Se encontraron {len(resultados)} registros con errores en MatriculaInmobiliaria.")
            else:
                messagebox.showinfo("Información", "Todos los valores en MatriculaInmobiliaria son numéricos o vacíos.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
    def validar_matricula_no_inicia_cero(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_matricula_no_inicia_cero")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Validar cada fila en la columna MatriculaInmobiliaria
            for _, row in df.iterrows():
                matricula = str(row.get('MatriculaInmobiliaria', '')).strip()
                if matricula and matricula.isdigit() and matricula.startswith('0'):  # Si es numérica y empieza con 0
                    resultado = {
                        'NroFicha': row.get('NroFicha', ''),
                    
                        'Observacion': 'El campo MatriculaInmobiliaria no debe iniciar con 0',
                        'Npn':row['Npn'],
                        'DestinoEconomico': row['DestinoEcconomico'],
                        'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                        'AreaTotalConstruida':row['AreaTotalConstruida'],
                        'CaracteristicaPredio':row['CaracteristicaPredio'],
                        'AreaTotalTerreno':row['AreaTotalTerreno'],
                        'ModoAdquisicion':row['ModoAdquisicion'],
                        'Tomo':row['Tomo'],
                        'PredioLcTipo':row['PredioLcTipo'],
                        'NumCedulaCatastral':row['NumCedulaCatastral'],
                        'AreaTotalLote':row['AreaTotalLote'],
                        'AreaLoteComun':row['AreaLoteComun'],
                        'AreaLotePrivada':row['AreaLotePrivada'],
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Error encontrado: {resultado}")
            '''
            
            # Manejar resultados
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_Matricula_Inicia_Cero.xlsx'
                sheet_name = 'Errores Matricula'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Se encontraron {len(resultados)} registros con MatriculaInmobiliaria que inicia con 0.")
            else:
                messagebox.showinfo("Información", "No se encontraron Matriculas Inmobiliarias que inicien con 0.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
    def validar_npn_modo_adquisicion(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_npn_modo_adquisicion")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre cada fila del DataFrame
            for _, row in df.iterrows():
                npn = str(row.get('Npn', '')).strip()
                modo_adquisicion = str(row.get('ModoAdquisicion', '')).strip()

                # Validar la longitud del Npn y el valor del dígito 22
                if len(npn) >= 22 and npn[21] == '2':
                    # Comprobar si el ModoAdquisicion no cumple las condiciones
                    if modo_adquisicion not in ['5|OCUPACIÓN', '2|POSESIÓN']:
                        resultado = {
                            'NroFicha': row.get('NroFicha', ''),
                            'Observacion': 'ModoAdquisicion inválido para Informalidad',
                            'Npn':row['Npn'],
                            'DestinoEconomico': row['DestinoEcconomico'],
                            'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                            'AreaTotalConstruida':row['AreaTotalConstruida'],
                            'CaracteristicaPredio':row['CaracteristicaPredio'],
                            'AreaTotalTerreno':row['AreaTotalTerreno'],
                            'ModoAdquisicion':row['ModoAdquisicion'],
                            'Tomo':row['Tomo'],
                            'PredioLcTipo':row['PredioLcTipo'],
                            'NumCedulaCatastral':row['NumCedulaCatastral'],
                            'AreaTotalLote':row['AreaTotalLote'],
                            'AreaLoteComun':row['AreaLoteComun'],
                            'AreaLotePrivada':row['AreaLotePrivada'], 
                            'Radicado':row['Radicado'],   
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        print(f"Error encontrado: {resultado}")
            '''
            
            # Manejar resultados
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_Npn_ModoAdquisicion.xlsx'
                sheet_name = 'Errores Npn ModoAdquisicion'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Se encontraron {len(resultados)} registros con errores en Npn y ModoAdquisicion.")
            else:
                messagebox.showinfo("Información", "No se encontraron errores en Npn y ModoAdquisicion.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
    
    
    def validar_matricula_repetida(self):
            archivo_excel = self.archivo_entry.get()
            nombre_hoja = 'Propietarios'

            if not archivo_excel or not nombre_hoja:
                messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
                return

            try:
                # Leer el archivo Excel
                df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

                print(f"funcion: validar_matricula_repetida")
                print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
                print(f"Dimensiones del DataFrame: {df.shape}")
                print(f"Columnas en el DataFrame: {df.columns.tolist()}")

                # Verificar que las columnas necesarias existan
                columnas_necesarias = ['MatriculaInmobiliaria', 'Documento']
                for columna in columnas_necesarias:
                    if columna not in df.columns:
                        messagebox.showerror("Error", f"La columna '{columna}' no existe en la hoja {nombre_hoja}.")
                        return
                df['MatriculaInmobiliaria'] = pd.to_numeric(df['MatriculaInmobiliaria'], errors='coerce')
                # Agrupar por 'MatriculaInmobiliaria' y contar ocurrencias
                duplicados = df.groupby('MatriculaInmobiliaria').filter(lambda x: len(x) > 1)

                # Lista para almacenar los errores encontrados
                errores = []

                # Validar si los duplicados tienen el mismo número de documento
                for matricula, grupo in duplicados.groupby('MatriculaInmobiliaria'):
                    documentos = grupo['Documento'].unique()
                    if len(documentos) == 1:
                        for _, fila in grupo.iterrows():
                            error = {
                                'MatriculaInmobiliaria': matricula,
                                'NroFicha': fila.get('NroFicha', 'N/A'),
                                'Observacion': 'Matricula inmobiliaria repetida',
                                'Documento': fila['Documento'],
                                'Radicado':fila['Radicado'],
                                'Nombre Hoja': nombre_hoja
                            }
                            errores.append(error)
                            print(f"Error encontrado: {error}")

                print(f"Total de errores encontrados: {len(errores)}")

                # Crear un DataFrame con los errores
                df_errores = pd.DataFrame(errores)

                '''
                # Guardar los errores en un archivo Excel
                output_file = 'Errores_MatriculaRepetida.xlsx'
                sheet_name = 'Errores'
                df_errores.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Proceso completado. Se encontró(n) {len(errores)} error(es).")
                '''
                return errores

            except Exception as e:
                print(f"Error: {str(e)}")
    '''
        
    def validar_destino_economico_nulo_o_0na(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return []

        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"función: validar_destino_economico_nulo_o_0na")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre cada fila del DataFrame
            for _, row in df.iterrows():
                destino_economico = row.get('DestinoEcconomico', None)  # Si no existe, es None

                # Validar si DestinoEcconomico es nulo o igual a '0|NA'
                if destino_economico is None or str(destino_economico).strip() == '0|NA':
                    resultado = {
                        'NroFicha': row.get('NroFicha', ''),  # Si no existe, devuelve vacío
                        'Observacion': 'DestinoEconómico nulo o igual a 0|NA',
                        'Npn':row['Npn'],
                        'DestinoEconomico': row['DestinoEcconomico'],
                        'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                        'AreaTotalConstruida':row['AreaTotalConstruida'],
                        'CaracteristicaPredio':row['CaracteristicaPredio'],
                        'AreaTotalTerreno':row['AreaTotalTerreno'],
                        'ModoAdquisicion':row['ModoAdquisicion'],
                        'Tomo':row['Tomo'],
                        'PredioLcTipo':row['PredioLcTipo'],
                        'NumCedulaCatastral':row['NumCedulaCatastral'],
                        'AreaTotalLote':row['AreaTotalLote'],
                        'AreaLoteComun':row['AreaLoteComun'],
                        'AreaLotePrivada':row['AreaLotePrivada'],
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Error encontrado: {resultado}")

            # Si se encontraron resultados, mostrarlos
            if resultados:
                for error in resultados:
                    print(error)
                messagebox.showinfo("Errores encontrados", f"Se encontraron {len(resultados)} errores en DestinoEconómico.")
            else:
                messagebox.showinfo("Información", "No se encontraron errores en DestinoEconómico.")
            
            # Retornar la lista de resultados
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
    '''
    def validar_destino_economico_nulo_o_0na(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return []

        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre cada fila del DataFrame
            for _, row in df.iterrows():
                destino_economico = row.get('DestinoEcconomico', None)  # Si no existe, es None
                npn = str(row.get('Npn', '')).strip()  # Convertir a string y limpiar espacios
                
                # Excluir las filas donde:
                # 1. condicion == 9 y unidad == 0000
                # 2. El dígito 22 de Npn es igual a 9 y los últimos 8 caracteres son ceros
                if not  (len(npn) >= 22 and (npn[21] in ['9', '8']) and npn[-8:] == '00000000'):
                # Validar si DestinoEconomico es NaN o igual a and npn[-8:] == '00000000'):
                    # Validar si DestinoEcconomico es NaN o igual a '0|NA'
                    if pd.isna(destino_economico) or str(destino_economico).strip() == '0|NA' :
                        resultado = {
                            'NroFicha': row.get('NroFicha', ''),  # Si no existe, devuelve vacío
                            'Observacion': 'Destino económico sin diligenciar',
                            'Npn':row['Npn'],
                            'DestinoEconomico': row['DestinoEcconomico'],
                            'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                            'AreaTotalConstruida':row['AreaTotalConstruida'],
                            'CaracteristicaPredio':row['CaracteristicaPredio'],
                            'AreaTotalTerreno':row['AreaTotalTerreno'],
                            'ModoAdquisicion':row['ModoAdquisicion'],
                            'Tomo':row['Tomo'],
                            'PredioLcTipo':row['PredioLcTipo'],
                            'NumCedulaCatastral':row['NumCedulaCatastral'],
                            'AreaTotalLote':row['AreaTotalLote'],
                            'AreaLoteComun':row['AreaLoteComun'],
                            'AreaLotePrivada':row['AreaLotePrivada'],
                            'Radicado':row['Radicado'],
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        print(f"Error encontrado: {resultado}")

            # Retornar la lista de resultados
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
    
    def validar_caracteristica_predio(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return []

        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre cada fila del DataFrame
            for _, row in df.iterrows():
                caracteristica_predio = row.get('CaracteristicaPredio', None)  # Si no existe, es None
                
                
                    # Validar si DestinoEcconomico es NaN o igual a '0|NA'
                if pd.isna(caracteristica_predio) or str(caracteristica_predio).strip() == '8|LOTE' :
                    resultado = {
                        'NroFicha': row.get('NroFicha', ''),  # Si no existe, devuelve vacío        
                        'Observacion': 'Caracteristica predio incorrecto',
                        'Npn':row['Npn'],
                        'DestinoEconomico': row['DestinoEcconomico'],
                        'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                        'AreaTotalConstruida':row['AreaTotalConstruida'],
                        'CaracteristicaPredio':row['CaracteristicaPredio'],
                        'AreaTotalTerreno':row['AreaTotalTerreno'],
                        'ModoAdquisicion':row['ModoAdquisicion'],
                        'Tomo':row['Tomo'],
                        'PredioLcTipo':row['PredioLcTipo'],
                        'NumCedulaCatastral':row['NumCedulaCatastral'],
                        'AreaTotalLote':row['AreaTotalLote'],
                        'AreaLoteComun':row['AreaLoteComun'],
                        'AreaLotePrivada':row['AreaLotePrivada'],
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Error encontrado: {resultado}")

            # Retornar la lista de resultados
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
        
    
    def validar_agricola_urb(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            # Leer el archivo Excel
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_cedula_destino")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Verificar que las columnas necesarias existan
            columnas_necesarias = ['NumCedulaCatastral', 'DestinoEcconomico']
            for columna in columnas_necesarias:
                if columna not in df.columns:
                    messagebox.showerror("Error", f"La columna '{columna}' no existe en la hoja {nombre_hoja}.")
                    return

            # Lista para almacenar los errores encontrados
            errores = []

            # Iterar sobre las filas para validar las condiciones
            for index, row in df.iterrows():
                num_cedula = str(row['NumCedulaCatastral'])  # Asegurarse de que sea una cadena
                destino = row['DestinoEcconomico']

                # Verificar si el cuarto dígito es '1' y el DestinoEcconomico es '24|AGRICOLA'
                if len(num_cedula) > 3 and num_cedula[3] == '1' and destino == '24|AGRICOLA':
                    error = {
                        'Observacion': 'DestinoEcconomico es 24|AGRICOLA en predio Urbano',
                        'Npn':row['Npn'],
                        'DestinoEconomico': row['DestinoEcconomico'],
                        'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                        'AreaTotalConstruida':row['AreaTotalConstruida'],
                        'CaracteristicaPredio':row['CaracteristicaPredio'],
                        'AreaTotalTerreno':row['AreaTotalTerreno'],
                        'ModoAdquisicion':row['ModoAdquisicion'],
                        'Tomo':row['Tomo'],
                        'PredioLcTipo':row['PredioLcTipo'],
                        'NumCedulaCatastral':row['NumCedulaCatastral'],
                        'AreaTotalLote':row['AreaTotalLote'],
                        'AreaLoteComun':row['AreaLoteComun'],
                        'AreaLotePrivada':row['AreaLotePrivada'],
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': nombre_hoja
                    }
                    errores.append(error)
                    print(f"Error encontrado en fila {index}: {error}")

            print(f"Total de errores encontrados: {len(errores)}")

            # Crear un DataFrame con los errores
            df_errores = pd.DataFrame(errores)

            '''
            # Guardar los errores en un archivo Excel
            output_file = 'Errores_CedulaDestino.xlsx'
            sheet_name = 'Errores'
            df_errores.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            messagebox.showinfo("Éxito", f"Proceso completado. Se encontró(n) {len(errores)} error(es).")
            '''
            return errores

        except Exception as e:
            print(f"Error: {str(e)}")