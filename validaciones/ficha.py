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

            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"funcion: terreno_cero")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                valor_b = row['NumCedulaCatastral']
                valor_p = row['AreaTotalTerreno']

                # Verificar si valor_b no es nulo o vacío, y si tiene al menos 22 caracteres
                if pd.notna(valor_b) and len(str(valor_b)) > 21:
                    valor_b_str = str(valor_b)  # Convertir el valor a cadena
                    print(f"Fila {index}: Valor B = '{valor_b_str}', condicion: {valor_b_str[21]}, Valor P = '{valor_p}'")

                    # Verificar las condiciones
                    if valor_b_str[21] == '0' and (valor_p == '0' or valor_p == 0):
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'Observacion': 'Terreno en ceros para ficha que no es mejora'
                        }
                        resultados.append(resultado)
                        print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
                else:
                    print(f"Fila {index}: NumCedulaCatastral no tiene suficientes caracteres o es nulo.")

            print(f"Total de resultados encontrados: {len(resultados)}")

            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)
            
            messagebox.showinfo("Éxito", f"Proceso completado Terreno cero. con {len(resultados)} registros.")
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
                valor_a = str(row['NumCedulaCatastral'])  # Convertir a cadena por si acaso
                valor_b = row['AreaTotalTerreno']

            print(f"Fila {index}: Valor A = '{valor_a}'")

            # Verificar que valor_a tenga al menos 22 caracteres
            if len(valor_a) > 21:
                print(f"CARACTER22B : {valor_a[21]}")

                if valor_a[21] == '2' and pd.isna(valor_b):
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'NumCedulaCatastral': row['NumCedulaCatastral'],
                        'Condicion de predio': valor_a[21],
                        'AreaTotalTerreno': valor_b,
                        'Observacion': 'Terreno nulo para condición de predio'
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

            '''
            
            messagebox.showinfo("Éxito",
                                f"Proceso completado Terreno null. con {len(resultados)} registros.")
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

            print(f"funcion: informal_matricula")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                valor_a = str(row['MatriculaInmobiliaria']) if pd.notna(row['MatriculaInmobiliaria']) else ''
                valor_b = row['ModoAdquisicion']

                print(f"Fila {index}: Valor A = '{valor_a}', condicion: {valor_b}")

                # Verificar las condiciones: valor_b es '2|POSESIÓN' y valor_a no está vacío
                if valor_b == '2|POSESIÓN' and (valor_a != '' or pd.notna(valor_a)):
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'ModoAdquisicion': row['ModoAdquisicion'],
                        'MatriculaInmobiliaria': valor_a,
                        'Observacion': 'Matricula invalida para posesión'
                         
                    }
                    resultados.append(resultado)
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

            print(f"Total de resultados encontrados: {len(resultados)}")
            
            
       
            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)

            # Guardar el resultado en un nuevo archivo Excel
            output_file = 'INFORMA_MATRICULA.xlsx'
            sheet_name = 'INFORMAL_MATRICULA'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            messagebox.showinfo("Éxito", f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
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
                valor_a = str(row['NumCedulaCatastral']) if pd.notna(row['NumCedulaCatastral']) else ''
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
                            'NumCedulaCatastral': valor_a,
                            'Condicion de predio': valor_a[21],
                            'ModoAdquisicion': row['ModoAdquisicion'],
                            'MatriculaInmobiliaria': row['MatriculaInmobiliaria'],
                            'Observacion': 'Informalidad con matrícula'
                            
                        }
                        
                        resultados.append(resultado)
                        print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
                else:
                    print(f"Fila {index} tiene una NumCedulaCatastral demasiado corta: '{valor_a}'")

            print(f"Total de resultados encontrados: {len(resultados)}")

            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)

            # Guardar el resultado en un nuevo archivo Excel
            output_file = 'MATRICULA_MEJORA.xlsx'
            sheet_name = 'MATRICULA_MEJORA'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            messagebox.showinfo("Éxito",
                                f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
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
                valor_a = row['NumCedulaCatastral']
                valor_b = row['circulo']

            print(f"Fila {index}: Valor A = '{valor_a}'")

            # Verificar que 'valor_a' tiene al menos 22 caracteres
            if len(valor_a) > 21:
                print(f"CARACTER22B : {valor_a[21]}")

                # Verificar las condiciones
                if valor_a[21] == '2' and pd.notna(valor_b):
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'NumCedulaCatastral': row['NumCedulaCatastral'],
                        'Condicion de predio': valor_a[21],
                        'ModoAdquisicion': row['ModoAdquisicion'],
                        'circulo': row['circulo'],
                        'Observacion': 'Informalidad con matrícula'
                    }
                    resultados.append(resultado)
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
            else:
                print(f"El valor de 'NumCedulaCatastral' en la fila {index} no tiene suficientes caracteres.")

            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)

            # Guardar el resultado en un nuevo archivo Excel
            output_file = 'CIRCULO_MEJORA.xlsx'
            sheet_name = 'CIRCULO_MEJORA'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            messagebox.showinfo("Éxito",
                                f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
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
                valor_a = row['NumCedulaCatastral']
                valor_b = row['Tomo']

                print(f"Fila {index}: Valor A = '{valor_a}'")

                # Verificar que 'valor_a' tiene al menos 22 caracteres antes de acceder al índice 21
                if len(valor_a) > 21:
                    # Verificar las condiciones
                    if valor_a[21] == '2' and valor_b != 0:
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'NumCedulaCatastral': row['NumCedulaCatastral'],
                            'Condicion de predio': valor_a[21],
                            'ModoAdquisicion': row['ModoAdquisicion'],
                            'Tomo': row['Tomo'],
                            'Observacion': 'Informalidad con Tomo'
                        }
                        resultados.append(resultado)
                        print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
                else:
                    print(f"El valor de 'NumCedulaCatastral' en la fila {index} no tiene suficientes caracteres.")

            print(f"Total de resultados encontrados: {len(resultados)}")

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
                valor_a = row['NumCedulaCatastral']
                valor_b = row['ModoAdquisicion']

                print(f"Fila {index}: Valor A = '{valor_a}'")

                # Verificar que 'valor_a' tiene al menos 22 caracteres antes de acceder al índice 21
                if len(valor_a) > 21:
                    # Verificar las condiciones
                    if valor_a[21] == '2' and valor_b != '2|POSESIÓN':
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'NumCedulaCatastral': row['NumCedulaCatastral'],
                            'Condicion de predio': valor_a[21],
                            'ModoAdquisicion': row['ModoAdquisicion'],
                            'Observacion': 'La informalidad no puede tener modo de adquisición diferente a posesión'
                        }
                        resultados.append(resultado)
                        print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
                else:
                    print(f"El valor de 'NumCedulaCatastral' en la fila {index} no tiene suficientes caracteres.")

            print(f"Total de resultados encontrados: {len(resultados)}")

            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)

            # Guardar el resultado en un nuevo archivo Excel
            output_file = 'MODO_ADQUISICION_INFORMAL.xlsx'
            sheet_name = 'MODO_ADQUISICION_INFORMAL'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            messagebox.showinfo("Éxito",
                                f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
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

            if duplicados.shape[0] > 0:
                # Guardar los resultados en un nuevo archivo Excel
                output_file = 'FICHAS_REPETIDAS.xlsx'
                sheet_name = 'fichas_repetidas'
                duplicados.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                print(f"Dimensiones del DataFrame de duplicados: {duplicados.shape}")

                messagebox.showinfo("Éxito", f"Proceso completado. Se ha creado el archivo '{output_file}' con {duplicados.shape[0]} registros duplicados.")
                return duplicados
            else:
                print("No se encontraron registros duplicados.")
                messagebox.showinfo("Información", "No se encontraron registros duplicados.")
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
                        'DestinoEcconomico': row['DestinoEcconomico'],
                        'Observacion': 'En sector rural no es valido destinaciones 12,13 y 14'
                    }
                    resultados.append(resultado)
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

            print(f"Total de resultados encontrados: {len(resultados)}")

            # Crear un nuevo DataFrame con los resultados

            df_resultado = pd.DataFrame(resultados)
            # Guardar el resultado en un nuevo archivo Excel
            output_file = 'rural_destino_invalido.xlsx'
            sheet_name = 'rural_destino_invalido'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            messagebox.showinfo("Éxito",
                                f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
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
