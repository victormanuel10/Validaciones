import pandas as pd
from tkinter import messagebox

class FichasRPH:
    def __init__(self, archivo_entry):
        # archivo_entry can be either a string (file path) or tkinter.Entry
        self.archivo_entry = archivo_entry
        self.resultados_generales = []

    def obtener_archivo(self):
        """ Helper function to get the file path from either a tkinter.Entry or a string. """
        if isinstance(self.archivo_entry, str):
            return self.archivo_entry
        elif hasattr(self.archivo_entry, 'get'):
            return self.archivo_entry.get()
        else:
            return None

    def validar_coeficiente_copropiedad_por_npn(self):
        """ 
        Valida que la suma de CoeficienteCopropiedad para los primeros 22 dígitos de Npn 
        sea igual a 100 en la hoja 'Fichas Prediales'. Además, verifica que para registros 
        donde el 22° dígito sea '8' o '9' y los últimos 4 dígitos no sean '0000', 
        el CoeficienteCopropiedad sea mayor a 0.
        """
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        xls = pd.ExcelFile(archivo_excel)  # Carga el archivo Excel
        print(xls.sheet_names)  # Lista todas las hojas disponibles

        if 'Construcciones' not in xls.sheet_names:
            print("Error", "La hoja 'Construcciones' no existe en el archivo Excel.")
            return

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja específica "Fichas Prediales"
            df_fichas = pd.read_excel(archivo_excel, sheet_name=nombre_hoja, dtype={'Npn': str})

            # Crear una columna 'Npn_22' con los primeros 22 caracteres de Npn
            df_fichas['Npn_22'] = df_fichas['Npn'].str[:22]

            # Filtrar registros donde el 22° dígito de 'Npn' sea '8' o '9'
            df_filtrado = df_fichas[df_fichas['Npn'].str[21].isin(['8', '9'])]

            # Verificar que los últimos 4 dígitos sean diferentes de '0000'
            df_filtrado = df_filtrado[df_filtrado['Npn'].str[-4:] != '0000']

            # Validar que el CoeficienteCopropiedad sea mayor a 0, si no, generar error
            errores_coeficiente_cero = df_filtrado[df_filtrado['CoeficienteCopropiedad'] <= 0]

            resultados = []

            for _, fila in errores_coeficiente_cero.iterrows():
                resultado = {
                    'Npn': fila['Npn'],
                    'Radicado': fila.get('Radicado', ''),  # Usar get para evitar error si la columna no existe
                    'Suma CoeficienteCopropiedad': fila['CoeficienteCopropiedad'],
                    'Observacion': 'El CoeficienteCopropiedad debe ser mayor a 0',
                    'Nombre Hoja': 'FichasPrediales',
                    'NroFicha': fila['NroFicha']
                }
                resultados.append(resultado)
                print(f"Error agregado: {resultado}")

            # Agrupar por 'Npn_22' y sumar 'CoeficienteCopropiedad'
            suma_coeficientes = df_filtrado.groupby('Npn_22')['CoeficienteCopropiedad'].sum().reset_index()

            # Ajustar los coeficientes que estén dentro del margen de 0.01 de 100 a 100
            suma_coeficientes['CoeficienteCopropiedad'] = suma_coeficientes['CoeficienteCopropiedad'].apply(
                lambda x: 100 if abs(x - 100) <= 0.01 else round(x, 3)
            )

            # Filtrar donde la suma no es 100
            errores_suma = suma_coeficientes[suma_coeficientes['CoeficienteCopropiedad'] != 100]

            for _, row in errores_suma.iterrows():
                npn_22 = row['Npn_22']
                coeficiente_suma = row['CoeficienteCopropiedad']

                # Obtener todas las filas que corresponden al 'Npn_22' para extraer también 'Radicado'
                filas_error = df_fichas[df_fichas['Npn_22'] == npn_22]

                for _, fila in filas_error.iterrows():
                    resultado = {
                        'Npn': fila['Npn'],
                        'Radicado': fila.get('Radicado', ''),
                        'Suma CoeficienteCopropiedad': coeficiente_suma,
                        'Observacion': 'La suma de CoeficienteCopropiedad no es 100',
                        'Nombre Hoja': 'FichasPrediales',
                        'NroFicha': fila['NroFicha']
                    }
                    resultados.append(resultado)
                    print(f"Error agregado: {resultado}")

            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
        
    def ficha_resumen_sin_unidades(self):
        """
        Verifica que en la hoja 'FichasPrediales' existan duplicados en los primeros 22 caracteres de Npn
        solo si el dígito 22 de 'Npn' es '8' o '9'. Si no hay duplicados, genera un error.
        """
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido y especifica la hoja.")
            return []

        try:
            # Leer la hoja específica 'FichasPrediales'
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            # Crear una columna 'Npn_22' con los primeros 22 caracteres de Npn
            df['Npn'] = df['Npn'].astype(str).str.zfill(24)  # Asegurarse de que todos los Npn tengan al menos 24 caracteres
            df['Npn_22'] = df['Npn'].str[:22]

            # Filtrar los registros donde el dígito 22 de 'Npn' es '8' o '9'
            df_filtrado = df[df['Npn'].str[21].isin(['8', '9'])]

            # Contar ocurrencias de cada valor en 'Npn_22' dentro del DataFrame filtrado
            conteo_npn = df_filtrado['Npn_22'].value_counts()

            # Filtrar los valores de 'Npn_22' que no tienen duplicados
            sin_duplicados = conteo_npn[conteo_npn == 1].index.tolist()

            # Generar una lista de errores para los registros sin duplicados
            resultados = []
            if sin_duplicados:
                for npn_22 in sin_duplicados:
                    filas_error = df_filtrado[df_filtrado['Npn_22'] == npn_22]
                    for _, fila in filas_error.iterrows():
                        resultado = {
                            'NroFicha': fila['NroFicha'],
                            'Npn': fila['Npn'],
                            'Radicado': fila['Radicado'],  # Agregar columna Radicado
                            'Observacion': 'No existe Unidades prediales para la ficha resumen',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        print(f"Error agregado: {resultado}")

            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
    def edificio_en_cero_rph(self):
        """
        Valida que en la columna 'Npn' de la hoja 'FichasPrediales', el dígito 22 sea 8 o 9.
        Solo un registro en cada grupo de Npn con los mismos primeros 22 dígitos puede tener '00' en los dígitos 23 y 24.
        Si hay más de un registro con '00' en esos dígitos, se genera un error.
        """
        archivo_excel = self.archivo_entry.get()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja específica 'FichasPrediales'
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')
            
            resultados = []

            # Asegurarse de que los valores en 'Npn' sean cadenas de texto y llenar valores nulos
            df_fichas['Npn'] = df_fichas['Npn'].fillna('').astype(str)

            # Agrupar registros por los primeros 22 dígitos de 'Npn'
            df_fichas['Npn_22_digitos'] = df_fichas['Npn'].str[:22]
            grupos_npn = df_fichas.groupby('Npn_22_digitos')

            for npn_22, grupo in grupos_npn:
                # Convertir los valores de 'Npn' en el grupo a cadenas de 24 caracteres
                grupo['Npn_24_digitos'] = grupo['Npn'].apply(lambda x: x.zfill(24))
                
                # Filtrar registros con dígito 22 igual a 8 o 9
                grupo = grupo[grupo['Npn_24_digitos'].str[21] == '9']

                # Separar los registros donde los dígitos 23 y 24 son '00'
                npn_con_cero = grupo[grupo['Npn_24_digitos'].str[22:24] == '00']

                # Si hay más de un registro con '00' en los dígitos 23 y 24 en el mismo grupo, genera error
                if len(npn_con_cero) > 1:
                    for _, row in npn_con_cero.iterrows():
                        npn = row['Npn']
                        digitos_27_30 = npn[26:30]
                        if digitos_27_30.isdigit() and sum(int(d) for d in digitos_27_30) > 0:
                            resultado = {
                                'NroFicha': row['NroFicha'],
                                'Npn': npn,
                                'Observacion': 'Edificio no puede ser 00 en RPH',
                                'Radicado':row['Radicado'],
                                'Nombre Hoja': 'FichasPrediales'
                            }
                            resultados.append(resultado)

                # Para los otros registros en el grupo, verificar que la suma de los dígitos 23 y 24 sea mayor o igual a 1
                npn_no_cero = grupo[grupo['Npn_24_digitos'].str[22:24] != '00']
                for _, row in npn_no_cero.iterrows():
                    ultimos_dos_digitos = row['Npn_24_digitos'][22:24]
                    suma_digitos = sum(int(d) for d in ultimos_dos_digitos if d.isdigit())
                    if suma_digitos < 1:
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'Npn': row['Npn'],
                            'Observacion': 'Edificio en 0 para condición de predio 9',
                            'Radicado':row['Radicado'],
                            'Nombre Hoja': 'FichasPrediales'
                        }
                        resultados.append(resultado)

            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []

    
    def piso_en_cero_rph(self):
        """
        Valida que en la columna 'Npn' de la hoja 'FichasPrediales', el dígito 22 sea 8 o 9.
        Solo un registro en cada grupo de Npn con los mismos primeros 22 dígitos puede tener '00' en los dígitos 23 y 24.
        Si hay más de un registro con '00' en esos dígitos, se genera un error.
        """
        archivo_excel = self.archivo_entry.get()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja específica 'FichasPrediales'
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')
            
            resultados = []

            # Asegurarse de que los valores en 'Npn' sean cadenas de texto y llenar valores nulos
            df_fichas['Npn'] = df_fichas['Npn'].fillna('').astype(str)

            # Agrupar registros por los primeros 22 dígitos de 'Npn'
            df_fichas['Npn_22_digitos'] = df_fichas['Npn'].str[:22]
            grupos_npn = df_fichas.groupby('Npn_22_digitos')

            for npn_22, grupo in grupos_npn:
                # Convertir los valores de 'Npn' en el grupo a cadenas de 24 caracteres
                grupo['Npn_24_digitos'] = grupo['Npn'].apply(lambda x: x.zfill(24))
                
                # Filtrar registros con dígito 22 igual a 8 o 9
                grupo = grupo[grupo['Npn_24_digitos'].str[21] == '9']

                # Separar los registros donde los dígitos 23 y 24 son '00'
                npn_con_cero = grupo[grupo['Npn_24_digitos'].str[24:26] == '00']

                # Si hay más de un registro con '00' en los dígitos 23 y 24 en el mismo grupo, genera error
                if len(npn_con_cero) > 1:
                    for _, row in npn_con_cero.iterrows():
                        npn = row['Npn']
                        digitos_27_30 = npn[26:30]
                        if digitos_27_30.isdigit() and sum(int(d) for d in digitos_27_30) > 0:
                            resultado = {
                                'NroFicha': row['NroFicha'],
                                'Npn': npn,
                                'Observacion': 'Piso no puede ser 00 en RPH',
                                'Radicado':row['Radicado'],
                                'Nombre Hoja': 'FichasPrediales'
                            }
                            resultados.append(resultado)

                # Para los otros registros en el grupo, verificar que la suma de los dígitos 23 y 24 sea mayor o igual a 1
                npn_no_cero = grupo[grupo['Npn_24_digitos'].str[24:26] != '00']
                for _, row in npn_no_cero.iterrows():
                    ultimos_dos_digitos = row['Npn_24_digitos'][24:26]
                    suma_digitos = sum(int(d) for d in ultimos_dos_digitos if d.isdigit())
                    if suma_digitos < 1:
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'Npn': row['Npn'],
                            'Observacion': 'Piso en 0 para condición de predio 9',
                            'Radicado':row['Radicado'],
                            'Nombre Hoja': 'FichasPrediales'
                        }
                        resultados.append(resultado)

            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
    
    def validar_npn_suma_cero_unico(self):
        """
        Valida en la hoja 'Fichas' que:
        - Si el 22.º dígito de 'Npn' es '8' o '9', solo un registro puede tener suma cero en sus últimos 4 dígitos.
        - Si existen otros registros con los mismos primeros 22 dígitos y con suma cero en sus últimos 4 dígitos, se genera un error.
        """
        archivo_excel = self.archivo_entry.get()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'Fichas'
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')
            
            resultados = []

            # Asegurarnos de que los valores en 'Npn' sean cadenas de texto y llenar valores nulos
            df_fichas['Npn'] = df_fichas['Npn'].fillna('').astype(str)

            # Filtrar los registros donde el 22.º dígito de 'Npn' es '8' o '9'
            df_fichas = df_fichas[df_fichas['Npn'].str.len() >= 22]  # Asegura que tengan al menos 22 dígitos
            df_fichas = df_fichas[df_fichas['Npn'].str[21].isin(['8', '9'])]
            
            # Agrupar por los primeros 22 dígitos de 'Npn' para encontrar duplicados
            df_fichas['Npn_22_digitos'] = df_fichas['Npn'].str[:22]
            grupos_npn = df_fichas.groupby('Npn_22_digitos')

            for npn_22, grupo in grupos_npn:
                # Convertir los últimos 4 dígitos a enteros y calcular la suma
                def calcular_suma_ultimos_4(npn):
                    ultimos_4 = ''.join(filter(str.isdigit, str(npn)[-4:]))  # Extrae los últimos 4 dígitos numéricos
                    if len(ultimos_4) == 4:
                        return sum(int(d) for d in ultimos_4)
                    else:
                        return None  # Retorna None si no se tienen exactamente 4 dígitos numéricos al final

                grupo['Suma_ultimos_4'] = grupo['Npn'].apply(calcular_suma_ultimos_4)

                # Filtrar registros con suma cero en los últimos cuatro dígitos
                suma_cero = grupo[grupo['Suma_ultimos_4'] == 0]

                # Verificar si hay más de un registro con suma cero en los últimos 4 dígitos
                if len(suma_cero) > 1:
                    for _, row in suma_cero.iterrows():
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'Npn': row['Npn'],
                            'Observacion': 'Ficha Resumen sin Unidades Prediales',
                            'Radicado':row['Radicado'],
                            'Radicado':row['Radicado'],
                            'Nombre Hoja': 'Fichas'
                        }
                        resultados.append(resultado)
            '''
            # Guardar los errores en un archivo Excel si existen
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_Npn_SumaCero_Unico_Fichas.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Errores encontrados", f"Se encontraron {len(resultados)} errores en la validación de suma cero única para 'Npn'.")
            else:
                messagebox.showinfo("Validación completada", "No se encontraron errores en la validación de suma cero única para 'Npn'.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            
            #messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
    
    
    def validar_npn_y_caracteristica(self):
        """
        Valida en la hoja 'Fichas' los registros donde:
        - El 22.º dígito de 'Npn' es '9'.
        - Los primeros 22 dígitos de 'Npn' están duplicados.
        - La suma de los últimos cuatro dígitos (27.º a 30.º) es mayor a cero.
        - 'CaracteristicaPredio' debe coincidir con la del registro con suma cero.
        """
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'Fichas'
            df_fichas = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
            
            resultados = []

            # Filtrar los registros donde el 22.º dígito de 'Npn' es '9'
            df_fichas = df_fichas[df_fichas['Npn'].astype(str).str[21] == '9']
            
            # Agrupar por los primeros 22 dígitos de 'Npn' para encontrar duplicados
            df_fichas['Npn_22_digitos'] = df_fichas['Npn'].astype(str).str[:22]
            grupos_npn = df_fichas.groupby('Npn_22_digitos')

            for npn_22, grupo in grupos_npn:
                # Convertir los últimos 4 dígitos a enteros y calcular la suma
                grupo['Ultimos_4_digitos'] = grupo['Npn'].astype(str).str[-4:].astype(int)
                grupo['Suma_ultimos_4'] = grupo['Ultimos_4_digitos'].apply(lambda x: sum(int(d) for d in str(x)))

                # Identificar el registro con suma igual a cero, si existe
                referencia = grupo[grupo['Suma_ultimos_4'] == 0]
                if not referencia.empty:
                    caracteristica_referencia = referencia.iloc[0]['CaracteristicaPredio']

                    # Validar los otros registros en el grupo
                    for _, row in grupo.iterrows():
                        if row['Suma_ultimos_4'] > 0 and row['CaracteristicaPredio'] != caracteristica_referencia:
                            resultado = {
                                'NroFicha': row['NroFicha'],
                                'Npn': row['Npn'],
                                'CaracteristicaPredio': row['CaracteristicaPredio'],
                                'CaracteristicaPredioEsperada': caracteristica_referencia,
                                'Observacion': 'CaracteristicaPredio no coincide con la ficha resumen',
                                'Radicado':row['Radicado'],
                                'Nombre Hoja': 'Fichas'
                            }
                            resultados.append(resultado)
            '''
            
            # Guardar los errores en un archivo Excel si existen
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_Npn_CaracteristicaPredio_Fichas.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Errores encontrados", f"Se encontraron {len(resultados)} errores en la validación de 'Npn' y 'CaracteristicaPredio'.")
            else:
                messagebox.showinfo("Validación completada", "No se encontraron errores en 'Npn' y 'CaracteristicaPredio'.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
    def validar_npn_num_cedula(self):
        """
        Valida en la hoja 'Fichas' los registros donde:
        - El 22.º dígito de 'Npn' es '9'.
        - Los primeros 22 dígitos de 'Npn' están duplicados.
        - La suma de los últimos cuatro dígitos (27.º a 30.º) es mayor a cero.
        - 'NumCedulaCatastral' de los registros con suma mayor a cero no debe coincidir con el del registro con suma cero.
        """
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'Fichas'
            df_fichas = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
            
            resultados = []

            # Filtrar los registros donde el 22.º dígito de 'Npn' es '9'
            df_fichas = df_fichas[df_fichas['Npn'].astype(str).str[21] == '9']
            
            # Agrupar por los primeros 22 dígitos de 'Npn' para encontrar duplicados
            df_fichas['Npn_22_digitos'] = df_fichas['Npn'].astype(str).str[:22]
            grupos_npn = df_fichas.groupby('Npn_22_digitos')

            for npn_22, grupo in grupos_npn:
                # Convertir los últimos 4 dígitos a enteros y calcular la suma
                grupo['Ultimos_4_digitos'] = grupo['Npn'].astype(str).str[-4:].astype(int)
                grupo['Suma_ultimos_4'] = grupo['Ultimos_4_digitos'].apply(lambda x: sum(int(d) for d in str(x)))

                # Identificar el registro con suma igual a cero, si existe
                referencia = grupo[grupo['Suma_ultimos_4'] == 0]
                if not referencia.empty:
                    referencia_cedula = referencia.iloc[0]['NumCedulaCatastral']
                    
                    # Validar los otros registros en el grupo
                    for _, row in grupo.iterrows():
                        if row['Suma_ultimos_4'] > 0 and row['NumCedulaCatastral'] == referencia_cedula:
                            resultado = {
                                'NroFicha': row['NroFicha'],
                                'Npn': row['Npn'],
                                'NumCedulaCatastral': row['NumCedulaCatastral'],
                                'Observacion': 'NumCedulaCatastral ya existe en ficha resumen',
                                'Radicado':row['Radicado'],
                                'Nombre Hoja': 'Fichas'
                            }
                            resultados.append(resultado)
            '''
            
            # Guardar los errores en un archivo Excel si existen
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_NumCedulaCatastral_Npn_Fichas.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Errores encontrados", f"Se encontraron {len(resultados)} errores en la validación de 'NumCedulaCatastral' y 'Npn'.")
            else:
                messagebox.showinfo("Validación completada", "No se encontraron errores en 'NumCedulaCatastral' y 'Npn'.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
    def validar_area_total_lote_npn(self):
        
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'Fichas'
            df_fichas = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            resultados = []

            # Iterar sobre cada fila para verificar las condiciones
            for index, row in df_fichas.iterrows():
                npn = str(row.get('Npn', '')).strip()  # Convertir a cadena y quitar espacios
                area_total_lote = row.get('AreaTotalLote', None)

                # Verificar si el 22º dígito de 'Npn' es '9' y la suma de los dígitos 27 a 30 es 0
                if len(npn) >= 30 and npn[21] == '9':
                    digitos_27_30 = npn[26:30]  # Obtener los dígitos 27, 28, 29, 30
                    
                    # Verificar que los dígitos son números y sumarlos
                    if digitos_27_30.isdigit() and sum(int(d) for d in digitos_27_30) == 0:
                        # Generar error si 'AreaTotalLote' está vacío
                        if pd.isna(area_total_lote) or area_total_lote == '' or area_total_lote==0:
                            resultado = {
                                'NroFicha': row['NroFicha'],
                                'AreaTotalLote':row['AreaTotalLote'],
                                'Npn': npn,
                                'Observacion': 'AreaTotalLote no debe ser CERO o VACIO en ficha resumen',
                                'Radicado':row['Radicado'],
                                'Nombre Hoja': nombre_hoja
                            }
                            resultados.append(resultado)
                            print(f"Fila {index} cumple las condiciones para error. Agregado: {resultado}")
            '''
            
            # Guardar los resultados en un archivo Excel si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_AreaTotalLote_Npn_Fichas.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros con Npn cuyo 22º dígito es 9 y sin AreaTotalLote.")
            else:
                messagebox.showinfo("Sin errores", "Todos los registros cumplen con las condiciones o tienen AreaTotalLote lleno.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        

    
    '''
    
    def validar_area_privada(self):
        """
        Verifica en la hoja 'Fichas' que cuando el 22º dígito de 'Npn' es '9' y la suma de los
        dígitos 27, 28, 29 y 30 es 0, el campo 'AreaTotalLote' no esté vacío. Si está vacío, genera un error.
        """
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'
        
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'Fichas'
            df_fichas = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            resultados = []

            # Iterar sobre cada fila para verificar las condiciones
            for index, row in df_fichas.iterrows():
                npn = str(row.get('Npn', '')).strip()  # Convertir a cadena y quitar espacios
                arealoteprivada = row.get('AreaLotePrivada', None)

                # Verificar si el 22º dígito de 'Npn' es '9' y la suma de los dígitos 27 a 30 es 0
                if len(npn) >= 30 and npn[21] == '9':
                    digitos_27_30 = npn[26:30]  # Obtener los dígitos 27, 28, 29, 30
                    
                    # Verificar que los dígitos son números y sumarlos
                    if digitos_27_30.isdigit() and sum(int(d) for d in digitos_27_30) > 0:
                        # Generar error si 'AreaTotalLote' está vacío
                        if pd.isna(arealoteprivada) or arealoteprivada == '' or arealoteprivada==0:
                            resultado = {
                                'NroFicha': row['NroFicha'],
                                'AreaLoteComun':row['AreaLoteComun'],
                                'AreaLotePrivada':row['AreaLotePrivada'],
                                'Npn': npn,
                                'Observacion': 'AreaLotePrivada no debe estar vacío en Unidad Predial',
                                'Nombre Hoja': nombre_hoja
                            }
                            resultados.append(resultado)
                            print(f"Fila {index} cumple las condiciones para error. Agregado: {resultado}")
            
            
            # Guardar los resultados en un archivo Excel si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_AreaLoteComun_Npn_Fichas.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros con Npn cuyo 22º dígito es 9 y sin AreaTotalLote.")
            else:
                messagebox.showinfo("Sin errores", "Todos los registros cumplen con las condiciones o tienen AreaTotalLote lleno.")
            
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        '''
    def validar_unidades_rph(self):
        """
        Valida que para los registros donde el dígito 22 en 'Npn' es '9' y la suma de los últimos 4 dígitos es cero,
        el valor de 'UnidadesEnRPH' sea igual a la cantidad de otros 'Npn' con los mismos primeros 22 caracteres,
        y la suma de los últimos 4 dígitos de estos otros 'Npn' sea mayor a cero.
        """
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'Fichas'
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
            
            resultados = []
            
            # Filtrar los registros donde el dígito 22 es '9' y la suma de los últimos 4 dígitos es cero
            for index, row in df.iterrows():
                npn = str(row['Npn']).zfill(30)  # Asegura longitud de 30 caracteres en Npn
                if len(npn) >= 30 and npn[21] == '9' and sum(int(digit) for digit in npn[26:30]) == 0:
                    # Obtener los primeros 22 dígitos del Npn actual
                    npn_22 = npn[:22]
                    
                    # Contar otros Npn que tienen los mismos primeros 22 caracteres y suma mayor a cero en los últimos 4 dígitos
                    conteo_npn_relacionados = df[(df['Npn'].astype(str).str[:22] == npn_22) &
                                                (df['Npn'].astype(str).apply(lambda x: sum(int(d) for d in str(x)[26:30]) > 0))].shape[0]
                    
                    # Verificar si el conteo es igual al valor de 'UnidadesEnRPH'
                    if conteo_npn_relacionados != row['UnidadesEnRPH']:
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'Npn': row['Npn'],
                            'DestinoEconomico': row['DestinoEcconomico'],
                            'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                            'AreaTotalConstruida':row['AreaTotalConstruida'],
                            'CaracteristicaPredio':row['CaracteristicaPredio'],
                            'AreaTotalTerreno':row['AreaTotalTerreno'],
                            'ModoAdquisicion':row['ModoAdquisicion'],
                            'UnidadesEnRPH': row['UnidadesEnRPH'],
                            'Unidades Prediales': conteo_npn_relacionados,
                            'Observacion': 'Unidades Prediales en ficha resumen no coinciden con el total de Unidades.',
                            'Radicado':row['Radicado'],
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
            '''
            
            # Guardar resultados en un archivo si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_UnidadesEnRPH_Fichas.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros con discrepancia en UnidadesEnRPH.")
            else:
                messagebox.showinfo("Sin errores", "No se encontraron registros con discrepancias en UnidadesEnRPH.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
        
    def validar_informalidad_edificio(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_informalidad_edificio")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre cada fila del DataFrame
            for _, row in df.iterrows():
                npn = str(row['Npn'])
                if len(npn) >= 24:
                    # Validar el dígito 22 y que los dígitos 23 y 24 sean '00'
                    if npn[21] == '2' and npn[22:24] != '00':  # Dígito 22 es igual a '2' y los dígitos 23 y 24 no son '00'
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'Npn': row['Npn'],
                            'Observacion': 'Informalidad Con edificio: Dígitos 23 y 24 deben ser 00',
                            'Radicado':row['Radicado'],
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        print(f"Condición de error encontrada: {resultado}")
            '''
            
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Informalidadconedificio.xlsx'
                sheet_name = 'Fichas Faltantes'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Se encontraron {len(resultados)} registros con errores de informalidad con edificio.")
            else:
                messagebox.showinfo("Información", "No se encontraron errores de informalidad con edificio en las fichas.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
    def validar_informalidad_con_piso(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Fichas'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            # Leer el archivo Excel
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_informalidad_con_piso")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            resultados = []

            # Iterar sobre cada fila del DataFrame
            for index, row in df.iterrows():
                npn = str(row.get('Npn', ''))  # Manejar valores nulos con get
                if len(npn) >= 30:  # Validar longitud mínima
                    try:
                        # Obtener el número formado por los últimos 4 dígitos
                        ultimos_cuatro_digitos = int(npn[-4:])

                        # Verificar condición
                        if ultimos_cuatro_digitos >= 1000:
                            resultado = {
                                'NroFicha': row.get('NroFicha', 'Sin dato'),
                                'Npn': npn,
                                'Observacion': 'Informalidad mal diligenciada',
                                'Radicado':row['Radicado'],
                                'Nombre Hoja': nombre_hoja
                            }
                            resultados.append(resultado)
                            print(f"Fila {index}: Error encontrado: {resultado}")
                    except ValueError as ve:
                        print(f"Error al procesar fila {index}: {ve} - Npn: {npn}")
                else:
                    print(f"Fila {index}: Npn no cumple con la longitud mínima (30 caracteres)")

            print(f"Total de errores encontrados: {len(resultados)}")
            '''
            
            if resultados:
                # Crear DataFrame con resultados
                df_resultado = pd.DataFrame(resultados)
                output_file = 'InformalidadConpiso.xlsx'
                sheet_name = 'ErroresInformalidadConPiso'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Se encontraron {len(resultados)} errores.")
            else:
                print("No se encontraron errores.")
                messagebox.showinfo("Información", "No se encontraron registros con errores.")
            '''
            return resultados

        except Exception as e:
            print(f"Error general: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
   
    