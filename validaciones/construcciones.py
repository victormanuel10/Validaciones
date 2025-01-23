# -- coding: utf-8 --
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
            # Leer el archivo Excel, especificando las hojas 'Construcciones' y 'Fichas'
            df_construcciones = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            print(f"función: validar_construcciones_No_convencionales")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            '''
            
            print(f"Dimensiones del DataFrame Construcciones: {df_construcciones.shape}")
            print(f"Dimensiones del DataFrame Fichas: {df_fichas.shape}")
            print(f"Columnas en el DataFrame Construcciones: {df_construcciones.columns.tolist()}")
            print(f"Columnas en el DataFrame Fichas: {df_fichas.columns.tolist()}")
            '''
            # Realizar un merge para incluir la columna 'Npn' desde la hoja 'Fichas'
            df_construcciones = df_construcciones.merge(
                df_fichas[['NroFicha', 'Npn']],
                on='NroFicha',
                how='left'
            )

            #print(f"Dimensiones del DataFrame después del merge: {df_construcciones.shape}")
            #print(f"Columnas después del merge: {df_construcciones.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame fusionado
            for index, row in df_construcciones.iterrows():
                conv = row['ConvencionalNoConvencional']
                calificacion = row['calificacionNoConvencional']
                Puntos = row['Puntos ']
                npn = row.get('Npn')  # Traer la columna Npn después del merge

                # Verificar las condiciones de validación
                
                if conv == 'No Convencional' and pd.isna(calificacion) and pd.isna(Puntos):
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'Npn': npn,# Desde Construcciones
                        'secuencia': row['secuencia'],  # Desde Construcciones
                        'TipoConstruccion': row['TipoConstruccion'],  # Desde Construcciones
                        'ConvencionalNoConvencional': row['ConvencionalNoConvencional'],  # Desde Construcciones
                        'calificacionNoConvencional': row['calificacionNoConvencional'],  # Desde Construcciones
                        'Observacion': 'Calificación de construcción no convencional es nula',
                        'Radicado':row['Radicado'],  # Desde Construcciones
                        'Nombre Hoja': 'Construcciones'  # Constante
                    }
                resultados.append(resultado)
                #print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

            # Mostrar resultados
            if resultados:
                #df_resultado = pd.DataFrame(resultados)
                #print(f"Total de registros no convencionales encontrados: {len(resultados)}")

                '''
                # Descomentar para guardar resultados en un archivo Excel
                output_file = 'CONSTRUCCIONES_VALIDACION.xlsx'
                sheet_name = 'CONSTRUCCIONES_VALIDACION'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"No convencionales. {len(resultados)} registros.")
                
            else:
                messagebox.showinfo("Información", "No se encontraron registros No Convencionales.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
    def tipo_construccion_noconvencionales(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Construcciones'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            # Leer las hojas Construcciones y Fichas
            df_construcciones = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            print(f"función: tipo_construccion_noconvencionales")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            '''print(f"Dimensiones del DataFrame Construcciones: {df_construcciones.shape}")
            print(f"Dimensiones del DataFrame Fichas: {df_fichas.shape}")
            print(f"Columnas en el DataFrame Construcciones: {df_construcciones.columns.tolist()}")
            print(f"Columnas en el DataFrame Fichas: {df_fichas.columns.tolist()}")'''

            # Realizar el merge para incluir la columna Npn
            df_construcciones = df_construcciones.merge(
                df_fichas[['NroFicha', 'Npn']],
                on='NroFicha',
                how='left'
            )

            '''print(f"Dimensiones del DataFrame después del merge: {df_construcciones.shape}")
            print(f"Columnas después del merge: {df_construcciones.columns.tolist()}")'''

            resultados = []

            # Validación
            for index, row in df_construcciones.iterrows():
                tipo_construccion = row['TipoConstruccion']
                no_convencional = row['ConvencionalNoConvencional']
                npn = row.get('Npn')  # Extraer Npn del DataFrame fusionado

                # Condición: TipoConstruccion no debe tener valor (vacío o nulo) o ser diferente de 'N'
                if no_convencional == 'No Convencional' and tipo_construccion not in ['', 'N']:
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'Npn': npn,
                        'secuencia': row['secuencia'],
                        'TipoConstruccion': row['TipoConstruccion'],
                        'ConvencionalNoConvencional': row['ConvencionalNoConvencional'],
                        'calificacionNoConvencional': row['calificacionNoConvencional'],
                        'Observacion': 'TipoConstruccion debe ser vacío o "N" si es No Convencional',
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    #print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

            '''
            if resultados:
                df_resultado = pd.DataFrame(resultados)

                output_file = 'TipoConstruccion.xlsx'
                sheet_name = 'TipoConstruccion'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")

                messagebox.showinfo("Éxito", f"TipoConstruccion inválido en No Convencionales. {len(resultados)} registros.")
            else:
                messagebox.showinfo("Información", "No se encontraron registros tipo construcción inválidos.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
    
    def areaconstruida_mayora1000(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Construcciones'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            # Leer las hojas Construcciones y Fichas
            df_construcciones = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            print(f"función: areaconstruida_mayora1000")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            '''print(f"Dimensiones del DataFrame Construcciones: {df_construcciones.shape}")
            print(f"Dimensiones del DataFrame Fichas: {df_fichas.shape}")
            print(f"Columnas en el DataFrame Construcciones: {df_construcciones.columns.tolist()}")
            print(f"Columnas en el DataFrame Fichas: {df_fichas.columns.tolist()}")'''

            # Realizar el merge para incluir la columna Npn
            df_construcciones = df_construcciones.merge(
                df_fichas[['NroFicha', 'Npn']],
                on='NroFicha',
                how='left'
            )

            '''print(f"Dimensiones del DataFrame después del merge: {df_construcciones.shape}")
            print(f"Columnas después del merge: {df_construcciones.columns.tolist()}")'''

            resultados = []

            # Validación
            for index, row in df_construcciones.iterrows():
                AreaConstruida = row['AreaConstruida']
                npn = row.get('Npn')  # Extraer Npn del DataFrame fusionado

                # Validar si el área construida es mayor o igual a 1000
                if AreaConstruida >= 1000:
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'Npn': npn,
                        'secuencia': row['secuencia'],
                        'AreaConstruida': row['AreaConstruida'],
                        'Observacion': 'El área construida es mayor a 1000 mts (verificar)',
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    #print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

            '''
            if resultados:
                df_resultado = pd.DataFrame(resultados)

                output_file = 'Areamayor1000.xlsx'
                sheet_name = 'Areamayor1000'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"AreaMayor a 1000. {len(resultados)} registros.")
            else:
                messagebox.showinfo("Información", "No se encontraron registros Areamayor1000.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
    def validar_secuencia_construcciones_vs_calificaciones(self):
        archivo_excel = self.archivo_entry.get()

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return

        try:
            # Leer las hojas Construcciones, CalificacionesConstrucciones y Fichas
            df_construcciones = pd.read_excel(archivo_excel, sheet_name='Construcciones')
            df_calificaciones = pd.read_excel(archivo_excel, sheet_name='CalificacionesConstrucciones')
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            print(f"Leyendo archivo: {archivo_excel}")
            print(f"Dimensiones de Construcciones: {df_construcciones.shape}")
            '''print(f"Dimensiones de CalificacionesConstrucciones: {df_calificaciones.shape}")
            print(f"Dimensiones de Fichas: {df_fichas.shape}")'''

            # Realizar merge entre Construcciones y Fichas para incluir la columna Npn
            df_construcciones = df_construcciones.merge(
                df_fichas[['NroFicha', 'Npn']],
                on='NroFicha',
                how='left'
            )

            #print(f"Dimensiones de Construcciones después del merge: {df_construcciones.shape}")
            #print(f"Columnas después del merge: {df_construcciones.columns.tolist()}")

            # Filtrar Construcciones: excluir registros donde ConvencionalNoConvencional sea 'No Convencional'
            df_construcciones_filtrado = df_construcciones[
                df_construcciones['ConvencionalNoConvencional'] != 'No Convencional'
            ]

            #print(f"Filtrado de Construcciones (excluyendo 'No Convencional'): {df_construcciones_filtrado.shape}")

            # Extraer las secuencias de cada hoja
            secuencia_construcciones = set(df_construcciones_filtrado['secuencia'].dropna())
            secuencia_calificaciones = set(df_calificaciones['secuencia'].dropna())

            # Encontrar secuencias en Construcciones que no están en CalificacionesConstrucciones
            secuencias_faltantes = secuencia_construcciones - secuencia_calificaciones

            resultados = []
            for secuencia in secuencias_faltantes:
                # Buscar los datos asociados a la secuencia faltante en la hoja Construcciones
                construccion_fila = df_construcciones_filtrado.loc[df_construcciones_filtrado['secuencia'] == secuencia]

                if not construccion_fila.empty:
                    nro_ficha = construccion_fila.iloc[0]['NroFicha']  # Obtener NroFicha
                    radicado = construccion_fila.iloc[0].get('Radicado', 'N/A')  # Obtener Radicado, si existe
                    npn = construccion_fila.iloc[0].get('Npn', 'N/A')  # Obtener Npn desde la hoja Fichas

                    resultado = {
                        'NroFicha': nro_ficha,
                        'secuencia': secuencia,
                        'Observacion': 'Construcción (secuencia) está en Construcciones pero no en CalificacionesConstrucciones',
                        'Npn': npn,  # Agregar Npn al resultado
                        'Radicado': radicado,
                        'Nombre Hoja': 'Construcciones'
                    }
                    resultados.append(resultado)
                    #print(f"secuencia faltante: {resultado}")

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
            # Leer las hojas Construcciones y Fichas
            df_construcciones = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            '''print(f"Dimensiones de Construcciones: {df_construcciones.shape}")
            print(f"Dimensiones de Fichas: {df_fichas.shape}")'''

            # Realizar merge para incluir la columna Npn desde la hoja Fichas
            df_construcciones = df_construcciones.merge(
                df_fichas[['NroFicha', 'Npn']],
                on='NroFicha',
                how='left'
            )

            #print(f"Dimensiones de Construcciones después del merge: {df_construcciones.shape}")
            #print(f"Columnas después del merge: {df_construcciones.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df_construcciones.iterrows():
                edad_construccion = row.get('EdadConstruccion', None)

                # Verificar si 'EdadConstruccion' es <= 0
                if edad_construccion is not None and edad_construccion <= 0:
                    resultado = {
                        'NroFicha': row.get('NroFicha', ''),
                        'Npn':row.get('Npn',''),
                        'secuencia': row.get('secuencia', ''),
                        'EdadConstruccion': edad_construccion,
                        'Observacion': 'Edad de construcción inválida (<= 0)',
                        'Radicado': row.get('Radicado', ''),
                        'Nombre Hoja': 'Construcciones'
                    }
                    resultados.append(resultado)
                    #print(f"Fila {index}: Agregado a resultados: {resultado}")

            #print(f"Total de errores encontrados: {len(resultados)}")

            # Agregar resultados a la lista general
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    def validar_porcentaje_construido(self):
        """
        Verifica que en la hoja 'Construcciones' no haya valores en la columna 'PorcentajeConstruido' 
        que sean menores a 100. Si los hay, genera un error por cada registro que cumple la condición
        e incluye la columna 'Npn' de la hoja 'Fichas'.
        """
        archivo_excel = self.archivo_entry.get()
        
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer las hojas 'Construcciones' y 'Fichas'
            df_construcciones = pd.read_excel(archivo_excel, sheet_name='Construcciones')
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')
            
            # Realizar merge para incluir la columna Npn desde la hoja Fichas
            df_construcciones = df_construcciones.merge(
                df_fichas[['NroFicha', 'Npn']],
                on='NroFicha',
                how='left'
            )

            print(f"Leyendo archivo: {archivo_excel}")
            '''print(f"Dimensiones de Construcciones después del merge: {df_construcciones.shape}")
            print(f"Columnas después del merge: {df_construcciones.columns.tolist()}")'''

            # Filtrar registros donde 'PorcentajeConstruido' es menor a 100
            errores = df_construcciones[df_construcciones['PorcentajeConstruido'] < 100]
            
            resultados = []

            # Generar lista de errores
            for _, row in errores.iterrows():
                resultado = {
                    'NroFicha': row['NroFicha'],
                    'secuencia':row['secuencia'],
                    'Npn': row.get('Npn', ''),
                    'PorcentajeConstruido': row['PorcentajeConstruido'],
                    'Observacion': 'El valor de PorcentajeConstruido es menor a 100',
                    'Radicado': row.get('Radicado', ' '),
                    'Nombre Hoja': 'Construcciones'
                }
                resultados.append(resultado)
                #print(f"Error encontrado: {resultado}")

            #print(f"Total de errores encontrados: {len(resultados)}")

            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
    
    def validar_construcciones_puntos(self):
        archivo_excel = self.archivo_entry.get()

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return []

        try:
            # Leer la hoja 'Construcciones' y 'Fichas'
            df_construcciones = pd.read_excel(archivo_excel, sheet_name='Construcciones')
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            # Realizar merge para incluir la columna Npn desde la hoja Fichas
            df_construcciones = df_construcciones.merge(
                df_fichas[['NroFicha', 'Npn']],
                on='NroFicha',
                how='left'
            )

            df_construcciones = df_construcciones[df_construcciones['ConvencionalNoConvencional'] == 'Convencional']
            # Lista para almacenar errores
            errores = []

            # Validar cada fila en la columna 'Puntos'
            for index, row in df_construcciones.iterrows():
                # Validar si 'Puntos' es nulo
                
                
                if pd.isnull(row['Puntos ']):
                    errores.append({
                        
                        'Npn': row.get('Npn', ''),
                        'NroFicha': row['NroFicha'],
                        'Observacion': "La columna 'Puntos' contiene valores nulos.",
                        'Puntos': row['Puntos '],
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': 'Construcciones'
                    })

                # Validar si 'Puntos' es menor a 1
                elif row['Puntos '] < 1:
                    errores.append({
                        'Puntos': row['Puntos '],
                        'Observacion': "La columna 'Puntos' contiene valores menores a 1.",
                        'NroFicha': row['NroFicha'],
                        'Npn': row.get('Npn', 'Sin Npn'),  # Agregar Npn al resultado
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': 'Construcciones'
                    })

            # Solo crear y guardar el DataFrame si hay errores
            '''
            if errores:
                df_errores = pd.DataFrame(errores)
                output_file = 'Errores_Construcciones_Puntos.xlsx'
                sheet_name = 'Errores Construcciones'
                df_errores.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Errores encontrados", f"Se encontraron {len(errores)} errores en la columna 'Puntos' de la hoja 'Construcciones'.")
            else:
                messagebox.showinfo("Validación completada", "No se encontraron errores en la columna 'Puntos' de la hoja 'Construcciones'.")
            '''
            return errores

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
        
    def validar_secuencia_calificaciones_vs_construcciones(self):
        archivo_excel = self.archivo_entry.get()

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return
        
        try:
            # Leer las hojas Construcciones, CalificacionesConstrucciones y Fichas
            df_construcciones = pd.read_excel(archivo_excel, sheet_name='Construcciones')
            df_calificaciones = pd.read_excel(archivo_excel, sheet_name='CalificacionesConstrucciones')
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            print(f"Leyendo archivo: {archivo_excel}")
            print(f"Dimensiones de Construcciones: {df_construcciones.shape}")
            '''print(f"Dimensiones de CalificacionesConstrucciones: {df_calificaciones.shape}")
            print(f"Dimensiones de Fichas: {df_fichas.shape}")'''

            # Extraer las secuencias de cada hoja
            secuencia_construcciones = set(df_construcciones['secuencia'].dropna())
            secuencia_calificaciones = set(df_calificaciones['secuencia'].dropna())

            # Encontrar secuencias en CalificacionesConstrucciones que no están en Construcciones
            secuencias_faltantes = secuencia_calificaciones - secuencia_construcciones

            # Realizar un merge con las hojas Fichas y Construcciones para obtener NroFicha y Npn
            df_construcciones_fichas = df_construcciones.merge(
                df_fichas[['NroFicha', 'Npn']], 
                on='NroFicha', 
                how='left'
            )

            resultados = []
            for secuencia in secuencias_faltantes:
                # Buscar la fila correspondiente a la secuencia faltante en CalificacionesConstrucciones
                fila = df_calificaciones[df_calificaciones['secuencia'] == secuencia].iloc[0]
                
                # Buscar la fila correspondiente a la secuencia faltante en el merge de Construcciones y Fichas
                fila_construccion = df_construcciones_fichas[df_construcciones_fichas['secuencia'] == secuencia].iloc[0]
                
                resultado = {
                    'secuencia': secuencia,
                    'Npn': fila_construccion.get('Npn', ''),  # Obtener la columna Npn
                    'Observacion': 'Construccion (secuencia) está en CalificacionesConstrucciones pero no en Construcciones',
                    'Nombre Hoja': 'CalificacionesConstrucciones'
                }
                resultados.append(resultado)
                #print(f"secuencia faltante: {resultado}")

            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    
    def validar_secuencia_convencional(self):
        """
        Valida que si en la columna 'ConvencionalNoConvencional' de la hoja 'Construcciones' el valor es 'Convencional',
        al menos una secuencia de la hoja 'Construcciones' exista en el campo 'secuencia' de la hoja 'ConstruccionesGenerales'.
        """
        archivo_excel = self.archivo_entry.get()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer las hojas 'Construcciones', 'ConstruccionesGenerales' y 'Fichas'
            df_construcciones = pd.read_excel(archivo_excel, sheet_name='Construcciones')
            df_generales = pd.read_excel(archivo_excel, sheet_name='ConstruccionesGenerales')
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            # Filtrar las secuencias de 'Construcciones' donde 'ConvencionalNoConvencional' es 'Convencional'
            convencionales = df_construcciones[df_construcciones['ConvencionalNoConvencional'] == 'Convencional']
            
            # Obtener las secuencias únicas en 'ConstruccionesGenerales'
            secuencias_generales = set(df_generales['Secuencia'].dropna().unique())

            # Realizar un merge con la hoja Fichas para incluir la columna 'Npn'
            df_construcciones = convencionales.merge(
                df_fichas[['NroFicha', 'Npn']], 
                on='NroFicha', 
                how='left'
            )

            # Lista para almacenar los errores
            errores = []

            # Validar cada secuencia en 'convencionales' para asegurarse de que exista en 'ConstruccionesGenerales'
            for index, row in df_construcciones.iterrows():
                secuencia = row['secuencia']
                if pd.isna(secuencia) or secuencia not in secuencias_generales:
                    errores.append({
                        'NroFicha': row['NroFicha'],
                        'Npn': row.get('Npn', ''),
                        'secuencia': secuencia,
                        'Observacion': "Construcción convencional (secuencia) no encontrada en 'ConstruccionesGenerales'",
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': 'Construcciones'
                    })
            
            '''
            # Guardar los resultados en un archivo Excel si hay errores
            if errores:
                df_errores = pd.DataFrame(errores)
                output_file = 'Errores_Convencional_Secuencia.xlsx'
                df_errores.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Errores encontrados", f"Se encontraron {len(errores)} secuencias sin correspondencia en 'ConstruccionesGenerales'.")
            else:
                messagebox.showinfo("Sin errores", "Todas las secuencias 'Convencional' están presentes en 'ConstruccionesGenerales'.")
            '''
            return errores

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        

    def validar_secuencia_convencional_calificaciones(self):
        """
        Valida que si en la columna 'ConvencionalNoConvencional' de la hoja 'Construcciones' el valor es 'Convencional',
        al menos una secuencia de la hoja 'Construcciones' exista en el campo 'secuencia' de la hoja 'CalificacionesConstrucciones'.
        """
        archivo_excel = self.archivo_entry.get()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer las hojas 'Construcciones', 'CalificacionesConstrucciones' y 'Fichas'
            df_construcciones = pd.read_excel(archivo_excel, sheet_name='Construcciones')
            df_calificaciones = pd.read_excel(archivo_excel, sheet_name='CalificacionesConstrucciones')
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            # Filtrar las secuencias de 'Construcciones' donde 'ConvencionalNoConvencional' es 'Convencional'
            convencionales = df_construcciones[df_construcciones['ConvencionalNoConvencional'] == 'Convencional']
            
            # Obtener las secuencias únicas en 'CalificacionesConstrucciones'
            secuencias_calificaciones = set(df_calificaciones['secuencia'].dropna().unique())

            # Realizar un merge con la hoja Fichas para incluir la columna 'Npn'
            df_construcciones = convencionales.merge(
                df_fichas[['NroFicha', 'Npn']], 
                on='NroFicha', 
                how='left'
            )

            # Lista para almacenar los errores
            errores = []

            # Validar cada secuencia en 'convencionales' para asegurarse de que exista en 'CalificacionesConstrucciones'
            for index, row in df_construcciones.iterrows():
                secuencia = row['secuencia']
                if secuencia not in secuencias_calificaciones:
                    errores.append({
                        'NroFicha': row['NroFicha'],
                        'secuencia': secuencia,
                        'Npn': row.get('Npn', ''),  # Agregar Npn al resultado
                        'Observacion': "Construccion convencional (secuencia) no encontrada en 'CalificacionesConstrucciones'",
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': 'Construcciones'
                    })
            
            '''
            # Guardar los resultados en un archivo Excel si hay errores
            if errores:
                df_errores = pd.DataFrame(errores)
                output_file = 'Errores_Convencional_Secuencia.xlsx'
                df_errores.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Errores encontrados", f"Se encontraron {len(errores)} secuencias sin correspondencia en 'CalificacionesConstrucciones'.")
            else:
                messagebox.showinfo("Sin errores", "Todas las secuencias 'Convencional' están presentes en 'CalificacionesConstrucciones'.")
            '''
            return errores

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
        
    def validar_no_convencional_secuencia(self):
        """
        Valida que si en la columna 'ConvencionalNoConvencional' de la hoja 'Construcciones' el valor es 'No Convencional',
        la secuencia de esa fila no esté presente en la columna 'secuencia' de la hoja 'CalificacionesConstrucciones'.
        """
        archivo_excel = self.archivo_entry.get()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer las hojas 'Construcciones', 'CalificacionesConstrucciones' y 'Fichas'
            df_construcciones = pd.read_excel(archivo_excel, sheet_name='Construcciones')
            df_calificaciones = pd.read_excel(archivo_excel, sheet_name='CalificacionesConstrucciones')
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            # Filtrar las secuencias de 'Construcciones' donde 'ConvencionalNoConvencional' es 'No Convencional'
            no_convencionales = df_construcciones[df_construcciones['ConvencionalNoConvencional'] == 'No Convencional']
            
            # Obtener las secuencias únicas en 'CalificacionesConstrucciones'
            secuencias_calificaciones = set(df_calificaciones['secuencia'].dropna().unique())

            # Realizar un merge con la hoja Fichas para incluir la columna 'Npn'
            df_construcciones = no_convencionales.merge(
                df_fichas[['NroFicha', 'Npn']], 
                on='NroFicha', 
                how='left'
            )

            # Lista para almacenar los errores
            errores = []

            # Validar cada secuencia en 'no_convencionales' para asegurarse de que no exista en 'CalificacionesConstrucciones'
            for index, row in df_construcciones.iterrows():
                secuencia = row['secuencia']
                if secuencia in secuencias_calificaciones:
                    errores.append({
                        'NroFicha': row['NroFicha'],
                        'secuencia': secuencia,
                        'Npn': row.get('Npn', 'Sin Npn'),  # Agregar Npn al resultado
                        'Observacion': "secuencia 'No Convencional' encontrada en 'CalificacionesConstrucciones'",
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': 'Construcciones'
                    })
            
            '''
            # Guardar los resultados en un archivo Excel si hay errores
            if errores:
                df_errores = pd.DataFrame(errores)
                output_file = 'Errores_No_Convencional_Secuencia.xlsx'
                df_errores.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Errores encontrados", f"Se encontraron {len(errores)} secuencias 'No Convencional' que no cumplen la condición.")
            else:
                messagebox.showinfo("Sin errores", "Todas las secuencias 'No Convencional' cumplen la condición en 'CalificacionesConstrucciones'.")
            '''
            return errores

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
        
    def validar_secuencias_repetidas(self):
        """
        Verifica si existen secuencias repetidas en las hojas 'Construcciones' y 'CalificacionesConstrucciones'.
        Genera un error si una secuencia se encuentra en ambas hojas.
        """
        archivo_excel = self.archivo_entry.get()
        
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer las hojas 'Construcciones', 'CalificacionesConstrucciones' y 'Fichas'
            df_construcciones = pd.read_excel(archivo_excel, sheet_name='Construcciones')
            df_calificaciones = pd.read_excel(archivo_excel, sheet_name='CalificacionesConstrucciones')
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            # Extraer las secuencias únicas de cada hoja
            secuencias_construcciones = df_construcciones['secuencia'].dropna().unique()
            secuencias_calificaciones = df_calificaciones['secuencia'].dropna().unique()

            # Encontrar secuencias repetidas en ambas hojas
            secuencias_repetidas = set(secuencias_construcciones) & set(secuencias_calificaciones)

            # Realizar un merge con la hoja Fichas para incluir la columna 'Npn'
            df_construcciones_repetidas = df_construcciones[df_construcciones['secuencia'].isin(secuencias_repetidas)]
            df_construcciones_repetidas = df_construcciones_repetidas.merge(
                df_fichas[['secuencia', 'Npn']], 
                on='secuencia', 
                how='left'
            )

            # Lista para almacenar los errores
            resultados = []
            for index, row in df_construcciones_repetidas.iterrows():
                secuencia = row['secuencia']
                resultados.append({
                    'secuencia': secuencia,
                    'Npn': row.get('Npn', 'Sin Npn'),  # Agregar Npn al resultado
                    'Observacion': 'secuencia duplicada en ambas hojas',
                    'Nombre Hoja': 'Construcciones'
                })
                #print(f"Secuencia duplicada encontrada: {row['secuencia']} - Npn: {row.get('Npn', 'Sin Npn')}")
            
            '''
            # Si se encuentran duplicados, guardarlos en un archivo Excel
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_Secuencias_Duplicadas.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Errores encontrados", f"Se encontraron {len(resultados)} secuencias duplicadas en ambas hojas.")
            else:
                messagebox.showinfo("Validación completada", "No se encontraron secuencias duplicadas en ambas hojas.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
        
    def validar_secuencia_unica_por_ficha(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Construcciones'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            # Leer el archivo Excel, especificando la hoja 'Construcciones' y la hoja 'Fichas'
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            print(f"funcion: validar_secuencia_unica_por_ficha")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            '''print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")'''

            # Lista para almacenar los resultados
            resultados = []

            # Agrupar el DataFrame por 'NroFicha'
            grouped = df.groupby('NroFicha')

            for name, group in grouped:
                # Verificar si hay duplicados en la columna 'Secuencia' dentro del grupo
                duplicados = group[group.duplicated(subset='secuencia', keep=False)]
                if not duplicados.empty:
                    for _, row in duplicados.iterrows():
                        # Realizar un merge con la hoja 'Fichas' para incluir la columna 'Npn'
                        npn_value = df_fichas.loc[df_fichas['secuencia'] == row['secuencia'], 'Npn'].values
                        npn_value = npn_value[0] if len(npn_value) > 0 else 'Sin Npn'  # Si no existe, se asigna 'Sin Npn'

                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'secuencia': row['secuencia'],
                            'Npn': npn_value,
                            'Observacion': 'Secuencia duplicada para el mismo NroFicha',
                            'Radicado':row['Radicado'],
                              # Incluir Npn en los resultados
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        #print(f"Secuencia duplicada encontrada: {resultado}")

            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    
    def validar_id_uso(self):
        archivo_excel = self.archivo_entry.get()

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return []

        try:
            # Leer la hoja 'Construcciones' y 'Fichas'
            df_construcciones = pd.read_excel(archivo_excel, sheet_name='Construcciones')
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

            # Realizar merge para incluir la columna Npn desde la hoja Fichas
            df_construcciones = df_construcciones.merge(
                df_fichas[['NroFicha', 'Npn']],
                on='NroFicha',
                how='left'
            )

            errores = []

            # Validar cada fila en la columna 'IdUso'
            for index, row in df_construcciones.iterrows():
                id_uso = str(row.get('IdUso', '')).strip()  # Convertir a string y eliminar espacios
                
                # Extraer los números antes del separador '|'
                try:
                    numero_inicial = int(id_uso.split('|')[0])  # Tomar el primer valor numérico
                except ValueError:
                    numero_inicial = None  # Si no es un número válido, asignar None

                # Validar si el número inicial es menor a 800
                if numero_inicial is None or numero_inicial < 800:
                    errores.append({
                        'Npn': row.get('Npn', ''),
                        'NroFicha': row['NroFicha'],
                        'Observacion': "Identificador de uso incorrecto (menor a 800)",
                        'IdUso': row['IdUso'],
                        'Radicado': row.get('Radicado', ''),
                        'Nombre Hoja': 'Construcciones'
                    })

            # Devolver los errores encontrados
            return errores

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []