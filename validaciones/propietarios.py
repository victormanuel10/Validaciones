import pandas as pd
from tkinter import messagebox
from datetime import datetime
from validaciones.ficha import Ficha

class Propietarios:
    def __init__(self, archivo_entry):
        self.archivo_entry = archivo_entry

    def procesar(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"Procesando archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            #Propietarios
            self.cedula_mujer(df)
            self.cedula_hombre(df)
            self.primer_apellido_blanco(df)
            self.primer_nombre_blanco(df)
            self.calidad_propietario_mun(df)
            self.nit_diferente_mun(df)
            self.derecho_diferente_cien()
            self.documento_blanco_cod_asig()
            self.fecha_escritura_inferior()
            self.fecha_registro_inferior()
            #Ficha
            ficha = Ficha(self.archivo_entry)
            ficha.terreno_cero()
            ficha.terreno_null()
            ficha.informal_matricula()
            ficha.matricula_mejora()
            ficha.tomo_mejora()
            ficha.modo_adquisicion_informal()
            ficha.ficha_repetida()
            ficha.rural_destino_invalido()

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")

    def cedula_mujer(self, df):
        resultados = []
        for index, row in df.iterrows():
            valor_a = str(row['TipoDocumento'])
            valor_b = row['Documento']

            print(f"Fila {index}: Valor A = '{valor_a}', Valor B = '{valor_b}'")

            if valor_a == '2|CEDULA DE CIUDADANIA MUJER' and (valor_b <= 20000000 or valor_b >= 70000000):
                resultado = {
                    'NroFicha': row['NroFicha'],
                    'TipoDocumento': row['TipoDocumento'],
                    'Documento': row['Documento'],
                    'PrimerNombre': row['PrimerNombre'],
                    'SegundoNombre': row['SegundoNombre'],
                    'PrimerApellido': row['PrimerApellido'],
                    'SegundoApellido': row['SegundoApellido'],
                    'Observacion': 'Documento no esta en rango de mujeres'
                }
                resultados.append(resultado)
                print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

        df_resultado = pd.DataFrame(resultados)
        output_file = 'CEDULA_MUJER.xlsx'
        sheet_name = 'cedula_mujer'
        df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
        print(f"Archivo guardado: {output_file}")
        print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

        messagebox.showinfo("Éxito",
                            f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")

    def cedula_hombre(self, df):
        resultados = []
        for index, row in df.iterrows():
            valor_a = str(row['TipoDocumento'])
            valor_b = row['Documento']

            print(f"Fila {index}: Valor A = '{valor_a}', Valor B = '{valor_b}'")

            if valor_a == '1|CEDULA DE CIUDADANIA HOMBRE' and ((valor_b >= 20000000 and valor_b <= 69999999) or valor_b >= 100000000):
                resultado = {
                    'NroFicha': row['NroFicha'],
                    'TipoDocumento': row['TipoDocumento'],
                    'Documento': row['Documento'],
                    'PrimerNombre': row['PrimerNombre'],
                    'SegundoNombre': row['SegundoNombre'],
                    'PrimerApellido': row['PrimerApellido'],
                    'SegundoApellido': row['SegundoApellido'],
                    'Observacion': 'Documento no esta en rango de hombre'
                }
                resultados.append(resultado)
                print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

        df_resultado = pd.DataFrame(resultados)
        output_file = 'CEDULA_HOMBRE.xlsx'
        sheet_name = 'cedula_hombre'
        df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
        print(f"Archivo guardado: {output_file}")
        print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

        messagebox.showinfo("Éxito",
                            f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
        
    def primer_apellido_blanco(self, df):
        resultados = []
        for index, row in df.iterrows():
            valor_a = row['PrimerApellido']
            valor_b = row['TipoDocumento']

            print(f"Fila {index}: Valor A = '{valor_a}'")

            if valor_b != '3|NIT' and (pd.isna(valor_a) or valor_a == ''):
                resultado = {
                    'NroFicha': row['NroFicha'],
                    'Documento': row['Documento'],
                    'PrimerNombre': row['PrimerNombre'],
                    'SegundoNombre': row['SegundoNombre'],
                    'PrimerApellido': row['PrimerApellido'],
                    'SegundoApellido': row['SegundoApellido'],
                    'Observacion': 'Primer apellido en blanco'
                }
                resultados.append(resultado)
                print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

        df_resultado = pd.DataFrame(resultados)
        output_file = 'PRIMER_APELLIDO_BLANCO.xlsx'
        sheet_name = 'PRIMER_APELLIDO'
        df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
        print(f"Archivo guardado: {output_file}")
        print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

        messagebox.showinfo("Éxito",
                            f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
        
    def primer_nombre_blanco(self, df):
        resultados = []
        for index, row in df.iterrows():
            valor_a = row['PrimerNombre']
            valor_b = row['TipoDocumento']

            print(f"Fila {index}: Valor A = '{valor_a}'")

            if valor_b != '3|NIT' and (valor_a == '' or pd.isna(valor_a)):
                resultado = {
                    'NroFicha': row['NroFicha'],
                    'Documento': row['Documento'],
                    'PrimerNombre': row['PrimerNombre'],
                    'SegundoNombre': row['SegundoNombre'],
                    'PrimerApellido': row['PrimerApellido'],
                    'SegundoApellido': row['SegundoApellido'],
                    'Observacion': 'Primer nombre en blanco'
                }
                resultados.append(resultado)
                print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

        df_resultado = pd.DataFrame(resultados)
        output_file = 'PRIMER_NOMBRE_BLANCO.xlsx'
        sheet_name = 'PRIMER_NOMBRE'
        df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
        print(f"Archivo guardado: {output_file}")
        print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

        messagebox.showinfo("Éxito",
                            f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
        
    def calidad_propietario_mun(self,df):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: calidad_propietario_mun")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                valor_a = row['CalidadPropietario']
                valor_b = row['TipoDocumento']

                print(f"Fila {index}: Valor A = '{valor_a}'")

                # Verificar las condiciones
                if valor_a != '4|MUNICIPAL' and valor_b == '3|NIT':
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'TipoDocumento': row['TipoDocumento'],
                        'Documento': row['Documento'],
                        'CalidadPropietario': row['CalidadPropietario'],
                        'RazonSocial': row['RazonSocial'],
                        'Observacion': 'Calidad del propietario diferente para nit del Municipio'
                    }
                    resultados.append(resultado)
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

            print(f"Total de resultados encontrados: {len(resultados)}")

            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)

            # Guardar el resultado en un nuevo archivo Excel
            output_file = 'CALIDAD_PROP_MUN.xlsx'
            sheet_name = 'CALIDAD_PROP_MUN'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            messagebox.showinfo("Éxito",
                                f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    def nit_diferente_mun(self,df):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: nit_diferente_mun")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                valor_a = row['CalidadPropietario']
                valor_b = row['TipoDocumento']

                print(f"Fila {index}: Valor A = '{valor_a}'")

                # Verificar las condiciones
                if valor_a == '4|MUNICIPAL' and valor_b != '3|NIT':
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'TipoDocumento': row['TipoDocumento'],
                        'Documento': row['Documento'],
                        'CalidadPropietario': row['CalidadPropietario'],
                        'RazonSocial': row['RazonSocial'],
                        'Observacion': 'tipo de documento diferente para nit del municipio'
                    }
                    resultados.append(resultado)
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

            print(f"Total de resultados encontrados: {len(resultados)}")

            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)

            # Guardar el resultado en un nuevo archivo Excel
            output_file = 'NIT_DIFERENTE_MUN.xlsx'
            sheet_name = 'NIT_DIFERENTE_MUN'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            messagebox.showinfo("Éxito",
                                f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    def derecho_diferente_cien(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: derecho_diferente_cien")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Agrupar el DataFrame por 'NroFicha'
            grouped = df.groupby('NroFicha')

            for name, group in grouped:
                valor_b_sum = group['Derecho'].sum()

                # Si la suma de 'Derecho' no es 100, guardar los valores
                if round(valor_b_sum, 3) != 100:
                    print(f"suma de derechos no es 100: {valor_b_sum}")
                    for _, row in group.iterrows():
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'TipoDocumento': row['TipoDocumento'],
                            'Documento': row['Documento'],
                            'Derecho': row['Derecho'],
                            'Observacion': 'Porcentaje de dominio incompleto diferente a cero, falta: ' + str(100 - valor_b_sum)
                        }
                        resultados.append(resultado)
                        print(f"Fila {_} agregado: {resultado}")

            print(f"Total de resultados encontrados: {len(resultados)}")

            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)

            # Guardar el resultado en un nuevo archivo Excel
            output_file = 'DERECHO_DIFERENTE_CIEN.xlsx'
            sheet_name = 'DERECHO_DIFERENTE_CIEN'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            messagebox.showinfo("Éxito",
                                f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    def documento_blanco_cod_asig(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: documento_blanco_cod_asig")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iloc[0:].iterrows():
                valor_a = row['TipoDocumento']
                valor_b = row['Documento']

                print(f"Fila {index}: Valor A = '{valor_a}', Valor B = '{valor_b}'")

                # Verificar las condiciones
                if valor_a == '8|CODIGO ASIGNADO' and pd.notna(valor_b):
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'TipoDocumento': row['TipoDocumento'],
                        'Documento': row['Documento'],
                        'PrimerNombre': row['PrimerNombre'],
                        'SegundoNombre': row['SegundoNombre'],
                        'PrimerApellido': row['PrimerApellido'],
                        'SegundoApellido': row['SegundoApellido'],
                        'Escritura': row['Escritura'],
                        'FechaEscritura': row['FechaEscritura'],
                        'Entidad': row['Entidad'],
                        'Observacion': 'Documento diferente a blanco para código asignado'
                    }
                    resultados.append(resultado)
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

            print(f"Total de resultados encontrados: {len(resultados)}")

            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)

            # Guardar el resultado en un nuevo archivo Excel
            output_file = 'DOCUMENTO_CODIGO_ASIGNADO.xlsx'
            sheet_name = 'DOCUMENTO_CODIGO_ASIGNADO'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            messagebox.showinfo("Éxito",
                                f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    def fecha_escritura_inferior(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'

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

            # Umbral de fecha (1 de enero de 1778)
            fecha_umbral = datetime(1778, 1, 1)

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                fecha_str = row['FechaEscritura']

                try:
                    # Convertir la cadena de texto a objeto de fecha
                    fecha_obj = datetime.strptime(fecha_str, "%d/%m/%Y")

                    # Verificar si la fecha es anterior a 1778
                    if fecha_obj < fecha_umbral:
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'FechaEscritura': fecha_str,
                            'Observacion': 'Fecha anterior a 1778'
                        }
                        resultados.append(resultado)
                        print(f"Fila {index}: Fecha '{fecha_str}' es anterior a 1778. Agregado a resultados.")

                except ValueError:
                    # Manejo de errores si la conversión falla
                    print(f"Error en el formato de fecha en la fila {index}: '{fecha_str}'")

            print(f"Total de fechas anteriores a 1778 encontradas: {len(resultados)}")

            if len(resultados) > 0:
                # Crear un nuevo DataFrame con los resultados
                df_resultado = pd.DataFrame(resultados)

                # Guardar el resultado en un nuevo archivo Excel
                output_file = 'FECHAS_ESCRITURA_INFERIORES_1778.xlsx'
                sheet_name = 'fechas_inferiores_1778'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

                messagebox.showinfo("Éxito",
                                    f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
            else:
                print("No se encontraron fechas anteriores a 1778.")
                messagebox.showinfo("Información", "No se encontraron fechas anteriores a 1778.")

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
    def fecha_registro_inferior(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'

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

            # Umbral de fecha (1 de enero de 1778)
            fecha_umbral = datetime(1778, 1, 1)

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                fecha_str = row['FechaEscritura']

                try:
                    # Convertir la cadena de texto a objeto de fecha
                    fecha_obj = datetime.strptime(fecha_str, "%d/%m/%Y")

                    # Verificar si la fecha es anterior a 1778
                    if fecha_obj < fecha_umbral:
                        resultado = {
                            'NroFicha': row['NroFicha'],
                            'FechaEscritura': fecha_str,
                            'Observacion': 'Fecha registro inferior a 1778'
                        }
                        resultados.append(resultado)
                        print(f"Fila {index}: Fecha '{fecha_str}' es anterior a 1778. Agregado a resultados.")

                except ValueError:
                    # Manejo de errores si la conversión falla
                    print(f"Error en el formato de fecha en la fila {index}: '{fecha_str}'")

            print(f"Total de fechas anteriores a 1778 encontradas: {len(resultados)}")

            if len(resultados) > 0:
                # Crear un nuevo DataFrame con los resultados
                df_resultado = pd.DataFrame(resultados)

                # Guardar el resultado en un nuevo archivo Excel
                output_file = 'FECHAS_REGISTRO_INFERIORES_1778.xlsx'
                sheet_name = 'fechas_inferiores_1778'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

                messagebox.showinfo("Éxito",
                                    f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
            else:
                print("No se encontraron fechas anteriores a 1778.")
                messagebox.showinfo("Información", "No se encontraron fechas anteriores a 1778.")

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")