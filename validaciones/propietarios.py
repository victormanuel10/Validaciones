# -- coding: utf-8 --
import pandas as pd
from tkinter import messagebox
from collections import Counter
from datetime import datetime


class Propietarios:
    def __init__(self, archivo_entry):
        self.archivo_entry = archivo_entry
        
    
    def validar_documento_inicia_con_cero(self):
        """
        Verifica que en la hoja 'Propietarios' no haya valores en la columna 'Documento' que inicien con '0'.
        Si los hay, genera un error por cada registro que cumple la condición.
        """
        archivo_excel = self.archivo_entry.get()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'Propietarios'
            df_propietarios = pd.read_excel(archivo_excel, sheet_name='Propietarios')
            
            # Filtrar los documentos que inician con '0'
            errores = df_propietarios[df_propietarios['Documento'].astype(str).str.startswith('0')]
            
            resultados = []

            # Generar una lista de errores
            for _, row in errores.iterrows():
                resultado = {
                    'NroFicha': row['NroFicha'],
                    'Observacion': 'El documento inicia con "0"',
                    'TipoDocumento':row['TipoDocumento'],
                    'Documento': row['Documento'],
                    'CalidadPropietario':row['CalidadPropietario'],
                    'Derecho':row['Derecho'],
                    'CalidadPropietarioOficial':row['CalidadPropietarioOficial'],
                    'Fecha':row['Fecha'],
                    'CodigoFideicomiso':row['CodigoFideicomiso'],
                    'Escritura':row['Escritura'],
                    'Entidad':row['Entidad'],
                    'EntidadDepartamento':row['EntidadDepartamento'],
                    'EntidadMunicipio':row['EntidadMunicipio'],
                    'NumeroFallo':row['NumeroFallo'],
                    'RazonSocial':row['RazonSocial'],
                    'PrimerNombre':row['PrimerNombre'],
                    'SegundoNombre':row['SegundoNombre'],
                    'PrimerApellido':row['PrimerApellido'],
                    'SegundoApellido':row['SegundoApellido'],
                    'Sexo':row['Sexo'],
                    
                    'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                    'Tomo':row['Tomo'],
                    'Radicado':row['Radicado'],
                    
                    'Nombre Hoja': 'Propietarios'
                }
                resultados.append(resultado)
            '''
            
            # Guardar los resultados en un archivo Excel si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_Documento_Inicia_Con_Cero_Propietarios.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Errores encontrados: {len(resultados)} registros con Documento que inicia con '0'.")
            else:
                messagebox.showinfo("Sin errores", "No se encontraron Documentos que inicien con '0' en la hoja 'Propietarios'.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        
    def validar_documento_sexo_femenino(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_documento_sexo")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                tipo_documento = row['TipoDocumento']
                documento = row['Documento']
                sexo = row['Sexo']

                # Intentar convertir el valor de 'Documento' a entero
                try:
                    documento = int(documento)
                except ValueError:
                    # Si no se puede convertir a entero, saltar la fila
                    print(f"Fila {index}: El valor del documento no es un número válido. Saltando fila.")
                    continue

                # Verificar si el Tipo de Documento es '10|CEDULA DE CIUDADANIA'
                if tipo_documento == '10|CEDULA DE CIUDADANIA':
                    # Validar que el Documento esté fuera del rango [20000000, 70000000] y que el Sexo sea 'F|FEMENINO'
                    if (documento <= 20000000 or documento >= 70000000) and sexo == 'F|FEMENINO':
                        resultado = {
                            'NroFicha': row['NroFicha'],  # Suponiendo que existe esta columna
                            'Observacion': 'Documento fuera del rango para Sexo Femenino',
                            'TipoDocumento':row['TipoDocumento'],
                            'Documento': row['Documento'],
                            'CalidadPropietario':row['CalidadPropietario'],
                            'Derecho':row['Derecho'],
                            'CalidadPropietarioOficial':row['CalidadPropietarioOficial'],
                            'Fecha':row['Fecha'],
                            'CodigoFideicomiso':row['CodigoFideicomiso'],
                            'Escritura':row['Escritura'],
                            'Entidad':row['Entidad'],
                            'EntidadDepartamento':row['EntidadDepartamento'],
                            'EntidadMunicipio':row['EntidadMunicipio'],
                            'NumeroFallo':row['NumeroFallo'],
                            'RazonSocial':row['RazonSocial'],
                            'PrimerNombre':row['PrimerNombre'],
                            'SegundoNombre':row['SegundoNombre'],
                            'PrimerApellido':row['PrimerApellido'],
                            'SegundoApellido':row['SegundoApellido'],
                            'Sexo':row['Sexo'],
                            
                            'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                            'Tomo':row['Tomo'],
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
                
                # Guardar el resultado en un archivo Excel
                output_file = 'ERRORES_DOCUMENTO_SEXO.xlsx'
                sheet_name = 'ErroresDocumentoSexo'
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

        
   

    def validar_tipo_documento_sexo(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_tipo_documento_sexo")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                tipo_documento = row['TipoDocumento']
                sexo = row['Sexo']

                # Verificar si el Tipo de Documento es '3|NIT'
                if tipo_documento == '3|NIT':
                    # Validar que el Sexo sea 'N|NO BINARIO'
                    if pd.notna(sexo):
                        resultado = {
                            'NroFicha': row['NroFicha'],  # Suponiendo que existe esta columna
                            
                            'Observacion': 'El tipo de documento es 3|NIT, pero el sexo no es Correcto',
                            'TipoDocumento':row['TipoDocumento'],
                            'Documento': row['Documento'],
                            'CalidadPropietario':row['CalidadPropietario'],
                            'Derecho':row['Derecho'],
                            'CalidadPropietarioOficial':row['CalidadPropietarioOficial'],
                            'Fecha':row['Fecha'],
                            'CodigoFideicomiso':row['CodigoFideicomiso'],
                            'Escritura':row['Escritura'],
                            'Entidad':row['Entidad'],
                            'EntidadDepartamento':row['EntidadDepartamento'],
                            'EntidadMunicipio':row['EntidadMunicipio'],
                            'NumeroFallo':row['NumeroFallo'],
                            'RazonSocial':row['RazonSocial'],
                            'PrimerNombre':row['PrimerNombre'],
                            'SegundoNombre':row['SegundoNombre'],
                            'PrimerApellido':row['PrimerApellido'],
                            'SegundoApellido':row['SegundoApellido'],
                            'Sexo':row['Sexo'],
                            
                            'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                            'Tomo':row['Tomo'],
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
                
                # Guardar el resultado en un archivo Excel
                output_file = 'ERRORES_TIPO_DOCUMENTO_SEXO.xlsx'
                sheet_name = 'ErroresTipoDocumentoSexo'
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
    
    def validar_documento_sexo_masculino(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'
        
        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_documento_sexo_masculino")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados
            resultados = []

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                tipo_documento = row['TipoDocumento']
                documento = row['Documento']
                sexo = row['Sexo']

                # Intentar convertir el valor de 'Documento' a entero
                try:
                    documento = int(documento)
                except ValueError:
                    # Si no se puede convertir a entero, saltar la fila
                    print(f"Fila {index}: El valor del documento no es un número válido. Saltando fila.")
                    continue

                # Verificar si el Tipo de Documento es '10|CEDULA DE CIUDADANIA'
                if tipo_documento == '10|CEDULA DE CIUDADANIA':
                    # Validar si el Documento está entre 20000000 y 69999999, y el Sexo es 'M|MASCULINO'
                    if 20000000 <= documento <= 69999999 and sexo == 'M|MASCULINO':
                        resultado = {
                            'NroFicha': row['NroFicha'],  # Suponiendo que existe esta columna
                            'Observacion': 'Documento en rango para Cédula de Ciudadanía y Sexo Masculino',
                            'TipoDocumento':row['TipoDocumento'],
                            'Documento': row['Documento'],
                            'CalidadPropietario':row['CalidadPropietario'],
                            'Derecho':row['Derecho'],
                            'CalidadPropietarioOficial':row['CalidadPropietarioOficial'],
                            'Fecha':row['Fecha'],
                            'CodigoFideicomiso':row['CodigoFideicomiso'],
                            'Escritura':row['Escritura'],
                            'Entidad':row['Entidad'],
                            'EntidadDepartamento':row['EntidadDepartamento'],
                            'EntidadMunicipio':row['EntidadMunicipio'],
                            'NumeroFallo':row['NumeroFallo'],
                            'RazonSocial':row['RazonSocial'],
                            'PrimerNombre':row['PrimerNombre'],
                            'SegundoNombre':row['SegundoNombre'],
                            'PrimerApellido':row['PrimerApellido'],
                            'SegundoApellido':row['SegundoApellido'],
                            'Sexo':row['Sexo'],
                            
                            'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                            'Tomo':row['Tomo'],
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
                
                # Guardar el resultado en un archivo Excel
                output_file = 'ERRORES_DOCUMENTO_SEXO_MASCULINO.xlsx'
                sheet_name = 'ErroresDocumentoSexoMasculino'
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
    
    

        
        
    def primer_apellido_blanco(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'
        df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
        if df is None:
            return
        
        
        resultados = []
        for index, row in df.iterrows():
            valor_a = row['PrimerApellido']
            valor_b = row['TipoDocumento']

            print(f"Fila {index}: Valor A = '{valor_a}'")

            if valor_b != '3|NIT' and (pd.isna(valor_a) or valor_a == ''):
                resultado = {
                    'NroFicha': row['NroFicha'],
                    'Observacion': 'Primer apellido en blanco',
                    'TipoDocumento':row['TipoDocumento'],
                    'Documento': row['Documento'],
                    'CalidadPropietario':row['CalidadPropietario'],
                    'Derecho':row['Derecho'],
                    'CalidadPropietarioOficial':row['CalidadPropietarioOficial'],
                    'Fecha':row['Fecha'],
                    'CodigoFideicomiso':row['CodigoFideicomiso'],
                    'Escritura':row['Escritura'],
                    'Entidad':row['Entidad'],
                    'EntidadDepartamento':row['EntidadDepartamento'],
                    'EntidadMunicipio':row['EntidadMunicipio'],
                    'NumeroFallo':row['NumeroFallo'],
                    'RazonSocial':row['RazonSocial'],
                    'PrimerNombre':row['PrimerNombre'],
                    'SegundoNombre':row['SegundoNombre'],
                    'PrimerApellido':row['PrimerApellido'],
                    'SegundoApellido':row['SegundoApellido'],
                    'Sexo':row['Sexo'],
                    
                    'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                    'Tomo':row['Tomo'],
                    'Radicado':row['Radicado'],
                    
                    'Nombre Hoja': 'Propietarios'
                }
                resultados.append(resultado)
                
                print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

        df_resultado = pd.DataFrame(resultados)
        '''
        output_file = 'PRIMER_APELLIDO_BLANCO.xlsx'
        sheet_name = 'PRIMER_APELLIDO'
        df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
        print(f"Archivo guardado: {output_file}")
        messagebox.showinfo("Éxito",
                            f"Proceso completado Primer Apellido. con {len(resultados)} registros.")
        '''
        return resultados

        
        print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")
        
    def primer_nombre_blanco(self): 
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'
        df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
        if df is None:
            return
       
        resultados = []
        for index, row in df.iterrows():
            valor_a = row['PrimerNombre']
            valor_b = row['TipoDocumento']

            print(f"Fila {index}: Valor A = '{valor_a}'")

            if valor_b != '3|NIT' and (valor_a == '' or pd.isna(valor_a)):
                resultado = {
                    'NroFicha': row['NroFicha'],
                    'Observacion': 'Primer nombre en blanco',
                    'TipoDocumento':row['TipoDocumento'],
                    'Documento': row['Documento'],
                    'CalidadPropietario':row['CalidadPropietario'],
                    'Derecho':row['Derecho'],
                    'CalidadPropietarioOficial':row['CalidadPropietarioOficial'],
                    'Fecha':row['Fecha'],
                    'CodigoFideicomiso':row['CodigoFideicomiso'],
                    'Escritura':row['Escritura'],
                    'Entidad':row['Entidad'],
                    'EntidadDepartamento':row['EntidadDepartamento'],
                    'EntidadMunicipio':row['EntidadMunicipio'],
                    'NumeroFallo':row['NumeroFallo'],
                    'RazonSocial':row['RazonSocial'],
                    'PrimerNombre':row['PrimerNombre'],
                    'SegundoNombre':row['SegundoNombre'],
                    'PrimerApellido':row['PrimerApellido'],
                    'SegundoApellido':row['SegundoApellido'],
                    'Sexo':row['Sexo'],
                    
                    'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                    'Tomo':row['Tomo'],
                    'Radicado':row['Radicado'],
                    'Nombre Hoja': 'Propietarios'
                }
                resultados.append(resultado)
                
                print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
        return resultados
        '''
        df_resultado = pd.DataFrame(resultados)
        
        output_file = 'PRIMER_NOMBRE_BLANCO.xlsx'
        sheet_name = 'PRIMER_NOMBRE'
        df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
        print(f"Archivo guardado: {output_file}")
        
        
        

        messagebox.showinfo("Éxito",
                            f"Proceso completado PRIMER_NOMBRE. con {len(resultados)} registros.")
        
        print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")
        '''
        
    
            
    def derecho_diferente_cien(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
            print(f"funcion: derecho_diferente_cien")
            
            resultados = []

            # Agrupar por 'NroFicha'
            grouped = df.groupby('NroFicha')

            for name, group in grouped:
                valor_b_sum = group['Derecho'].sum()

                # Si la suma de 'Derecho' no es 100, agregar una sola observación para el grupo
                if round(valor_b_sum, 3) != 100:
                    falta_derecho = round(100 - valor_b_sum, 3)
                    radicados = ', '.join(group['Radicado'].dropna().astype(str).unique())
                    resultado = {
                        'NroFicha': name,
                        'Observacion': f'Porcentaje de derecho diferente a 100, falta: {falta_derecho}',
                        'TipoDocumento': group['TipoDocumento'].iloc[0],
                        'Documento': group['Documento'].iloc[0],
                        'Suma Derecho': valor_b_sum,
                        'Radicado':radicados,
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Resultado agregado para NroFicha {name}: {resultado}")
            '''
            
            # Agregar los resultados al consolidado solo una vez
            if resultados:
                
                print(f"Total de resultados agregados: {len(resultados)}")

            # Crear el DataFrame de resultados
            df_resultado = pd.DataFrame(resultados)
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")
            '''
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
    '''        
    def documento_blanco_cod_asig(self): COMENTADO PORQUE NO SE USA
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
                        'Observacion': 'Documento diferente a blanco para código asignado',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    # Agregar solo el resultado actual
                    
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")

            print(f"Total de resultados encontrados: {len(resultados)}")

            # Crear un nuevo DataFrame con los resultados
            df_resultado = pd.DataFrame(resultados)

            # Guardar el resultado en un nuevo archivo Excel
            
            output_file = 'DOCUMENTO_CODIGO_ASIGNADO.xlsx'
            sheet_name = 'DOCUMENTO_CODIGO_ASIGNADO'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            

            
            messagebox.showinfo("Éxito", f"Proceso completado Codigo Asignado. con {len(resultados)} registros.")
            
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")

            '''
            
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
                fecha_escritura = row['FechaEscritura']

                # Verificar si 'FechaEscritura' es un string y convertirlo si es necesario
                if isinstance(fecha_escritura, str):
                    try:
                        fecha_obj = datetime.strptime(fecha_escritura, "%d/%m/%Y")
                    except ValueError:
                        print(f"Error en el formato de fecha en la fila {index}: '{fecha_escritura}'")
                        continue  # Saltar a la siguiente fila si la conversión falla
                elif isinstance(fecha_escritura, datetime):
                    fecha_obj = fecha_escritura
                else:
                    print(f"Tipo no válido en la fila {index}: '{fecha_escritura}'")
                    continue  # Saltar a la siguiente fila si el tipo es inesperado

                # Verificar si la fecha es anterior a 1778
                if fecha_obj < fecha_umbral:
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'FechaEscritura': fecha_obj.strftime("%d/%m/%Y"),
                        'Observacion': 'Fecha de escritura inferior al año 1778',
                        'TipoDocumento':row['TipoDocumento'],
                        'Documento': row['Documento'],
                        'CalidadPropietario':row['CalidadPropietario'],
                        'Derecho':row['Derecho'],
                        'CalidadPropietarioOficial':row['CalidadPropietarioOficial'],
                        'Fecha':row['Fecha'],
                        'CodigoFideicomiso':row['CodigoFideicomiso'],
                        'Escritura':row['Escritura'],
                        'Entidad':row['Entidad'],
                        'EntidadDepartamento':row['EntidadDepartamento'],
                        'EntidadMunicipio':row['EntidadMunicipio'],
                        'NumeroFallo':row['NumeroFallo'],
                        'RazonSocial':row['RazonSocial'],
                        'PrimerNombre':row['PrimerNombre'],
                        'SegundoNombre':row['SegundoNombre'],
                        'PrimerApellido':row['PrimerApellido'],
                        'SegundoApellido':row['SegundoApellido'],
                        'Sexo':row['Sexo'],
                        
                        'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                        'Tomo':row['Tomo'],
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    
                    print(f"Fila {index}: Fecha '{fecha_obj}' es anterior a 1778. Agregado a resultados.")

            print(f"Total de fechas anteriores a 1778 encontradas: {len(resultados)}")
            '''
            
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
                                    f"Proceso completado. Fechas inferiores a 1778: {len(resultados)} registros.")
                
            else:
                print("No se encontraron fechas anteriores a 1778.")
                messagebox.showinfo("Información", "No se encontraron registros con fechas anteriores a 1778.")
            '''
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
    
    
    def fecha_escritura_mayor(self):
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

            # Obtener la fecha actual (sin tiempo para evitar diferencias por horas/minutos)
            fecha_actual = datetime.now().date()

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                fecha_escritura = row['FechaEscritura']

                # Verificar si 'FechaEscritura' es nula o NaT
                if pd.isnull(fecha_escritura):
                    print(f"Fila {index}: Fecha de escritura es nula. Ignorada.")
                    continue

                # Convertir a fecha si es necesario
                if isinstance(fecha_escritura, str):
                    try:
                        fecha_escritura = datetime.strptime(fecha_escritura, "%d/%m/%Y").date()
                    except ValueError:
                        print(f"Error en el formato de fecha en la fila {index}: '{fecha_escritura}'")
                        continue

                elif isinstance(fecha_escritura, datetime):
                    fecha_escritura = fecha_escritura.date()  # Tomar solo la parte de la fecha

                # Verificar si la fecha es superior a la fecha actual
                if fecha_escritura > fecha_actual:
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'FechaEscritura': fecha_escritura.strftime("%d/%m/%Y"),
                        'Observacion': 'Fecha de escritura es superior a la fecha actual',
                        'TipoDocumento':row['TipoDocumento'],
                        'Documento': row['Documento'],
                        'CalidadPropietario':row['CalidadPropietario'],
                        'Derecho':row['Derecho'],
                        'CalidadPropietarioOficial':row['CalidadPropietarioOficial'],
                        'Fecha':row['Fecha'],
                        'CodigoFideicomiso':row['CodigoFideicomiso'],
                        'Escritura':row['Escritura'],
                        'Entidad':row['Entidad'],
                        'EntidadDepartamento':row['EntidadDepartamento'],
                        'EntidadMunicipio':row['EntidadMunicipio'],
                        'NumeroFallo':row['NumeroFallo'],
                        'RazonSocial':row['RazonSocial'],
                        'PrimerNombre':row['PrimerNombre'],
                        'SegundoNombre':row['SegundoNombre'],
                        'PrimerApellido':row['PrimerApellido'],
                        'SegundoApellido':row['SegundoApellido'],
                        'Sexo':row['Sexo'],
                        
                        'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                        'Tomo':row['Tomo'],
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    # Solo se agrega el resultado actual, no toda la lista
                    
                    print(f"Fila {index}: Fecha '{fecha_escritura}' es superior a la fecha actual. Agregado a resultados.")

            print(f"Total de fechas superiores a la fecha actual encontradas: {len(resultados)}")
            '''
            
            if len(resultados) > 0:
                # Crear un nuevo DataFrame con los resultados
                df_resultado = pd.DataFrame(resultados)
                
                # Guardar el resultado en un nuevo archivo Excel
                output_file = 'FECHAS_ESCRITURA_SUPERIORES.xlsx'
                sheet_name = 'fechas_superiores'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

                messagebox.showinfo("Éxito", f"Proceso completado. Se ha creado el archivo '{output_file}' con {len(resultados)} registros.")
                
            else:
                print("No se encontraron fechas superiores a la fecha actual.")
                messagebox.showinfo("Información", "No se encontraron registros con fechas superiores a la fecha actual.")
            '''
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
               
               
    '''
    
    def entidadvacio(self):
        
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

            resultados = []

            

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
                
               Entidad=row['Entidad'] 
               EntidadDepartamento = row['EntidadDepartamento']
               EntidadMunicipio= row['EntidadMunicipio']
                        
               if pd.isna(EntidadDepartamento) or EntidadDepartamento=='' or EntidadMunicipio=='' or pd.isna(EntidadMunicipio) or Entidad=='':
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'EntidadDepartamento':row['EntidadDepartamento'],
                        'EntidadMunicipio':row['EntidadMunicipio'],
                        'Observacion': 'Entidad no puede ser vacia o null',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    
                    print(f"Fila {index}: Agregado a resultados: {resultado}")
                    
            print(f"Entidades vacias: {len(resultados)}")
            
            if len(resultados) > 0:
                # Crear un nuevo DataFrame con los resultados
                df_resultado = pd.DataFrame(resultados)
                
                # Guardar el resultado en un nuevo archivo Excel
                output_file = 'EntidadesVacias.xlsx'
                sheet_name = 'fechas_superiores'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

                messagebox.showinfo("Éxito", f"Proceso completado. Entidades vacias '{output_file}' con {len(resultados)} registros.")
               
            else:
                print("No se encontraron Entidades vacias.")
                messagebox.showinfo("Información", "No se encontraron registros con fechas superiores a la fecha actual.")
            
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
        '''  
    def numerofallocero(self):
        
        archivo_excel= self.archivo_entry.get()
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

            resultados = []

            

            # Iterar sobre las filas del DataFrame
            for index, row in df.iterrows():
               NumeroFallo = row['NumeroFallo']
               
                        
               if   NumeroFallo== '0' or NumeroFallo=='' or pd.isna(NumeroFallo):
                    resultado = {
                        'NroFicha': row['NroFicha'],
                        'Observacion': 'El numero fallo es cero o vacio',
                        'TipoDocumento':row['TipoDocumento'],
                        'Documento': row['Documento'],
                        'CalidadPropietario':row['CalidadPropietario'],
                        'Derecho':row['Derecho'],
                        'CalidadPropietarioOficial':row['CalidadPropietarioOficial'],
                        'Fecha':row['Fecha'],
                        'CodigoFideicomiso':row['CodigoFideicomiso'],
                        'Escritura':row['Escritura'],
                        'Entidad':row['Entidad'],
                        'EntidadDepartamento':row['EntidadDepartamento'],
                        'EntidadMunicipio':row['EntidadMunicipio'],
                        'NumeroFallo':row['NumeroFallo'],
                        'RazonSocial':row['RazonSocial'],
                        'PrimerNombre':row['PrimerNombre'],
                        'SegundoNombre':row['SegundoNombre'],
                        'PrimerApellido':row['PrimerApellido'],
                        'SegundoApellido':row['SegundoApellido'],
                        'Sexo':row['Sexo'],
                        
                        'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                        'Tomo':row['Tomo'],
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    
                    
                    
            print(f"Entidades vacias: {len(resultados)}")
            '''
            if len(resultados) > 0:
                # Crear un nuevo DataFrame con los resultados
                df_resultado = pd.DataFrame(resultados)
                
                # Guardar el resultado en un nuevo archivo Excel
                output_file = 'NumeroFallo.xlsx'
                sheet_name = 'Numero fallo'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

                messagebox.showinfo("Éxito", f"Proceso completado. Numerofallo '{output_file}' con {len(resultados)} registros.")
                
            else:
                print("No se encontraron Numerofallo.")
                messagebox.showinfo("Información", "No se encontraron registros con fechas superiores a la fecha actual.")
            '''
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
        
        
  
    
    def validar_matricula_entidad(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return

        try:
            # Leer el archivo Excel, especificando la hoja
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: validar_matricula_entidad")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            # Lista para almacenar los resultados que no cumplen la condición
            resultados = []

            # Iterar sobre cada fila del DataFrame
            for _, row in df.iterrows():
                matricula_inmobiliaria = str(row.get('MatriculaInmobiliaria', '')).strip()
                entidad_departamento = str(row.get('EntidadDepartamento', '')).strip()
                entidad_municipio = str(row.get('EntidadMunicipio', '')).strip()

                # Validar la condición
                if matricula_inmobiliaria and (entidad_departamento == 'null|null' or not entidad_municipio):
                    resultado = {
                        'NroFicha': row.get('NroFicha'),
                        'Observacion': 'EntidadDepartamento no puede ser null|null y EntidadMunicipio no puede ser vacío si MatriculaInmobiliaria tiene valor',
                        'TipoDocumento':row['TipoDocumento'],
                        'Documento': row['Documento'],
                        'CalidadPropietario':row['CalidadPropietario'],
                        'Derecho':row['Derecho'],
                        'CalidadPropietarioOficial':row['CalidadPropietarioOficial'],
                        'Fecha':row['Fecha'],
                        'CodigoFideicomiso':row['CodigoFideicomiso'],
                        'Escritura':row['Escritura'],
                        'Entidad':row['Entidad'],
                        'EntidadDepartamento':row['EntidadDepartamento'],
                        'EntidadMunicipio':row['EntidadMunicipio'],
                        'NumeroFallo':row['NumeroFallo'],
                        'RazonSocial':row['RazonSocial'],
                        'PrimerNombre':row['PrimerNombre'],
                        'SegundoNombre':row['SegundoNombre'],
                        'PrimerApellido':row['PrimerApellido'],
                        'SegundoApellido':row['SegundoApellido'],
                        'Sexo':row['Sexo'],
                        
                        'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                        'Tomo':row['Tomo'],
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    print(f"Condición de error encontrada: {resultado}")
                   
            # Generar reporte si hay resultados
            '''
            
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Validacion_Matricula_Entidad.xlsx'
                sheet_name = 'Propietarios_Errores'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo de reporte guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Se encontraron {len(resultados)} registros con errores de MatriculaInmobiliaria y Entidad.")
            else:
                messagebox.showinfo("Información", "No se encontraron errores en los registros de MatriculaInmobiliaria y Entidad.")
            '''
            
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    
    def contar_nph_calidad_propietario(self):
        archivo_excel = self.archivo_entry.get()
        hoja_fichas = 'Fichas'
        hoja_propietarios = 'Propietarios'

        if not archivo_excel or not hoja_fichas or not hoja_propietarios:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de las hojas.")
            return

        try:
            # Leer ambas hojas del archivo Excel
            df_fichas = pd.read_excel(archivo_excel, sheet_name=hoja_fichas)
            df_propietarios = pd.read_excel(archivo_excel, sheet_name=hoja_propietarios)

            print(f"funcion: validar_calidad_propietario")
            print(f"Leyendo archivo: {archivo_excel}")
            print(f"Hoja Fichas: {hoja_fichas}, Dimensiones: {df_fichas.shape}")
            print(f"Hoja Propietarios: {hoja_propietarios}, Dimensiones: {df_propietarios.shape}")

            resultados = []
            # Contador de registros que cumplen las condiciones
            

            # Iterar sobre las filas del DataFrame de Fichas
            for index, row in df_fichas.iterrows():
                npn = row.get('Npn')
                matricula = row.get('MatriculaInmobiliaria')
                nro_ficha = row.get('NroFicha')

                # Verificar condiciones en la hoja Fichas
                if pd.notna(npn) and len(str(npn)) > 21:
                    npn = str(npn)  # Convertir a string si no lo es
                    digito_22 = npn[21]

                    if digito_22 == '0' and (pd.isna(matricula) or matricula == '' or matricula == 0):
                        # Buscar el mismo NroFicha en la hoja Propietarios
                        propietarios_ficha = df_propietarios[df_propietarios['NroFicha'] == nro_ficha]

                        # Validar que CalidadPropietarioOficial no sea '4|MUNICIPAL' ni '2|NACIONAL'
                        for _, propietario in propietarios_ficha.iterrows():
                            calidad = propietario.get('CalidadPropietarioOficial')
                            matricula=propietario.get('MatriculaInmobiliaria')
                            if calidad not in ['4|MUNICIPAL', '2|NACIONAL']:
                                resultado = {
                                'NroFicha': row.get('NroFicha'),
                                'Observacion': 'El predio es NPH , la matricula es 0 o vacia y CalidadPropietarioOficial es diferente de la Nacion o el municipio',
                                'Radicado':row['Radicado'],
                                'Nombre Hoja': 'Propietarios'
                                
                                
                                }
                                resultados.append(resultado)
                        
                      # Contar solo una vez por NroFicha si cumple

           

            '''
            # Guardar los resultados en un archivo Excel
            output_file = 'VALIDACION_CALIDAD_PROPIETARIO.xlsx'
            sheet_name = 'Resumen'
            df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"Archivo guardado: {output_file}")
            print(f"Dimensiones del DataFrame de resultados: {df_resultado.shape}")

            messagebox.showinfo("Éxito",
                                f"Proceso completado. Se ha creado el archivo '{output_file}' con un resumen del total de registros.")
            '''
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
            
    def validar_tipo_documento(self):
        """
        Verifica que en la hoja 'Fichas', los valores en la columna 'TipoDocumento' no sean
        '10|CEDULA CIUDADANIA HOMBRE' o '10|CEDULA CIUDADANIA MUJER'.
        Si se encuentran estos valores, genera un error indicando que deben ser '10|CEDULA DE CIUDADANIA'.
        """
        archivo_excel = self.archivo_entry.get()
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo válido.")
            return []

        try:
            # Leer la hoja 'Fichas'
            df_fichas = pd.read_excel(archivo_excel, sheet_name='Propietarios')

            # Valores no permitidos en 'TipoDocumento'
            valores_invalidos = [
                "10|CEDULA CIUDADANIA HOMBRE",
                "10|CEDULA CIUDADANIA MUJER"
            ]
            
            # Lista para almacenar los errores encontrados
            resultados = []

            # Validar cada fila en la columna 'TipoDocumento'
            for index, row in df_fichas.iterrows():
                tipo_documento = row.get('TipoDocumento', '')
                  
                # Si 'TipoDocumento' contiene un valor no permitido
                if tipo_documento in valores_invalidos:
                    resultados.append({
                        'NroFicha': row['NroFicha'],
                        'Observacion': "Debe ser '10|CEDULA DE CIUDADANIA'",
                        'TipoDocumento':tipo_documento,
                        'Documento': row['Documento'],
                        'CalidadPropietario':row['CalidadPropietario'],
                        'Derecho':row['Derecho'],
                        'CalidadPropietarioOficial':row['CalidadPropietarioOficial'],
                        'Fecha':row['Fecha'],
                        'CodigoFideicomiso':row['CodigoFideicomiso'],
                        'Escritura':row['Escritura'],
                        'Entidad':row['Entidad'],
                        'EntidadDepartamento':row['EntidadDepartamento'],
                        'EntidadMunicipio':row['EntidadMunicipio'],
                        'NumeroFallo':row['NumeroFallo'],
                        'RazonSocial':row['RazonSocial'],
                        'PrimerNombre':row['PrimerNombre'],
                        'SegundoNombre':row['SegundoNombre'],
                        'PrimerApellido':row['PrimerApellido'],
                        'SegundoApellido':row['SegundoApellido'],
                        'Sexo':row['Sexo'],
                        
                        'MatriculaInmobiliaria':row['MatriculaInmobiliaria'],
                        'Tomo':row['Tomo'],
                        'Radicado':row['Radicado'],
                        'Nombre Hoja': 'Propietarios'
                    })
                    
            '''
            
            # Guardar los resultados en un archivo Excel si hay errores
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_TipoDocumento_Fichas.xlsx'
                df_resultado.to_excel(output_file, index=False)
                print(f"Archivo de errores guardado: {output_file}")
                messagebox.showinfo("Errores encontrados", f"Se encontraron {len(resultados)} registros con valores incorrectos en 'TipoDocumento'.")
            else:
                messagebox.showinfo("Sin errores", "No se encontraron valores incorrectos en la columna 'TipoDocumento' en la hoja 'Fichas'.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            return []
        