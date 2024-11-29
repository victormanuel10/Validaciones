# -- coding: utf-8 --
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
                
                if ( Tamaniobanio == '311|SIN BAÑO') and (pd.notna(EnchapesBanio) or pd.notna(MobiliarioBanio) or pd.notna(ConservacionBanio)):
                    resultado = {
                        'secuencia':row['secuencia'],
                        'Tamaño baño': row['TamanioBanio'],
                        'Observacion': 'No puede tener EnchapesBanio, MobiliarioBanio, ConservacionBanio (aviso) ',
                        'Nombre Hoja': nombre_hoja
                    }
                    resultados.append(resultado)
                    
                    print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
                
            '''    
            if resultados:
                df_resultado = pd.DataFrame(resultados)

                
                output_file = 'Validar_Baños.xlsx'
                sheet_name = 'Validar Baños'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")

                
                
                
                
                
                #messagebox.showinfo( f"Tamaño baño. {len(resultados)} registros.")
                
            else:
                messagebox.showinfo("Información", "No se encontraron registros que cumplan con la condición.")
            '''
            return resultados      
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
    
    def validar_sinCocina(self):
            
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
                    TamanioCocina = row['TamanioCocina']
                    Enchape= row['Enchape']
                    MobiliarioCocina = row['MobiliarioCocina']
                    ConservacionCocina = row['ConservacionCocina']
                    
                    if ( TamanioCocina == '411|SIN COCINA') and (pd.notna(Enchape) or pd.notna(MobiliarioCocina) or pd.notna(ConservacionCocina)):
                        resultado = {
                            'secuencia':row['secuencia'],
                            'Observacion': 'No puede tener Enchape, MobiliarioCocina, ConservacionCocina (aviso)',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        
                        print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
                    
                '''        
                if resultados:
                    df_resultado = pd.DataFrame(resultados)

                    
                    output_file = 'Validar_Cocinas.xlsx'
                    sheet_name = 'Validar_Cocinas'
                    df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                    print(f"Archivo guardado: {output_file}")

                    
                    
                    
                    messagebox.showinfo( f"Validar_Cocinas. {len(resultados)} registros.")
                    
                    
                    
                else:
                    messagebox.showinfo("Información", "No se encontraron registros que cumplan con la condición.")
                '''
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
                resultado = {}
                print(f"Fila {index}: {row}") 
                Armazon = row['Armazon']
                Muro = row['Muro']
                Cubierta=row['Cubierta']
                Piso=row['Piso']
                muros_validos_Madera_Tapia = ['122|BAHAREQUE,ADOBE, TAPIA', '121|MATERIALES DE DESECHOS,ESTERILLA', '123|MADERA']
                Cubierta_validas_Madera_Tapia=['131|MATERIALES DE DESECHO','132|ZINC,TEJA DE BARRO','134|ETERNIT O TEJA DE BARRO']
                
                if Armazon == '111|MADERA, TAPIA':
                    if Muro not in muros_validos_Madera_Tapia:
                        resultado = {
                            'secuencia': row['secuencia'],
                            'Armazon': row['Armazon'],
                            'Muro': row['Muro'],
                            'Cubierta':row['Cubierta'],
                            'Observacion': 'Muro invalido para armazon (aviso)',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)

                    
                    if Piso == '235|TABLETA, CAUCHO, ACRÍLICO, GRANITO, BALDOSAS FINA, CERÁMICA':
                        resultado = {
                            'secuencia': row['secuencia'],
                            'Armazon':row['Armazon'],
                            'Piso': row['Piso'],
                            'Observacion': 'Piso invalido para armazon (aviso)',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                        
                    if Cubierta not in Cubierta_validas_Madera_Tapia:
                        resultado = {
                            'secuencia': row['secuencia'],
                            'Piso': row['Piso'],
                            'Cubierta':row['Cubierta'],
                            'Observacion': 'Cubierta invalido para armazon (aviso)',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)

                #elif Armazon == '124|CONCRETO PREFABRICADO' and 
                elif Armazon == '112|PREFABRICADO':
                    # Validación para el campo Muro
                    if Muro == '122|BAHAREQUE,ADOBE, TAPIA' or Muro == '121|MATERIALES DE DESECHOS,ESTERILLA':
                        resultado = {
                            'secuencia': row['secuencia'],
                            'Armazon': row['Armazon'],
                            'Muro': row['Muro'],
                            'Cubierta': row['Cubierta'],
                            'Piso': row['Piso'],
                            'Observacion': 'Muro invalido para armazon (aviso)',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                    
                    # Validación para el campo Cubierta
                    cubiertas_invalidas_prefabricado = ['121|MATERIALES DE DESECHOS', '135|AZOTEA, ALUMINIO,PLACAS CON ETERNIT', '136|PLACA IMPERMEABILI, CUBIERTA DE LUJO']
                    if Cubierta in cubiertas_invalidas_prefabricado:
                        resultado = {
                            'secuencia': row['secuencia'],
                            'Armazon': row['Armazon'],
                            'Muro': row['Muro'],
                            'Cubierta': row['Cubierta'],
                            'Piso': row['Piso'],
                            'Observacion': 'Cubierta invalida para armazon (aviso)',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                
                
                
                
                elif Armazon == '113|LADRILLO,BLOQUE, MADERA INMUNIZADA':
                    if Muro=='122|BAHAREQUE,ADOBE, TAPIA':
                        resultado = {
                            'secuencia': row['secuencia'],
                            'Armazon': row['Armazon'],
                            'Muro': row['Muro'],
                            'Cubierta':row['Cubierta'],
                            'Piso':row['Piso'],
                            'Observacion': 'Muro invalido para armazon (aviso)',
                            'Nombre Hoja': nombre_hoja
                        }

                        resultados.append(resultado)

                    if Cubierta=='131|MATERIALES DE DESECHO':
                        resultado = {
                            'secuencia': row['secuencia'],
                            'Armazon': row['Armazon'],
                            'Muro': row['Muro'],
                            'Cubierta':row['Cubierta'],
                            'Piso':row['Piso'],
                            'Observacion': 'Cubierta invalido para armazon (aviso)',
                            'Nombre Hoja': nombre_hoja
                        }
                    
                        resultados.append(resultado)
                    
                elif Armazon == '114|CONCRETO HASTA TRES PISOS':
                    if (Muro=='121|MATERIALES DE DESECHOS,ESTERILLA' or Muro=='123|MADERA' or Muro=='122|BAHAREQUE,ADOBE, TAPIA'):
                        resultado = {
                            'secuencia': row['secuencia'],
                            'Armazon': row['Armazon'],
                            'Muro': row['Muro'],
                            'Cubierta':row['Cubierta'],
                            'Observacion': 'Muro invalido para armazon (aviso)',
                            'Nombre Hoja': nombre_hoja
                        }
                    
                        resultados.append(resultado)
                    if (Cubierta=='131|MATERIALES DE DESECHO'):
                        resultado = {
                            'secuencia': row['secuencia'],
                            'Armazon': row['Armazon'],
                            'Muro': row['Muro'],
                            'Cubierta':row['Cubierta'],
                            'Observacion': 'Cubierta invalido para armazon(aviso)',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
            '''
            
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                
                
                
                output_file = 'MuroInvalido.xlsx'
                sheet_name = 'MuroInvalido'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")

                messagebox.showinfo("Éxito", f"Proceso completado. Muro Invalido con {len(resultados)} registros.")
                
            else:
                messagebox.showinfo("Información", "No se encontraron registros que cumplan con la condición.")
            '''
            return resultados
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")
            
    def Validar_fachada(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'CalificacionesConstrucciones'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return
        
        try:
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

            print(f"funcion: Validar_fachada")
            print(f"Leyendo archivo: {archivo_excel}, Hoja: {nombre_hoja}")
            print(f"Dimensiones del DataFrame: {df.shape}")
            print(f"Columnas en el DataFrame: {df.columns.tolist()}")

            resultados = []

            for index, row in df.iterrows():
                resultado = {}
                print(f"Fila {index}: {row}") 
                Fachada = row['Fachada']
                Cubrimiento_Muro = row['Cubrimiento Muro']
                Piso = row['Piso']

                # Lista de valores no permitidos en Cubrimiento Muro cuando Fachada es igual a 211|POBRE
                cubrimiento_invalidos = ['223|ESTUCO, CERÁMICA, PAPEL FINO', 
                                        '224|MADERA, PIEDRA ORNAMENT. LADRILLO FINO', 
                                        '225|MÁRMOL, LUJOSOS, OTROS']

                # Lista de valores no permitidos en Piso cuando Fachada es igual a 211|POBRE
                pisos_invalidos = ['235|TABLETA, CAUCHO, ACRÍLICO, GRANITO, BALDOSAS FINA, CERÁMICA',
                                '236|PARQUET, ALFONFRA, RETAL DE MÁRMOL',
                                '237|MÁRMOL, OTROS LUJOSOS']
                
                cubrimiento_invalidos_sencilla=['224|MADERA, PIEDRA ORNAMENT. LADRILLO FINO', 
                                        '225|MÁRMOL, LUJOSOS, OTROS']
                
                pisos_invalidos_sencilla=['236|PARQUET, ALFONFRA, RETAL DE MÁRMOL',
                                          '237|MÁRMOL, OTROS LUJOSOS']
                
                cubrimiento_invalidos_regular=['225|MÁRMOL, LUJOSOS, OTROS']
                
                pisos_invalidos_regular=['231|TIERRA PISADA','236|PARQUET, ALFONFRA, RETAL DE MÁRMOL'
                                         ,'237|MÁRMOL, OTROS LUJOSOS']
                
                cubrimiento_invalidos_bueno=['221|SIN CUBRIMIENTO','222|PAÑETE, PAPEL, COMÚN, LADRILLO PRENSADO']
                pisos_invalidos_bueno=['231|TIERRA PISADA','232|CEMENTO, MADERA BURDA',
                                        '233|BALDOSA COMÚN DE CEMENTO, TABLÓN LADR']
                
                if Fachada == '211|POBRE':
                    # Validación para Cubrimiento Muro
                    if Cubrimiento_Muro in cubrimiento_invalidos:
                        resultado = {
                            'secuencia': row['secuencia'],
                            'Fachada': row['Fachada'],
                            'Cubrimiento Muro': row['Cubrimiento Muro'],
                            'Observacion': 'Cubrimiento Muro invalido para fachada (aviso)',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)

                    # Validación para Piso
                    if Piso in pisos_invalidos:
                        resultado = {
                            'secuencia': row['secuencia'],
                            'Fachada': row['Fachada'],
                            'Piso': row['Piso'],
                            'Observacion': 'Piso invalido para fachada (aviso)',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                
                elif Fachada == '212|SENCILLA':
                    if Cubrimiento_Muro in cubrimiento_invalidos_sencilla:
                        resultado = {
                            'secuencia': row['secuencia'],
                            'Fachada': row['Fachada'],
                            'Cubrimiento Muro': row['Cubrimiento Muro'],
                            'Observacion': 'Cubrimiento Muro invalido para fachada (aviso)',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                    
                    if Piso in pisos_invalidos_sencilla:
                        resultado = {
                            'secuencia': row['secuencia'],
                            'Fachada': row['Fachada'],
                            'Piso': row['Piso'],
                            'Observacion': 'Piso invalido para fachada (aviso)',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                
                
                elif Fachada == '213|REGULAR':
                    if Cubrimiento_Muro in cubrimiento_invalidos_regular:
                        resultado = {
                            'secuencia': row['secuencia'],
                            'Fachada': row['Fachada'],
                            'Cubrimiento Muro': row['Cubrimiento Muro'],
                            'Observacion': 'Cubrimiento Muro invalido para fachada (aviso)',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                    
                    if Piso in pisos_invalidos_regular:
                        resultado = {
                            'secuencia': row['secuencia'],
                            'Fachada': row['Fachada'],
                            'Piso': row['Piso'],
                            'Observacion': 'Piso invalido para fachada (aviso)',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                    
                elif Fachada == '214|BUENA':
                    if Cubrimiento_Muro in cubrimiento_invalidos_bueno:
                        resultado = {
                            'secuencia': row['secuencia'],
                            'Fachada': row['Fachada'],
                            'Cubrimiento Muro': row['Cubrimiento Muro'],
                            'Observacion': 'Cubrimiento Muro invalido para fachada (aviso)',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                    
                    if Piso in pisos_invalidos_bueno:
                        resultado = {
                            'secuencia': row['secuencia'],
                            'Fachada': row['Fachada'],
                            'Piso': row['Piso'],
                            'Observacion': 'Piso invalido para fachada (aviso)',
                            'Nombre Hoja': nombre_hoja
                        }
                        resultados.append(resultado)
                print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
            '''
            
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                
                output_file = 'CubrimientoMuroInvalido.xlsx'
                sheet_name = 'CubrimientoMuroInvalido'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")

                messagebox.showinfo("Éxito", f"Proceso completado. Cubrimiento Muro invalido con {len(resultados)} registros.")
            else:
                messagebox.showinfo("Información", "No se encontraron registros que cumplan con la condición.")
            '''
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
                    secuencia = row['secuencia']

                    
                    construccion_row = df_construcciones[df_construcciones['secuencia'] == secuencia]

                    if not construccion_row.empty and construccion_row.iloc[0]['EdadConstruccion'] >= 20:
                        resultado = {
                            'secuencia': secuencia,
                            'Conservacion': conservacion,
                            'Observacion': 'La edad de la construcción es mayor o igual a 20 años (aviso)',
                            'Nombre Hoja': nombre_hoja_calificaciones
                        }
                        resultados.append(resultado)
                        print(f"Fila {index} cumple las condiciones. Agregado: {resultado}")
            '''
            if resultados:
                
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_Construcciones.xlsx'
                sheet_name = 'Errores'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Proceso completado. Se encontraron {len(resultados)} registros con errores.")
            
                
            else:
                messagebox.showinfo("Información", "No se encontraron registros que cumplan con las condiciones.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")

    def validar_cubierta_y_numero_pisos(self):
        archivo_excel = self.archivo_entry.get()
        hoja_calificaciones = 'CalificacionesConstrucciones'
        hoja_construcciones = 'Construcciones'

        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return

        try:
            # Leer ambas hojas del archivo Excel
            df_calificaciones = pd.read_excel(archivo_excel, sheet_name=hoja_calificaciones)
            df_construcciones = pd.read_excel(archivo_excel, sheet_name=hoja_construcciones)

            print(f"funcion: validar_cubierta_numero_pisos")
            print(f"Leyendo archivo: {archivo_excel}")
            print(f"Dimensiones Hoja Calificaciones: {df_calificaciones.shape}")
            print(f"Dimensiones Hoja Construcciones: {df_construcciones.shape}")

            # Lista para almacenar los resultados
            resultados = []

            # Filtrar las filas de CalificacionesConstrucciones donde Cubierta sea igual al valor especificado
            calificaciones_filtradas = df_calificaciones[df_calificaciones['Cubierta'] == '135|AZOTEA, ALUMINIO,PLACAS CON ETERNIT']

            for _, fila_calificaciones in calificaciones_filtradas.iterrows():
                secuencia = fila_calificaciones['secuencia']

                # Buscar la misma secuencia en la hoja Construcciones
                construccion_filtrada = df_construcciones[df_construcciones['secuencia'] == secuencia]

                if not construccion_filtrada.empty:
                    numero_pisos = construccion_filtrada.iloc[0]['NumeroPisos']

                    # Validar que el Número de Pisos sea menor a 3
                    if numero_pisos < 3:
                        resultado = {
                            'secuencia': secuencia,
                            'Cubierta': fila_calificaciones['Cubierta'],
                            'NumeroPisos': numero_pisos,
                            'Observacion': 'Número de pisos menor a 3 para la cubierta azotea (aviso)',
                            'Nombre Hoja': hoja_calificaciones
                            
                        }
                        resultados.append(resultado)
                        print(f"Error encontrado: {resultado}")
            '''
            
            # Manejar resultados
            if resultados:
                df_resultado = pd.DataFrame(resultados)
                output_file = 'Errores_Cubierta_NumeroPisos.xlsx'
                sheet_name = 'Errores'
                df_resultado.to_excel(output_file, sheet_name=sheet_name, index=False)
                print(f"Archivo guardado: {output_file}")
                messagebox.showinfo("Éxito", f"Se encontraron {len(resultados)} registros con errores en Cubierta y Número de Pisos.")
            else:
                messagebox.showinfo("Información", "No se encontraron errores en Cubierta y Número de Pisos.")
            '''
            return resultados

        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error durante el proceso: {str(e)}")