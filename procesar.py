import pandas as pd
from tkinter import messagebox
from collections import Counter
from datetime import datetime
from validaciones.ficha import Ficha
from validaciones.construcciones import Construcciones
from validaciones.califconstrucciones import CalificaionesConstrucciones
from validaciones.zonashomogeneas import ZonasHomogeneas
from validaciones.colindantes import Colindantes
from validaciones.cartografia import Cartografia
from ValidacionesRPH.fichasrph import FichasRPH
from validaciones.propietarios import Propietarios
from reportes import Reportes

class Procesar:
    
    def __init__(self, archivo_entry):
        self.archivo_entry = archivo_entry
        self.resultados_generales = []
        
    def agregar_resultados(self, resultados):
        if isinstance(resultados, list):
            for resultado in resultados:
                self.resultados_generales.append(resultado)
        elif isinstance(resultados, pd.DataFrame):
            self.resultados_generales.extend(resultados.to_dict(orient='records'))
       
    def procesar_errores(self):

        
        
        
        
        archivo_excel = self.archivo_entry.get()
        
        if not archivo_excel:
            messagebox.showerror("Error", "Por favor, selecciona un archivo.")
            return
        
        ficha = Ficha(self.archivo_entry)
        
        self.agregar_resultados(ficha.validar_fichas_en_propietarios())
        self.agregar_resultados(ficha.validar_destino_economico_nulo_o_0na())
        
        self.agregar_resultados(ficha.predios_con_direcciones_invalidas())
        
        
        self.agregar_resultados(ficha.validar_modo_adquisicion_caracteristica())
        
        self.agregar_resultados(ficha.validar_caracteristica_predio())
        self.agregar_resultados(ficha.validar_agricola_urb())
        self.agregar_resultados(ficha.validar_matricula_repetida())
        self.agregar_resultados(ficha.validar_area_construida_fichas_construcciones())
        self.agregar_resultados(ficha.modo_adquisicion_informal())
        self.agregar_resultados(ficha.informal_matricula())
        
        self.agregar_resultados(ficha.areaterrenocero())
        self.agregar_resultados(ficha.tomo_mejora())
        
        
        self.agregar_resultados(ficha.validar_nrofichas_propietarios())
        self.agregar_resultados(ficha.validar_matricula_numerica())
        
        self.agregar_resultados(ficha.validar_duplicados_npn())
        self.agregar_resultados(ficha.validar_matricula_inmobiliaria_PredioLc_Modo_Adquisicion())
        self.agregar_resultados(ficha.validar_npn_sin_cuatro_ceros())
        self.agregar_resultados(ficha.validar_Predios_Uso_Publico())
        self.agregar_resultados(ficha.Validar_Longitud_NPN())
        self.agregar_resultados(ficha.validar_direccion_referencia_y_nombre())
        self.agregar_resultados(ficha.validar_destino_economico_y_longitud_cedula())
        self.agregar_resultados(ficha.ultimo_digito())
        
        
        self.agregar_resultados(ficha.validar_npn14a17())
        self.agregar_resultados(ficha.validar_npn())
        self.agregar_resultados(ficha.porcentaje_litigiocero())
        self.agregar_resultados(ficha.destino_economico_mayorcero())
        self.agregar_resultados(ficha.matricula_mejora())
        self.agregar_resultados(ficha.terreno_cero())
        self.agregar_resultados(ficha.terreno_null())
        self.agregar_resultados(ficha.circulo_mejora())
        self.agregar_resultados(ficha.ficha_repetida())
        self.agregar_resultados(ficha.validar_matricula_no_inicia_cero())
        
        
        
        propietarios = Propietarios(self.archivo_entry)
        self.agregar_resultados(propietarios.validar_tipo_documento())
        self.agregar_resultados(propietarios.contar_nph_calidad_propietario())
        self.agregar_resultados(propietarios.validar_matricula_entidad())
        self.agregar_resultados(propietarios.derecho_diferente_cien())
        self.agregar_resultados(propietarios.validar_documento_inicia_con_cero())
        self.agregar_resultados(propietarios.validar_documento_sexo_masculino())
        self.agregar_resultados(propietarios.validar_tipo_documento_sexo())
        self.agregar_resultados(propietarios.validar_documento_sexo_femenino())
        self.agregar_resultados(propietarios.numerofallocero())
        self.agregar_resultados(propietarios.primer_apellido_blanco())
        self.agregar_resultados(propietarios.primer_nombre_blanco())
       #self.agregar_resultados(propietarios.documento_blanco_cod_asig())
        self.agregar_resultados(propietarios.fecha_escritura_inferior())
        self.agregar_resultados(propietarios.fecha_escritura_mayor())
        
        fichasrph=FichasRPH(self.archivo_entry)
        self.agregar_resultados(fichasrph.piso_en_cero_rph())
        self.agregar_resultados(fichasrph.validar_coeficiente_copropiedad_por_npn())
        
        self.agregar_resultados(fichasrph.edificio_en_cero_rph())
        self.agregar_resultados(fichasrph.validar_informalidad_con_piso())
        self.agregar_resultados(fichasrph.validar_informalidad_edificio())
        self.agregar_resultados(fichasrph.validar_area_total_lote_npn())
        self.agregar_resultados(fichasrph.validar_unidades_rph())
        self.agregar_resultados(fichasrph.validar_npn_num_cedula())
        self.agregar_resultados(fichasrph.validar_npn_suma_cero_unico())
        self.agregar_resultados(fichasrph.validar_duplicados_npn())
        
        
        construcciones = Construcciones(self.archivo_entry)
        self.agregar_resultados(construcciones.validar_secuencia_construcciones_vs_calificaciones())
        self.agregar_resultados(construcciones.validar_secuencia_convencional())
        self.agregar_resultados(construcciones.validar_secuencia_unica_por_ficha())
        self.agregar_resultados(construcciones.validar_construcciones_No_convencionales())
        self.agregar_resultados(construcciones.validar_secuencia_convencional_calificaciones())
        self.agregar_resultados(construcciones.validar_no_convencional_secuencia())       
        self.agregar_resultados(construcciones.validar_construcciones_puntos())
        self.agregar_resultados(construcciones.validar_porcentaje_construido())
        self.agregar_resultados(construcciones.validar_edad_construccion())
        self.agregar_resultados(construcciones.validar_construcciones_No_convencionales())
        self.agregar_resultados(construcciones.areaconstruida_mayora1000())
        self.agregar_resultados(construcciones.tipo_construccion_noconvencionales())         
        self.agregar_resultados(construcciones.validar_secuencia_calificaciones_vs_construcciones())
        
        calificonstrucciones= CalificaionesConstrucciones(self.archivo_entry)
        self.agregar_resultados(calificonstrucciones.validar_banios())
        self.agregar_resultados(calificonstrucciones.validar_cubierta_y_numero_pisos())
        self.agregar_resultados(calificonstrucciones.validar_sinCocina())
        self.agregar_resultados(calificonstrucciones.Validar_armazon())
        self.agregar_resultados(calificonstrucciones.Validar_fachada())
        self.agregar_resultados(calificonstrucciones.conservacion_cubierta_bueno())
        
        
        colindantes=Colindantes(self.archivo_entry)
        self.agregar_resultados(colindantes.validar_orientaciones_rph())
        self.agregar_resultados(colindantes.validar_orientaciones_colindantes())
        
        cartografia=Cartografia(self.archivo_entry)
        self.agregar_resultados(cartografia.validar_fichas_faltantes())
        self.agregar_resultados(cartografia.validar_cartografia_faltantes())
        self.agregar_resultados(cartografia.validar_cartografia_columnas())
        
        
        
        
        zonashomogeneas= ZonasHomogeneas(self.archivo_entry)
        self.agregar_resultados(zonashomogeneas.validar_tipo_zonas_homogeneas())
        
        
        self.generar_reporte_observaciones(archivo_excel)  
        
        errores_por_hoja = {}

        if self.resultados_generales:
            for resultado in self.resultados_generales:
                nombre_hoja = resultado.get('Nombre Hoja', 'Sin Nombre')
                if nombre_hoja not in errores_por_hoja:
                    errores_por_hoja[nombre_hoja] = []
                errores_por_hoja[nombre_hoja].append(resultado)

            with pd.ExcelWriter('ERRORES_CONSOLIDADOS.xlsx') as writer:
                for hoja, errores in errores_por_hoja.items():
                    df_resultado = pd.DataFrame(errores)
                    df_resultado.to_excel(writer, sheet_name=hoja, index=False)
                    print(f"Errores guardados en la hoja: {hoja}")
                
                # Asegúrate de que el reporte se agrega al archivo después de los errores
                self.agregar_reporte(writer)
                self.agregar_hoja_reportes(writer)
            messagebox.showinfo("Éxito", "Proceso completado. Se ha creado el archivo 'ERRORES_CONSOLIDADOS.xlsx'.")
        else:
            messagebox.showinfo("Sin errores", "No se encontraron errores en los archivos procesados.")
    
    def agregar_hoja_reportes(self, writer):
        """
        Genera y agrega una hoja llamada 'Reportes' al archivo Excel con las funciones de la clase Reportes.
        """
        try:
            # Instanciar la clase Reportes
            reportes = Reportes(self.archivo_entry)

            # Obtener resultados de las funciones de la clase Reportes
            resultados_reportes = []
            funciones_reportes = [
                reportes.matriz_con_matricula,
                reportes.matriz_sin_matricula,
                reportes.matriz_sin_circulo,
                reportes.matriz_con_circulo,
                reportes.contar_rph_matriz,
                reportes.contar_unidades_prediales,
                reportes.contar_nph
            ]

            # Ejecutar cada función y agregar resultados
            for funcion in funciones_reportes:
                resultado = funcion()
                if isinstance(resultado, pd.DataFrame):
                    # Concatenar los resultados
                    resultados_reportes.append(resultado)

            # Concatenar todos los DataFrames en uno solo
            if resultados_reportes:
                df_reportes = pd.concat(resultados_reportes, ignore_index=True)
                df_reportes.to_excel(writer, sheet_name='Reportes', index=False)
                print("Hoja 'Reportes' agregada con las funciones de la clase Reportes.")
            else:
                print("No se generaron resultados para la hoja 'Reportes'.")

        except Exception as e:
            print(f"Error al generar la hoja 'Reportes': {e}")
            
    
    def leer_archivo(self):
        archivo_excel = self.archivo_entry.get()
        nombre_hoja = 'Propietarios'

        if not archivo_excel or not nombre_hoja:
            messagebox.showerror("Error", "Por favor, selecciona un archivo y especifica el nombre de la hoja.")
            return None

        try:
            df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)
            return df
        except Exception as e:
            print(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error al leer el archivo: {str(e)}")
            return None
    
    
    
    def generar_reporte_observaciones(self, archivo_excel):
        """
        Genera un reporte con:
        1. El conteo de las observaciones y lo guarda en la hoja 'Resumen'.
        2. Una agrupación por cada 'NroFicha' con las observaciones asociadas y lo guarda en otra hoja,
        incluyendo la columna Npn de la hoja Fichas.
        """
        if not self.resultados_generales:
            print("No hay resultados generales para generar el reporte.")
            return  # Termina la función si no hay resultados

        try:
            # --- Reporte 1: Conteo de Observaciones ---
            contador_observaciones = Counter([resultado['Observacion'] for resultado in self.resultados_generales if 'Observacion' in resultado])

            # Crear el DataFrame con el conteo
            df_reporte = pd.DataFrame(contador_observaciones.items(), columns=['Observacion', 'Cantidad'])

            # Almacenar el reporte de observaciones
            self.reporte = df_reporte

            # --- Reporte 2: Agrupación por NroFicha ---
            # Crear un DataFrame de los resultados generales
            df_resultados = pd.DataFrame(self.resultados_generales)

            # Verificar si las columnas necesarias existen
            if 'NroFicha' in df_resultados.columns and 'Observacion' in df_resultados.columns:
                # Agrupar observaciones por NroFicha
                agrupacion_fichas = (
                    df_resultados.groupby('NroFicha')['Observacion']
                    .apply(lambda x: '; '.join(map(str, x.unique())))  # Convertir cada valor a cadena
                    .reset_index()
                )
                agrupacion_fichas.columns = ['NroFicha', 'Observaciones']

                # Leer la hoja Fichas del archivo Excel
                df_fichas = pd.read_excel(archivo_excel, sheet_name='Fichas')

                # Convertir NroFicha a numérico en ambas tablas
                agrupacion_fichas['NroFicha'] = pd.to_numeric(agrupacion_fichas['NroFicha'], errors='coerce')
                df_fichas['NroFicha'] = pd.to_numeric(df_fichas['NroFicha'], errors='coerce')

                # Realizar el merge para agregar la columna Npn
                self.agrupacion_fichas = pd.merge(
                    agrupacion_fichas,
                    df_fichas[['NroFicha', 'Npn','Radicado']],
                    on='NroFicha',
                    how='left'
                )
            else:
                print("No se encontraron las columnas 'NroFicha' o 'Observacion' en los resultados.")
                self.agrupacion_fichas = None

            # Verificación
            print("Reporte generado:")
            print(self.reporte)  # Esto debería mostrar el DataFrame con las observaciones
            print("Agrupación por NroFicha:")
            print(self.agrupacion_fichas)  # Esto debería mostrar la agrupación por NroFicha con la columna Npn

        except Exception as e:
            print(f"Error al generar el reporte: {e}")

    def agregar_reporte(self, writer):
        """
        Agrega las hojas 'Resumen' y 'Agrupación por Fichas' al archivo Excel.
        """
        if hasattr(self, 'reporte'):
            self.reporte.to_excel(writer, sheet_name='Resumen', index=False)
            print("Reporte de observaciones agregado a la hoja 'Resumen'.")
        else:
            print("No hay observaciones para generar el reporte.")

        if hasattr(self, 'agrupacion_fichas') and self.agrupacion_fichas is not None:
            self.agrupacion_fichas.to_excel(writer, sheet_name='Errores por Ficha', index=False)
            print("Agrupación por NroFicha agregada a la hoja 'Agrupación por Fichas'.")
        else:
            print("No hay agrupación por NroFicha para generar el reporte.")