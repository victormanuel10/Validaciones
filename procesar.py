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
        
        
        
        
        
       
        
        
        
        ficha = Ficha(self.archivo_entry)
        
        
        
        
        self.agregar_resultados(ficha.validar_area_construida())
        self.agregar_resultados(ficha.modo_adquisicion_informal())
        self.agregar_resultados(ficha.informal_matricula())
        self.agregar_resultados(ficha.validar_destino_economico_nulo_o_0na())
        self.agregar_resultados(ficha.areaterrenocero())
        self.agregar_resultados(ficha.tomo_mejora())
        self.agregar_resultados(ficha.validar_modo_adquisicion_caracteristica())
        self.agregar_resultados(ficha.validar_fichas_en_propietarios())
        self.agregar_resultados(ficha.validar_nrofichas_propietarios())
        self.agregar_resultados(ficha.validar_matricula_repetida())
        self.agregar_resultados(ficha.validar_matricula_numerica())
        self.agregar_resultados(ficha.predios_con_direcciones_invalidas())
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
        self.agregar_resultados(propietarios.entidadvacio())
        self.agregar_resultados(propietarios.primer_apellido_blanco())
        self.agregar_resultados(propietarios.primer_nombre_blanco())
       #self.agregar_resultados(propietarios.documento_blanco_cod_asig())
        self.agregar_resultados(propietarios.fecha_escritura_inferior())
        self.agregar_resultados(propietarios.fecha_escritura_mayor())
        
        fichasrph=FichasRPH(self.archivo_entry)
        self.agregar_resultados(fichasrph.validar_area_privada())
        self.agregar_resultados(fichasrph.validar_area_comun())
        self.agregar_resultados(fichasrph.edificio_en_cero_rph())
        self.agregar_resultados(fichasrph.validar_informalidad_con_piso())
        self.agregar_resultados(fichasrph.validar_informalidad_edificio())
        self.agregar_resultados(fichasrph.validar_area_total_lote_npn())
        self.agregar_resultados(fichasrph.validar_unidades_rph())
        self.agregar_resultados(fichasrph.validar_npn_num_cedula())
        self.agregar_resultados(fichasrph.validar_npn_suma_cero_unico())
        self.agregar_resultados(fichasrph.validar_duplicados_npn())
        self.agregar_resultados(fichasrph.validar_coeficiente_copropiedad_por_npn())
        
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
        self.agregar_resultados(calificonstrucciones.validar_cubierta_y_numero_pisos())
        self.agregar_resultados(calificonstrucciones.validar_sinCocina())
        self.agregar_resultados(calificonstrucciones.conservacion_cubierta_bueno())
        self.agregar_resultados(calificonstrucciones.validar_banios()) 
        self.agregar_resultados(calificonstrucciones.Validar_armazon())
        self.agregar_resultados(calificonstrucciones.Validar_fachada())
        
        
        
        cartografia=Cartografia(self.archivo_entry)
        self.agregar_resultados(cartografia.validar_fichas_faltantes())
        self.agregar_resultados(cartografia.validar_cartografia_faltantes())
        self.agregar_resultados(cartografia.validar_cartografia_columnas())
        
        
        colindantes=Colindantes(self.archivo_entry)
        self.agregar_resultados(colindantes.validar_orientaciones_rph())
        self.agregar_resultados(colindantes.validar_orientaciones_colindantes())
        
        
        zonashomogeneas= ZonasHomogeneas(self.archivo_entry)
        self.agregar_resultados(zonashomogeneas.validar_tipo_zonas_homogeneas())
        
        reportes=Reportes(self.archivo_entry)
        self.agregar_resultados(reportes.matriz_con_matricula())
        self.agregar_resultados(reportes.matriz_sin_matricula())
        self.agregar_resultados(reportes.matriz_sin_circulo())
        self.agregar_resultados(reportes.matriz_con_circulo())
        self.agregar_resultados(reportes.contar_rph_matriz())
        self.agregar_resultados(reportes.contar_unidades_prediales())
        self.agregar_resultados(reportes.contar_nph())
        
        
        self.generar_reporte_observaciones()  
        
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
    
            messagebox.showinfo("Éxito", "Proceso completado. Se ha creado el archivo 'ERRORES_CONSOLIDADOS.xlsx'.")
        else:
            messagebox.showinfo("Sin errores", "No se encontraron errores en los archivos procesados.")
    
    
    
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
    
    
    
    def generar_reporte_observaciones(self):
        """
        Genera un reporte con el conteo de las observaciones y lo guarda en la hoja 'Reporte'.
        """
        if not self.resultados_generales:
            print("No hay resultados generales para generar el reporte.")
            return  # Termina la función si no hay resultados

        # Verifica la estructura de los datos
        print("Estructura de resultados generales:")
        print(self.resultados_generales)  # Imprime los resultados para ver qué contiene

        # Asegúrate de que todos los diccionarios tengan la clave 'Observacion'
        for resultado in self.resultados_generales:
            if 'Observacion' not in resultado:
                print("Falta la clave 'Observacion' en uno de los resultados:", resultado)

        # Filtra los resultados que contienen la clave 'Observacion'
        contador_observaciones = Counter([resultado['Observacion'] for resultado in self.resultados_generales if 'Observacion' in resultado])

        # Crear el DataFrame con el conteo
        df_reporte = pd.DataFrame(contador_observaciones.items(), columns=['Observacion', 'Cantidad'])

        # Agregar la hoja 'Reporte' al archivo Excel con los errores
        self.reporte = df_reporte

        # Verificación
        print("Reporte generado:")
        print(self.reporte)  # Esto debería mostrar el DataFrame con las observaciones

    def agregar_reporte(self, writer):
        """
        Agrega la hoja 'Reporte' con el conteo de observaciones al archivo Excel.
        """
        if hasattr(self, 'reporte'):
            
            self.reporte.to_excel(writer, sheet_name='Resumen', index=False)
            print("Reporte de observaciones agregado a la hoja 'Reporte'.")
        else:
            print("No hay observaciones para generar el reporte.")