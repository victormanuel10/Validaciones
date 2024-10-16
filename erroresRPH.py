import pandas as pd
from tkinter import messagebox
from validaciones.fichasrph import FichasRPH

class FichasRPHProcesador:
    def __init__(self, archivo_entry):
        self.archivo_entry = archivo_entry
        self.resultados_generales = []

    def agregar_resultados(self, resultados):
        """ Agrega los resultados de validaciones a la lista general. """
        if isinstance(resultados, list):
            for resultado in resultados:
                self.resultados_generales.append(resultado)
        elif isinstance(resultados, pd.DataFrame):
            self.resultados_generales.extend(resultados.to_dict(orient='records'))

    def procesar_errores_rph(self):
        """ Ejecuta todas las validaciones específicas de RPH y genera un archivo consolidado. """
        
        
        fichas_rph = FichasRPH(self.archivo_entry)
        
        self.agregar_resultados(fichas_rph.validar_coeficiente_copropiedad())
        # Puedes agregar más validaciones si es necesario, como otras reglas dentro de FichasRPH

        # Verificar y generar el archivo de errores consolidado
        self.generar_archivo_errores()

    def generar_archivo_errores(self):
        """ Genera un archivo Excel con los errores recopilados. """
        errores_por_hoja = {}

        if self.resultados_generales:
            for resultado in self.resultados_generales:
                nombre_hoja = resultado.get('Nombre Hoja', 'Sin Nombre')  # Obtener el nombre de la hoja
                if nombre_hoja not in errores_por_hoja:
                    errores_por_hoja[nombre_hoja] = []  # Inicializa la lista para esa hoja
                errores_por_hoja[nombre_hoja].append(resultado)

            # Crear un archivo Excel con múltiples hojas
            with pd.ExcelWriter('ERRORES_FICHAS_RPH_CONSOLIDADOS.xlsx') as writer:
                for hoja, errores in errores_por_hoja.items():
                    df_resultado = pd.DataFrame(errores)
                    df_resultado.to_excel(writer, sheet_name=hoja, index=False)
                    print(f"Errores guardados en la hoja: {hoja}")

            messagebox.showinfo("Éxito", "Proceso completado. Se ha creado el archivo 'ERRORES_FICHAS_RPH_CONSOLIDADOS.xlsx'.")
        else:
            messagebox.showinfo("Sin errores", "No se encontraron errores en los archivos procesados.")
