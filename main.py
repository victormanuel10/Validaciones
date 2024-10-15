import tkinter as tk
import warnings
from tkinter import filedialog, messagebox
from interfaz import InterfazGrafica
from validaciones.propietarios import Propietarios


warnings.filterwarnings("ignore", category=UserWarning, message="Data Validation extension is not supported and will be removed")



class Application:
    def __init__(self, root):
        self.interfaz = InterfazGrafica(root, self)

    def seleccionar_archivo_nph(self):
        self.interfaz.seleccionar_archivo_nph()

    def procesar_archivo(self):
        processor = Propietarios(self.interfaz.archivo_entry_nph)
        processor.procesar_errores()
        
    

if __name__ == "__main__":
    root = tk.Tk()
    app = Application(root)
    root.mainloop()