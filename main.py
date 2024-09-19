import tkinter as tk
import warnings
from interfaz import InterfazGrafica
from validaciones.propietarios import Propietarios


warnings.filterwarnings("ignore", category=UserWarning, message="Data Validation extension is not supported and will be removed")



class Application:
    def __init__(self, root):
        self.interfaz = InterfazGrafica(root, self)

    def seleccionar_archivo(self):
        self.interfaz.seleccionar_archivo()

    def procesar_archivo(self):
        processor = Propietarios(self.interfaz.archivo_entry)
        processor.procesar()


if __name__ == "__main__":
    root = tk.Tk()
    app = Application(root)
    root.mainloop()