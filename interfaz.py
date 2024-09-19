import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk

class InterfazGrafica:
    def __init__(self, root, app):
        self.root = root
        self.app = app
        self.root.title("Procesador de Excel")
        self.root.geometry("800x400")
        self.root.state('zoomed')

        # Cargar la imagen y configurarla como fondo
        self.background_image = self.crear_imagen_semitransparente("./assets/Logo_Conestudios.png", 0.1)
        self.background_label = tk.Label(self.root, image=self.background_image)
        self.background_label.place(x=0, y=0, relwidth=1, relheight=1)

        # Crear y colocar widgets
        self.crear_widgets()

    def crear_imagen_semitransparente(self, image_path, alpha):
        # Cargar la imagen original
        image = Image.open(image_path)
        image = image.convert("RGBA")
        data = image.getdata()

        # Modificar el canal alpha de cada píxel
        new_data = [(item[:-1] + (int(255 * alpha),)) for item in data]
        image.putdata(new_data)

        # Redimensionar la imagen
        image = image.resize((self.root.winfo_screenwidth(), self.root.winfo_screenheight()), Image.LANCZOS)
        return ImageTk.PhotoImage(image)

    def crear_widgets(self):
        # Crear un marco para los widgets
        frame = tk.Frame(self.root, bg='#7ea7b9')
        frame.pack(pady=20)

        # Etiqueta y campo de entrada para el archivo
        tk.Label(frame, text="Archivo Excel:", font="arial 12 bold", bg='#7ea7b9').pack(side=tk.LEFT, padx=10)
        self.archivo_entry = tk.Entry(frame, width=50)
        self.archivo_entry.pack(side=tk.LEFT, padx=10)

        # Botón para seleccionar archivo
        tk.Button(frame, text="Seleccionar Archivo", command=self.app.seleccionar_archivo).pack(side=tk.LEFT, padx=10)

        # Botón para procesar el archivo
        tk.Button(self.root, text="Procesar", font="Arial 16 bold", command=self.app.procesar_archivo).pack(side=tk.BOTTOM, pady=20)

    def seleccionar_archivo(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        self.archivo_entry.delete(0, tk.END)
        self.archivo_entry.insert(0, filename)