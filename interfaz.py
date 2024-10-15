import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from PIL import Image, ImageTk
import os
from osgeo import ogr
import pandas as pd

class InterfazGrafica:
    def __init__(self, root, app):
        self.root = root
        self.app = app
        self.root.title("Carga Masiva")
        self.root.geometry("800x400")
        self.root.state('zoomed')
        self.gdb_path = tk.StringVar()  # Inicialización de gdb_path
        self.excel_file_path = tk.StringVar()  # Inicialización de excel_file_path
        # Crear el contenedor de pestañas
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=1, fill="both")

        # Crear las pestañas
        self.tab_validaciones = tk.Frame(self.notebook, bg='#7ea7b9')
        self.tab_convertir_gdb = tk.Frame(self.notebook, bg='#7ea7b9')
        self.tab_Agregar_Fichas = tk.Frame(self.notebook, bg='#7ea7b9')

        self.notebook.add(self.tab_validaciones, text="Validaciones")
        self.notebook.add(self.tab_Agregar_Fichas, text="Agregar Fichas")
        self.notebook.add(self.tab_convertir_gdb, text="Convertir GDB")

        self.background_image = self.crear_imagen_semitransparente("./assets/Logo_Conestudios.png", 0.1)

        # Configurar la pestaña de validaciones
        self.configurar_pestania_validaciones()

        # Configurar la pestaña de convertir GDB
        self.configurar_pestania_convertir_gdb()
        
        self.configurar_pestania_agregar_fichas()
        
    def configurar_pestania_validaciones(self):
        # Cargar la imagen y configurarla como fondo en la pestaña de validaciones
        self.background_image = self.crear_imagen_semitransparente("./assets/Logo_Conestudios.png", 0.1)
        self.background_label = tk.Label(self.tab_validaciones, image=self.background_image)
        self.background_label.place(x=0, y=0, relwidth=1, relheight=1)

        # Crear los widgets dentro de la pestaña "Validaciones"
        frame_nph = tk.Frame(self.tab_validaciones, bg='#7ea7b9')
        frame_nph.pack(pady=20)

        tk.Label(frame_nph, text="Carga Masiva NPH:", font="arial 12 bold", bg='#7ea7b9').pack(side=tk.LEFT, padx=10)
        self.archivo_entry_nph = tk.Entry(frame_nph, width=50)
        self.archivo_entry_nph.pack(side=tk.LEFT, padx=10)
        self.boton_nph = tk.Button(frame_nph, text="Seleccionar Archivo NPH", command=self.seleccionar_archivo_nph)
        self.boton_nph.pack(side=tk.LEFT, padx=10)

        frame_botones = tk.Frame(self.tab_validaciones, bg='#7ea7b9')
        frame_botones.pack(side=tk.BOTTOM, pady=10)

        self.boton_procesar = tk.Button(frame_botones, text="Procesar", font="Arial 16 bold", command=None, state=tk.DISABLED)
        self.boton_procesar.pack(side=tk.LEFT, padx=(0, 20))

        self.boton_limpiar = tk.Button(frame_botones, text="Limpiar", font="Arial 16 bold", command=self.limpiar_seleccion)
        self.boton_limpiar.pack(side=tk.LEFT)

    def configurar_pestania_convertir_gdb(self):
        # Usar la misma imagen de fondo en la pestaña de convertir GDB
        self.background_label_convertir_gdb = tk.Label(self.tab_convertir_gdb, image=self.background_image)
        self.background_label_convertir_gdb.place(x=0, y=0, relwidth=1, relheight=1)

        # Crear widgets para seleccionar la geodatabase
        tk.Label(self.tab_convertir_gdb, text="Convertir GDB a GPKG", font="arial 12 bold", bg='#7ea7b9').pack(pady=20)

        # Botón para iniciar la conversión
        self.boton_convertir = tk.Button(self.tab_convertir_gdb, text="Seleccionar carpeta .gdb", font="Arial 16 bold", command=self.select_gdb_folder)
        self.boton_convertir.pack(pady=10)
    
    def configurar_pestania_agregar_fichas(self):
        """Configura la pestaña para agregar fichas."""
        # Cargar la imagen y configurarla como fondo en la pestaña de agregar fichas
        self.background_label_agregar_fichas = tk.Label(self.tab_Agregar_Fichas, image=self.background_image)
        self.background_label_agregar_fichas.place(x=0, y=0, relwidth=1, relheight=1)

        # Crear los widgets dentro de la pestaña "Agregar Fichas" usando grid
        tk.Label(self.tab_Agregar_Fichas, text="Seleccionar GDB:", bg='#7ea7b9').grid(row=0, column=0, pady=10)
        tk.Button(self.tab_Agregar_Fichas, text="Seleccionar GDB", command=self.select_gdb).grid(row=0, column=1, pady=10)

        tk.Label(self.tab_Agregar_Fichas, text="GDB seleccionada:", bg='#7ea7b9').grid(row=1, column=0, pady=10)
        tk.Label(self.tab_Agregar_Fichas, textvariable=self.gdb_path, bg='#7ea7b9').grid(row=1, column=1, pady=10)

        tk.Label(self.tab_Agregar_Fichas, text="Seleccionar Excel:", bg='#7ea7b9').grid(row=2, column=0, pady=10)
        tk.Button(self.tab_Agregar_Fichas, text="Seleccionar Excel", command=self.select_excel).grid(row=2, column=1, pady=10)

        tk.Label(self.tab_Agregar_Fichas, text="Archivo Excel seleccionado:", bg='#7ea7b9').grid(row=3, column=0, pady=10)
        tk.Label(self.tab_Agregar_Fichas, textvariable=self.excel_file_path, bg='#7ea7b9').grid(row=3, column=1, pady=10)

        tk.Button(self.tab_Agregar_Fichas, text="Ejecutar", command=self.process_all).grid(row=4, column=0, pady=10)

    def crear_imagen_semitransparente(self, image_path, alpha):
        image = Image.open(image_path)
        image = image.convert("RGBA")
        data = image.getdata()

        new_data = [(item[:-1] + (int(255 * alpha),)) for item in data]
        image.putdata(new_data)
        image = image.resize((self.root.winfo_screenwidth(), self.root.winfo_screenheight()), Image.LANCZOS)
        return ImageTk.PhotoImage(image)

    def seleccionar_archivo_nph(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.archivo_entry_nph.delete(0, tk.END)
            self.archivo_entry_nph.insert(0, filename)
            self.boton_procesar.config(command=self.app.procesar_archivo, state=tk.NORMAL)
    
    
    def select_excel(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.excel_file_path.set(filename)
    
    def limpiar_seleccion(self):
        self.archivo_entry_nph.delete(0, tk.END)
        self.boton_procesar.config(state=tk.DISABLED)
    
    def limpiar_seleccion_fichas(self):
        self.archivo_entry_fichas.delete(0, tk.END)
        self.boton_procesar_fichas.config(state=tk.DISABLED)
    
    def select_gdb(self):
        gdb_folder = filedialog.askdirectory(title="Seleccionar carpeta .gdb")
        if gdb_folder:
            self.gdb_path.set(gdb_folder)
            
    def select_gdb_folder(self):
        gdb_folder = filedialog.askdirectory(title="Seleccionar carpeta .gdb")
        if gdb_folder:
            output_path = os.path.join(os.path.dirname(gdb_folder), os.path.basename(gdb_folder) + ".gpkg")  # Define la ruta de salida .gpkg
            self.convert_gdb_to_gpkg(gdb_folder, output_path)

    def convert_gdb_to_gpkg(self, gdb_folder, output_path):
        # Elimina el archivo de salida si ya existe
        if os.path.exists(output_path):
            os.remove(output_path)

        driver = ogr.GetDriverByName('OpenFileGDB')
        if driver is None:
            messagebox.showerror("Error", "Driver OpenFileGDB no está disponible.")
            return

        # Abre la carpeta GDB como un geodatabase
        gdb = driver.Open(gdb_folder, 0)

        if not gdb:
            messagebox.showerror("Error", f"No se pudo abrir la geodatabase GDB: {gdb_folder}")
            return

        output_driver = ogr.GetDriverByName('GPKG')
        if output_driver is None:
            messagebox.showerror("Error", "No se pudo encontrar el driver de salida.")
            return

        # Crea el DataSource de salida
        output_layer = output_driver.CreateDataSource(output_path)
        if output_layer is None:
            messagebox.showerror("Error", "No se pudo crear el archivo de salida.")
            return

        for i in range(gdb.GetLayerCount()):
            layer = gdb.GetLayerByIndex(i)
            output_layer.CopyLayer(layer, layer.GetName())

        messagebox.showinfo("Éxito", "Conversión completada.")

    def crear_imagen_semitransparente(self, image_path, alpha):
        image = Image.open(image_path)
        image = image.convert("RGBA")
        data = image.getdata()

        new_data = [(item[:-1] + (int(255 * alpha),)) for item in data]
        image.putdata(new_data)
        image = image.resize((self.root.winfo_screenwidth(), self.root.winfo_screenheight()), Image.LANCZOS)
        return ImageTk.PhotoImage(image)
    
    def procesar_fichas(self):
        messagebox.INFO("Aun no se han agregado funcionalidades")

    def process_all(self):
        """Ejecuta el proceso completo utilizando la GDB y el archivo Excel seleccionados."""
        gdb = self.gdb_path.get()
        excel_file = self.excel_file_path.get()

        if not gdb or not excel_file:
            messagebox.showerror("Error", "Por favor, selecciona una GDB y un archivo Excel.")
            return

        try:
            # Aquí agrega la lógica para procesar la GDB y el archivo Excel
            print(f"Procesando GDB: {gdb}")
            print(f"Procesando archivo Excel: {excel_file}")

            # Ejemplo de procesamiento con pandas
            df = pd.read_excel(excel_file)
            print(df.head())  # Solo un ejemplo, imprime las primeras filas del DataFrame

            # Lógica adicional para agregar fichas a la GDB...
            messagebox.showinfo("Éxito", "Proceso completado exitosamente.")
        except Exception as e:
            messagebox.showerror("Error al procesar", str(e))
            
    