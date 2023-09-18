import tkinter as tk
import customtkinter as ctk
from tkinter import ttk
from tkinter import PhotoImage
from tkinter import filedialog
from funciones import procesar_archivo_excel, buscar_duplicados, informar, crear_articulos
import pandas as pd
from tkinter import messagebox
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
import os 
import sys

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# DataFrames de Ejemplo
data = {"ItemCode": [],"ItemName": [],"BarCode": []}
df_t = pd.DataFrame(data)
creados_t = df_t
data_nuevos = {"ItemCode": [], "ItemName": [], "BarCode": []}
duplicados_t = pd.DataFrame(data_nuevos)
treeview_creados = None  # Declara la variable global para acceder desde otras funciones
ArticulosCreados = pd.DataFrame(columns=['ItemCode', 'ItemName', 'BarCode'])

def ejecutar_funciones():
    try:
        global df, duplicados
        boton_seleccionar.configure(state="disabled", text="Datos Cargados")
        df = procesar_archivo_excel()
        df, duplicados = buscar_duplicados(df)
        duplicados_t = duplicados[['ItemCode', 'ItemName', 'BarCode']]
        df_t = df[['ItemCode', 'ItemName', 'BarCode']]
        show_tables(df_t, duplicados_t, creados_t)
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al ejecutar las funciones:\n\n{str(e)}")


# FUNCION  PARA CREACIÓN DE ARTICULOS  
def update_creados():
    global creados_t, ArticulosCreados
    creados_t = ArticulosCreados[['ItemCode', 'ItemName', 'BarCode']]
    actualizar_tabla_creados()

def ejecutar_crear():
    try:
        global df, ArticulosCreados
        boton_crear_articulos.configure(state="disabled", text="Artículos Creados")
        # Llama a la función crear_articulos y verifica su salida
        ArticulosCreados = crear_articulos(df)
        print(df)  # Agrega esta línea para verificar el valor de df
        update_creados()  # Actualiza creados_t y treeview_creados
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al crear los artículos:\n\n{str(e)}")

def actualizar_tabla_creados():
    global creados_t, treeview_creados
    for index, row in creados_t.iterrows():
        treeview_creados.insert("", "end", values=(row['ItemCode'], row['ItemName'], row['BarCode']))
    show_tables(df_t, duplicados_t, creados_t)

# VENTANA
ctk.set_appearance_mode("light")  # Modes: system (default), light, dark
root = ctk.CTk()
root.after(0, lambda: root.wm_state('zoomed'))
root.title("Alta de Artículos - SAP Business One")

# Logo SUPERCLIN
width = round(280)
height = round(190)
image = PhotoImage(file=resource_path("logo2.png"))
image = image.subsample(image.width() // width, image.height() // height)
label = tk.Label(root, image=image)
label.grid(row=1, column=0, columnspan=2)

# Widget texto TITULO
font = ("Helvetica", 28, "bold")  # Puedes ajustar el nombre de la fuente y el tamaño
label1 = ctk.CTkLabel(root, font=font, text="Alta de Artículos en SAP Business One.")
label1.grid(row=1, column=2, columnspan=3)

# Widget de Boton 'Cargar_Datos'
boton_seleccionar = ctk.CTkButton(root, text="Cargar Datos", command=ejecutar_funciones)
boton_seleccionar.grid(row=4, column=1)

# Widget 'Texto df'
font = ("Helvetica", 15, "bold")  # Puedes ajustar el nombre de la fuente y el tamaño
label1 = ctk.CTkLabel(root, font=font, text="Artículos para crear")
label1.grid(row=3, column=2)

# Widget 'Texto df'
font = ("Helvetica", 15, "bold")  # Puedes ajustar el nombre de la fuente y el tamaño
label1 = ctk.CTkLabel(root, font=font, text="Artículos Duplicados")
label1.grid(row=3, column=3)

# Widgets de Tablas
# Crear una función para mostrar los DataFrames en la ventana de Tkinter
def show_tables(df_t, duplicados_t, creados_t):
    global treeview_creados  # Acceder a la variable global treeview_creados 
    # Crear tablas usando Treeview de ttk
    treeview_df = ttk.Treeview(root, columns=list(df_t.columns), show="headings")
    treeview_duplicados = ttk.Treeview(root, columns=list(duplicados_t.columns), show="headings")
    treeview_creados = ttk.Treeview(root, columns=list(creados_t.columns), show="headings")

    # Configurar encabezados de las columnas
    for col in df_t.columns:
        treeview_df.heading(col, text=col)
    for col in duplicados_t.columns:
        treeview_duplicados.heading(col, text=col)
    for col in creados_t.columns:
        treeview_creados.heading(col, text=col)     

    # Empacar las tablas en la ventana
    treeview_df.grid(row=4, column=2, padx=20, pady=20)
    treeview_duplicados.grid(row=4, column=3, padx=20, pady=20)
    treeview_creados.grid(row=7, column=2, pady=20)
    
    # Ajustar los anchos de las columnas individualmente
    treeview_df.column('ItemCode', anchor='center', width=70)
    treeview_df.column('ItemName', anchor='center', width=380)
    treeview_df.column('BarCode', anchor='center', width=100)

    treeview_duplicados.column('ItemCode', anchor='center', width=70)
    treeview_duplicados.column('ItemName', anchor='center', width=380)
    treeview_duplicados.column('BarCode', anchor='center', width=100)

    treeview_creados.column('ItemCode', anchor='center', width=70)
    treeview_creados.column('ItemName', anchor='center', width=380)
    treeview_creados.column('BarCode', anchor='center', width=100)

    # Limpiar las tablas
    treeview_df.delete(*treeview_df.get_children())
    treeview_duplicados.delete(*treeview_duplicados.get_children())
    treeview_creados.delete(*treeview_creados.get_children())

    # Actualizar tablas
    for index, row in df_t.iterrows():
        treeview_df.insert("", "end", values=(row['ItemCode'], row['ItemName'], row['BarCode']))

    for index, row in duplicados_t.iterrows():
        treeview_duplicados.insert("", "end", values=(row['ItemCode'], row['ItemName'], row['BarCode']))

    for index, row in creados_t.iterrows():
        treeview_creados.insert("", "end", values=(row['ItemCode'], row['ItemName'], row['BarCode']))



# Llamar a la función para mostrar las tablas
show_tables(df_t, duplicados_t, creados_t)



# Widget de Boton     Crear_Artículos
boton_crear_articulos = ctk.CTkButton(root, text="Crear Artículos", command=ejecutar_crear)
boton_crear_articulos.grid(row=7, column=1)


# Widget texto     ArticulosCreados
font = ("Helvetica", 15, "bold") 
label1 = ctk.CTkLabel(root, font=font, text="Artículos Creados")
label1.grid(row=6, column=2)


# Widget de Boton     informar
boton_informar = ctk.CTkButton(root, text="Informar", command=informar)
boton_informar.grid(row=9, column=5, pady=30)


# Crear una línea divisoria
line_divider = ctk.CTkFrame(root, bg_color="white", width=2, height=5)
line_divider.grid(row=0, column=0,  padx=100, pady=5)


# Crear un lienzo (Canvas) que ocupe todo el ancho de la ventana principal
canvas = ctk.CTkCanvas(root, width=root.winfo_screenwidth(), height=1)
canvas.grid(row=8, column=0, columnspan= 7, pady=5, sticky="ew")  # Utilizamos sticky="ew" para que ocupe todo el ancho
# Dibujar una línea horizontal
canvas.create_line(10, 1, canvas.winfo_width(), 1, fill="black", width=2)

root.mainloop()