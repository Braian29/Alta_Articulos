import customtkinter as ctk
import tkinter as tk
from tkinter import ttk
from tkinter import PhotoImage
import requests
import json
import subprocess
import urllib3 
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning) 
import pandas as pd 
import sys 
import os 

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def login_sap():
    url = "https://XXXXXXXXXXXXXX/b1s/v1/Login"
    headers = {'Content-Type': 'application/json'}
    company = company_combobox.get()
    username = username_entry.get()
    password = password_entry.get()
    data = {"CompanyDB": company, "Password": password, "UserName": username}
    # Mostrar mensaje "Iniciando sesión"
    status_label.configure(text="Iniciando sesión...")
    # Simular tiempo de espera.
    root.after(1000, lambda: process_login(url, headers, data))


def process_login(url, headers, data):
    response = requests.post(url, headers=headers, data=json.dumps(data), verify=False)
    if response.status_code == 200:
        # Inicio de sesión exitoso
        status_label.configure(text="Inicio de sesión exitoso.")
        # Obtener los cookies de la respuesta
        cookies = response.cookies.get_dict()
	# Guardar los cookies en un archivo JSON
        with open("cookies.json", "w") as f:
                json.dump(cookies, f)
        # Ocultar la ventana principal temporalmente
        root.withdraw()
        try:
            # Ejecutar el script 'interfaz.py' y pasar los cookies como argumento
            subprocess.run([sys.executable, resource_path("custom_interfaz.py")])
        except Exception as e:
            print("Error al ejecutar interfaz.py:", e)
        # Cerrar la ventana actual
        root.destroy()
    else:
        # Inicio de sesión fallido
        status_label.configure(text="Inicio de sesión fallido. Verifique los datos ingresados.")
        print(response.text)
ctk.set_appearance_mode("light")  # Modes: system (default), light, dark


# Crear la ventana principal
root = ctk.CTk()
root.title("Alta de Artículos - SAP Business One")
root.geometry("370x600")


# Logo SUPERCLIN
width = 362
height = 342
image = PhotoImage(file=resource_path("logo.png"))
image = image.subsample(image.width() // width, image.height() // height)
label = tk.Label(root, image=image)
label.pack(pady=2)


# Crear un marco para agrupar los widgets
frame = ctk.CTkFrame(root)
frame.pack(pady=10)

# Etiquetas y campos de entrada
ctk.CTkLabel(frame, text="Compañía:").pack(pady=5)
company_combobox = ttk.Combobox(frame, values=['Base1', 'Base2'])
company_combobox.pack(pady=2)
company_combobox.set('base1')  # Valor predeterminado

ctk.CTkLabel(frame, text="Usuario:").pack(pady=5)
username_entry = ctk.CTkEntry(frame)
username_entry.pack(pady=3)

ctk.CTkLabel(frame, text="Password:").pack(pady=5)
password_entry = ctk.CTkEntry(frame, show="*")
password_entry.pack(pady=4)

# Botón de inicio de sesión
login_button = ctk.CTkButton(frame, text="Iniciar Sesión", command=login_sap)
login_button.pack(pady=5)

# Etiqueta para mostrar el estado del inicio de sesión
status_label = ctk.CTkLabel(root, text="", fg_color="white")
status_label.pack(pady=6)

root.mainloop()
