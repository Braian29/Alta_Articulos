import tkinter as tk
from tkinter import filedialog
import pandas as pd
from tkinter import ttk
import requests
import json
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
import pandas as pd
from pandas import json_normalize
import datetime
import numpy as np
from numpy import NaN
from tkinter import messagebox
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import sys
import os 

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Datos obtenidos en InicioSesion.py
with open(resource_path("cookies.json"), "r") as f:
    cookies = json.load(f)
headers = {'Content-Type': 'application/json'}

# toma un valor y lo convierte en una cadena de texto, formateándolo como un número entero si no tiene parte decimal, o manteniendo su formato original si tiene parte decimal o si no se puede convertir a un número.
def format_number(x):
    try:
        x = float(x)
        if x % 1 == 0:
            return '{:.0f}'.format(x)
        else:
            return str(x)
    except ValueError:
        return x

#Procesar Excel
def procesar_archivo_excel():
    ruta = 'C:/Alta_Articulos/Alta articulos planilla 2023.xlsx'
    df = pd.read_excel(ruta, sheet_name='Service-Layer-Ok') #Leer Excel y crear DataFrame
    df = df.loc[df['ItemsGroupCode'] > 0] #Seleccionar los artículos donde 'ItemsGroupCode' sea mayor a 0 para eliminar filas innecesarias.  
    df = df.dropna(subset=['EAN'])#elimina las filas del DataFrame 'df' donde la columna 'EAN' tiene valores ausentes (NaN)
    # Seleccionar el registro con el ItemCode más grande
    max_item_code = df['ItemCode'].max()
    df = df[df['ItemCode'] == max_item_code]    
    #Verificar si el valor no es , Si el valor no es nulo, redondea y Convierte el valor redondeado a un entero.
    cols = ['ItemCode', 'ItemsGroupCode', 'EAN', 'Mainsupplier', 'SalesItemsPerUnit', 'U_B1SYS_MTXCode',  'UoMEntry',
            'DefaultPurchasingUoMEntry', 'SalesQtyPerPackUnit', 'SalesUnitLength', 'SalesUnitWidth', 'SalesUnitHeight',
            'SalesVolumeUnit', 'SalesUnitWeight', 'PurchaseUnitHeight1', 'PurchaseUnitLength1', 'PurchaseUnitWeight1', 'PurchaseUnitWidth1',
            'SalesUnitHeight1', 'SalesUnitLength1', 'SalesUnitWeight1', 'SalesUnitWidth1', 'UoMGroupEntry', 'InventoryUoMEntry', 'DefaultSalesUoMEntry','PricingUnit','Series']
    for col in cols:
        df[col] = df[col].apply(lambda x: int(round(x)) if pd.notna(x) else x) 
    #intenta convertir sus valores a números y, si es posible, redondea esos valores a enteros. Cualquier valor que no se pueda convertir a número se reemplaza con NaN
    cols = ['EAN', 'U_B1SYS_MTXCode']
    for col in cols:
        df[col] = pd.to_numeric(df[col], errors='coerce')
        df[col] = df[col].apply(lambda x: int(round(x)) if pd.notna(x) else x)
    #Agrupar datos en diccionario dentro del campo 'ItemBarCodeCollection'
    df['ItemBarCodeCollection'] = df.apply(lambda x: [{'UoMEntry': x['UoMEntry'], 'Barcode': x['BarCode'], 'FreeText': x['FreeText']}], axis=1)
    #Borrar datos agrupados    
    df = df.drop(['UoMEntry', 'BarCode', 'FreeText'], axis=1)
    #Renombrar columna
    df.rename(columns={'EAN': 'BarCode'}, inplace=True)
    #Reemplazar ',' por '.' 
    df['SalesUnitVolume'] = df['SalesUnitVolume'].astype(str).str.replace(',', '.').astype(float).round().astype(int)
    df['AvgStdPrice'] = df['AvgStdPrice'].astype(str).str.replace(',', '.')
    df['MovingAveragePrice'] = df['MovingAveragePrice'].astype(str).str.replace(',', '.')
    df['SupplierCatalogNo'] = df['SupplierCatalogNo'].apply(format_number)
    cols2 = ['MovingAveragePrice', 'AvgStdPrice']
    df[cols2] = df[cols2].round(2)
    return df

def buscar_duplicados(df):
    barcodes_url = "https://177.85.33.53:50695/b1s/v1/BarCodes"
    duplicados = pd.DataFrame(columns=df.columns)
    no_encontrados = []
    for index, row in df.iterrows():
        barcode = row.BarCode
        query = "?$filter=Barcode eq '{}'".format(barcode)
        barcodes_response = requests.get(barcodes_url + query, headers=headers, cookies=cookies, verify=False)

        if barcodes_response.status_code == 200:
            data = barcodes_response.json()
            if data['value']:
                duplicados = pd.concat([duplicados, row.to_frame().T])
                df.drop(index, inplace=True)
            else:
                no_encontrados.append(barcode)
        else:
            print("Error al obtener el código de barras {} - Código de estado: {}".format(barcode, barcodes_response.status_code))
    if no_encontrados:
        print("Los siguientes códigos de barras no están en la base de datos: {}".format(no_encontrados))
    if not duplicados.empty:
        print("Los siguientes códigos de barras están duplicados y han sido movidos al dataframe 'duplicados':")
        print(duplicados['BarCode'])
    return df, duplicados

# Función para crear artículos
def crear_articulos(df):
    global ArticulosCreados
    #Crear dataframe CodigoNCM
    CodigoNCM = df[['BarCode', 'ItemName']].copy()
    CodigoNCM['GroupCode'] = 'C'
    CodigoNCM['AbsEntry'] = ''
    CodigoNCM.rename(columns={'BarCode': 'NCMCode', 'ItemName': 'Description'}, inplace=True)
    codigo_ncm_url = "https://177.85.33.53:50695/b1s/v1/NCMCodesSetup"
    for index, row in CodigoNCM.iterrows():
        ncm_code = row['NCMCode']
        query = "?$filter=NCMCode eq '{}'".format(ncm_code)
        ncm_response = requests.get(codigo_ncm_url + query, headers=headers, cookies=cookies, verify=False)
        if ncm_response.status_code == 200:
            data = ncm_response.json()
            if data['value']:
                # Existe un registro con el mismo NCMCode en la entidad, obtener el valor de AbsEntry
                abs_entry = data['value'][0]['AbsEntry']
                # Asignar el valor de AbsEntry al DataFrame
                CodigoNCM.at[index, 'AbsEntry'] = abs_entry
            else:
                # No existe un registro con el mismo NCMCode, crear uno nuevo
                payload = {
                    "NCMCode": ncm_code,
                    "Description": row['Description'],
                    "GroupCode": row['GroupCode']
                }
                ncm_create_response = requests.post(codigo_ncm_url, headers=headers, data=json.dumps(payload), cookies=cookies, verify=False)
                if ncm_create_response.status_code == 201:
                    # Registro creado exitosamente, obtener el valor de AbsEntry
                    created_data = ncm_create_response.json()
                    abs_entry = created_data['AbsEntry']
                    # Asignar el valor de AbsEntry al DataFrame
                    CodigoNCM.at[index, 'AbsEntry'] = abs_entry
                else:
                    print("Error al crear el registro para NCMCode {}: {}".format(ncm_code, ncm_create_response.text))
        else:
            print("Error al obtener el registro para NCMCode {}: {}".format(ncm_code, ncm_response.text))
    # Crear un diccionario de mapeo usando 'CodigoNCM'
    codigo_ncm_mapping = dict(zip(CodigoNCM['NCMCode'], CodigoNCM['AbsEntry']))
    # Usar la función map para reemplazar los valores en 'U_B1SYS_MTXCode' con los valores de 'AbsEntry'
    df['U_B1SYS_MTXCode'] = df['BarCode'].map(codigo_ncm_mapping)
    #Conversión de Dataframe a Diccionario.
    json_records = df.to_json(orient='records')
    dict_data = json.loads(json_records)
    # Crear artículos nuevos con el objeto Items
    items_url = "https://177.85.33.53:50695/b1s/v1/Items"
    articulos_creados_data = []
    for item in dict_data:
        items_response = requests.post(items_url, headers=headers, data=json.dumps(item), cookies=cookies, verify=False)
        if items_response.status_code == 201:
            # Artículo creado exitosamente
            response_data = items_response.json()
            # Extraer los campos que necesitas (ItemCode, ItemName y BarCode)
            item_code = response_data['ItemCode']
            item_name = response_data['ItemName']
            bar_code = item['BarCode']
            # Agregar los datos a la lista 'articulos_creados_data'
            articulos_creados_data.append({'ItemCode': item_code, 'ItemName': item_name, 'BarCode': bar_code})
        else:
            # Error al crear el artículo
            print(items_response.text)
    # Crear el DataFrame 'ArticulosCreados' a partir de la lista 'articulos_creados_data'
    ArticulosCreados = pd.DataFrame(articulos_creados_data)
    messagebox.showinfo("Éxito", "Artículos creados exitosamente. Presione el Botón 'Informar' para pedir que Administración asigne los costos alos nuevos artículos.")  # O mensaje de error si corresponde.
    return ArticulosCreados
    
def informar():
    global ArticulosCreados
    # Configuración del correo
    smtp_server = 'smtp.super-clin.com.ar'  
    smtp_port = 587  
    sender_email = 'braian.alonso@super-clin.com.ar'  # Remitente
    sender_password = 'alon3786'  
    receiver_emails = ['compras@super-clin.com.ar','braian.alonso@super-clin.com.ar', 'roberto.carballo@superclin.com' , 'administracion@super-clin.com.ar', 'compras2@super-clin.com.ar', 'ignacio.galindez@super-clin.com.ar']  # Lista de destinatarios
    # Crear el mensaje
    subject = 'Nuevos artículos creados'
    body = """
    <html>
    <body>
    <p><strong>POR FAVOR NO RESPONDER ESTE CORREO.</strong></p>
    <p>Se han creado los siguientes artículos:</p>
    <p></p>
    {0}
    </body>
    </html>
    """.format(ArticulosCreados.to_html(index=False))
    # Conexión al servidor SMTP y envío del correo
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        
        for receiver_email in receiver_emails:
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = receiver_email
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'html'))
            
            text = msg.as_string()
            server.sendmail(sender_email, receiver_email, text)
            
        server.quit()
        messagebox.showinfo("Éxito", "Correo enviado correctamente.")
    except Exception as e:
        print("Error al enviar el correo:", str(e))
        messagebox.showerror("Error", "No se pudo enviar el correo.")



if __name__ == "__main__":
    df = procesar_archivo_excel()
    print("DataFrame principal:")
    df.to_excel("DataFrame principal.xlsx", index=False)
    print(df)
    df, duplicados = buscar_duplicados(df)
    print("DataFrame Tratado:")
    print(df)
    print("DataFrame de duplicados:")
    print(duplicados)
    duplicados.to_excel("duplicados.xlsx", index=False)
    df.to_excel("df.xlsx", index=False)