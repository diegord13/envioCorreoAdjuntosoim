#Importing the modules
import numpy as np
import openpyxl
from openpyxl_image_loader import SheetImageLoader
import pandas as pd
import win32com.client as client
import os
from datetime import date
import fnmatch

today = str(date.today())


def imagen(celda):
    pxl_doc = openpyxl.load_workbook(f'{file}')
    sheet = pxl_doc['imagenes']
    image_loader = SheetImageLoader(sheet)
    image = image_loader.get(celda)
    
    image.save(f'{file[0:17]}_Reintegros 1.png')
    
    

def proveedor(valor_buscado):
    df_im = pd.read_excel(file,sheet_name='imagenes')
    df_im=df_im.fillna('0')
    filas = df_im.loc[df_im['proveedor'].str.contains(valor_buscado)].index
    if filas is not None:
        
        info =df_im.loc[df_im['proveedor'].str.contains(valor_buscado)]
        imagen('A'+str(filas[0]+3))
        return info



def send_mail(recipient_email,
                         subject,
                         content):
    
    outlook = client.Dispatch("Outlook.Application")

    email = outlook.CreateItem(0)
    email.To = recipient_email
    email.CC = 'elizabeth.uribe@wom.co'
    email.Subject = subject
    email.Body = content
    email.Attachments.Add(os.path.join(os.getcwd(),f'{file[0:17]}_Reintegros 1.png'))
    email.Attachments.Add(os.path.join(os.getcwd(),f'{file[0:17]}_{proveedores}.xlsx'))
    email.Send()
    #outlook.Quit()

def get_excel_file():
    
    file_list = os.listdir()
    
    xlsx_file_list = [file for file in file_list if '.xlsx' in file]
    if len(xlsx_file_list) < 1: 
        raise ValueError('There is no excel file in the current directory!')
    else:
        
        
        #print(xlsx_filename)
    #df_reintegro = pd.read_excel(f'./{xlsx_filename}',sheet_name='Reintegros')
    
        return xlsx_file_list

def listado(valor):
    
    df2=pd.read_excel(f'{file}')
   
    df2 = df2.dropna()
    filtro = df2.loc[df2['Proveedor'].str.contains(valor)]   
    filtro.to_excel(f'{file[0:17]}_{proveedores}.xlsx')

def correo_proveedor(prov):
    df_correo=pd.read_excel("correo/correo_proveedores.xlsx")
    filtro_correo = df_correo.loc[df_correo['PROVEEDOR'].str.contains(prov)].index
    correo=df_correo['CORREO'][filtro_correo]

    s=correo.values
    if len(s) > 0:
        if s!=0:
            for val in s:
                
                return val
        else:
            val='sonia.rodriguezact@wom.co'
            return val
        
    else:
        val="sonia.rodriguezact@wom.co"
        
    return val
    
files = get_excel_file()
for file in files:
    
    df=pd.read_excel(f'{file}')
    
    df = df.dropna()
    df= df.drop_duplicates('Proveedor')
    serProveedor = df['Proveedor']
    asunto="Soporte Pagos"

    for proveedores in serProveedor:
        x=proveedor(proveedores) 
        listado(proveedores)
        
        email=correo_proveedor(proveedores)
        body_correo=f"""
        
        Buenos dias

        Este es el reporte de pago {file[0:17]}

        {str(x)}

        Cordialmente
    """
        
       # send_mail(email,f'Reporte de pago {file[0:17]}',body_correo)
# Filtra los archivos de la carpeta por extensión .xlsx y .xls
folder_path = os.path.dirname(os.path.abspath(__file__))

print(folder_path)
excel_files = fnmatch.filter(os.listdir(folder_path), '*.xlsx') + fnmatch.filter(os.listdir(folder_path), '*.xls')

# Elimina cada archivo filtrado
for file in excel_files:
    os.remove(os.path.join(folder_path, file))




