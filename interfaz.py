import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter import ttk #barra de carga
from tkinter import font 
from TelefonicaFija import procesar_Telefonica_Fija
from TelefonicaDatos import procesar_Telefonica_Datos
from TelefonicaMovil import procesar_Telefonica_Movil
from TelecomFija import procesar_Telecom_Fija
from TelecomMovil import procesar_Telecom_Movil
from TelecomDatos import procesar_Telecom_Datos
from TelecomOtro import procesar_Telecom_Otro

import os
import PyPDF2
from PyPDF2 import PdfReader
# ruta_icono = os.path.join("resources", "TorreonColor.ico")

def elegir_destino_excel():
    ruta_excel = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Archivos Excel", ".xlsx"), ("Todos los archivos", ".*")]
    )
    return ruta_excel

def cargar_Telefonica_Fija(barra_progreso):
  ruta_carpeta =filedialog.askdirectory()
  if not ruta_carpeta:
        print("No se seleccionó ninguna carpeta.")
        return
  ruta_excel = elegir_destino_excel()
  if not ruta_excel:
        print("No se seleccionó ningún archivo Excel para guardar.")
        return
  procesar_Telefonica_Fija(ruta_carpeta, barra_progreso,ruta_excel)
  
def cargar_Telefonica_Movil(barra_progreso):
  ruta_carpeta =filedialog.askdirectory()
  if not ruta_carpeta:
        print("No se seleccionó ninguna carpeta.")
        return
  ruta_excel = elegir_destino_excel()
  if not ruta_excel:
        print("No se seleccionó ningún archivo Excel para guardar.")
        return
  procesar_Telefonica_Movil(ruta_carpeta, barra_progreso,ruta_excel)  
  
def cargar_Telefonica_Datos(barra_progreso):
  ruta_carpeta =filedialog.askdirectory()
  if not ruta_carpeta:
        print("No se seleccionó ninguna carpeta.")
        return
  ruta_excel = elegir_destino_excel()
  if not ruta_excel:
        print("No se seleccionó ningún archivo Excel para guardar.")
        return
  procesar_Telefonica_Datos(ruta_carpeta, barra_progreso,ruta_excel)  
  
#-----------------------------------------------------------------------------------  
  
def cargar_Telecom_Fija(barra_progreso):
  ruta_carpeta =filedialog.askdirectory()
  if not ruta_carpeta:
        print("No se seleccionó ninguna carpeta.")
        return
  ruta_excel = elegir_destino_excel()
  if not ruta_excel:
        print("No se seleccionó ningún archivo Excel para guardar.")
        return
  procesar_Telecom_Fija(ruta_carpeta, barra_progreso,ruta_excel)
  
def cargar_Telecom_Movil(barra_progreso):
  ruta_carpeta =filedialog.askdirectory()
  if not ruta_carpeta:
        print("No se seleccionó ninguna carpeta.")
        return
  ruta_excel = elegir_destino_excel()
  if not ruta_excel:
        print("No se seleccionó ningún archivo Excel para guardar.")
        return
  procesar_Telecom_Movil(ruta_carpeta, barra_progreso,ruta_excel)  
  
def cargar_Telecom_Datos(barra_progreso):
  ruta_carpeta =filedialog.askdirectory()
  if not ruta_carpeta:
        print("No se seleccionó ninguna carpeta.")
        return
  ruta_excel = elegir_destino_excel()
  if not ruta_excel:
        print("No se seleccionó ningún archivo Excel para guardar.")
        return
  procesar_Telecom_Datos(ruta_carpeta, barra_progreso,ruta_excel)

def cargar_Telecom_Otro(barra_progreso):
  ruta_carpeta =filedialog.askdirectory()
  if not ruta_carpeta:
        print("No se seleccionó ninguna carpeta.")
        return
  ruta_excel = elegir_destino_excel()
  if not ruta_excel:
        print("No se seleccionó ningún archivo Excel para guardar.")
        return
  procesar_Telecom_Otro(ruta_carpeta, barra_progreso,ruta_excel)
  
def iniciar_interfaz():
 root = tk.Tk()
 root.title("Facturación PDF a Excel")
#  root.iconbitmap(ruta_icono)
 root.geometry("550x470")
 root.resizable(False,False)
 root.config(bg="#B7D2E9", bd=20, relief="groove")

 icono_imagen = PhotoImage(file="./resources/torreon.png")  # Reemplaza con la ruta a tu imagen
 icono_imagen = icono_imagen.subsample(6, 6)

 icono_label = tk.Label(root, image=icono_imagen, bg="#B7D2E9")
 icono_label.place(x=5, y=5)
 
 titulo1 = tk.Label(root,text="TELEFONICA | MOVISTAR", font=("Colibri", 20,"underline"), bg="#B7D2E9", fg="#000000")
 titulo1.pack(pady=10)

  # Crear un Frame que servirá como recuadro para los botones
 marco_botones = tk.Frame(root, relief="sunken", bd=2, bg="#B7D2E9")  # relief="raised" para el borde, bd=2 para el grosor
 marco_botones.place(x=50, y=63, width=400, height=70)  # Posición y tamaño del recuadro

 boton_origen = Button(root, text="Fijo", width=5, height=1,
                       font=("Colibri", 16), command=lambda: cargar_Telefonica_Fija(barra_progreso), relief="raised", bd=2, bg="#B7D2E9", fg="#000000")
 boton_origen.place(x=220, y=75)

 boton_origen2 = Button(root, text="Datos", width=5, height=1,
                       font=("Colibri", 16), command=lambda: cargar_Telefonica_Datos(barra_progreso), relief="raised", bd=2, bg="#B7D2E9", fg="#000000")
 boton_origen2.place(x=70, y=75)

 boton_origen3 = Button(root, text="Móvil", width=5, height=1,
                       font=("Colibri", 16), command=lambda: cargar_Telefonica_Movil(barra_progreso), relief="raised", bd=2, bg="#B7D2E9", fg="#000000")
 boton_origen3.place(x=360, y=75)

 barra_progreso = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
 barra_progreso.place(x=50,y=150)

 titulo2 = tk.Label(root,text="TELECOM", font=("Colibri", 20,"underline"), bg="#B7D2E9", fg="#000000")
 titulo2.place(x=180,y=175)
  # Crear un Frame que servirá como recuadro para los botones
 marco_botones = tk.Frame(root, relief="sunken", bd=2, bg="#B7D2E9")  # relief="raised" para el borde, bd=2 para el grosor
 marco_botones.place(x=50, y=215, width=400, height=70)  # Posición y tamaño del recuadro

 boton_origen4 = Button(root, text="Fijo", width=5, height=1,
                       font=("Colibri", 16), command=lambda: cargar_Telecom_Fija(barra_progreso), relief="raised", bd=2, bg="#B7D2E9", fg="#000000")
 boton_origen4.place(x=165, y=230)

 boton_origen5 = Button(root, text="Datos", width=5, height=1,
                       font=("Colibri", 16), command=lambda: cargar_Telecom_Datos(barra_progreso), relief="raised", bd=2, bg="#B7D2E9", fg="#000000")
 boton_origen5.place(x=60, y=230)

 boton_origen6 = Button(root, text="Móvil", width=5, height=1,
                       font=("Colibri", 16), command=lambda: cargar_Telecom_Movil(barra_progreso), relief="raised", bd=2, bg="#B7D2E9", fg="#000000")
 boton_origen6.place(x=265, y=230)

 boton_origen10 = Button(root, text="Otro", width=5, height=1,
                       font=("Colibri", 16), command=lambda: cargar_Telecom_Otro(barra_progreso), relief="raised", bd=2, bg="#B7D2E9", fg="#000000")
 boton_origen10.place(x=365, y=230)

 titulo3 = tk.Label(root,text="CLARO", font=("Colibri", 20,"underline"), bg="#B7D2E9", fg="#000000")
 titulo3.place(x=200,y=285)
  # Crear un Frame que servirá como recuadro para los botones
 marco_botones = tk.Frame(root, relief="sunken", bd=2, bg="#B7D2E9")  # relief="raised" para el borde, bd=2 para el grosor
 marco_botones.place(x=50, y=330, width=400, height=70)  # Posición y tamaño del recuadro

 boton_origen7 = Button(root, text="Fijo", width=5, height=1,
                       font=("Colibri", 16), command=lambda: cargar_Telecom_Fija(barra_progreso), relief="raised", bd=2, bg="#B7D2E9", fg="#000000")
 boton_origen7.place(x=220, y=345)

 boton_origen8 = Button(root, text="Datos", width=5, height=1,
                       font=("Colibri", 16), command=lambda: cargar_Telecom_Datos(barra_progreso), relief="raised", bd=2, bg="#B7D2E9", fg="#000000")
 boton_origen8.place(x=70, y=345)

 boton_origen9 = Button(root, text="Móvil", width=5, height=1,
                       font=("Colibri", 16), command=lambda: cargar_Telecom_Movil(barra_progreso), relief="raised", bd=2, bg="#B7D2E9", fg="#000000")
 boton_origen9.place(x=360, y=345) 

 root = mainloop()