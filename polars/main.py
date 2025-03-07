
from app.auxiliares import cruce_sistecredito, cruce_oms_mercado_pago_clase, cruce_addi_erp
from app.lectores.lector_oms import LectorOMS
from app.lectores.lector_mercadopago import LectorMercadoPago
from app.lectores.lector_erp import LectorERP
from app.lectores.lector_addi import LectorADDI
from app.lectores.lector_mercadolibre import LectorMercadoLibre

from app.lectores.sistecredito.lector_facturas import LectorSisCredFacturas
from app.lectores.sistecredito.lector_pagare import LectorSisCredPagare

from app.lectores.modelos import ConfiguracionLector

from tkinter import filedialog, messagebox, ttk

from app.settings import (RUTA_RAIZ,
                            RUTA_INSUMOS_ADDI,
                            RUTA_INSUMOS_ERP,
                            RUTA_INSUMOS_MP,
                            RUTA_INSUMOS_ML,
                            RUTA_INSUMOS_OMS,
                            RUTA_INSUMOS_FACTURA_SISTECREDITO,
                            RUTA_INSUMOS_PAGARE_SISTECREDITO,
                            RUTA_RESULTADOS)


from datetime import datetime

import customtkinter
import pandas as pd
import tkinter as tk
import polars as pl
import os

#ejecutar_procesamiento()

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")

root = customtkinter.CTk()
root.resizable(False, False)
root.minsize(1200, 400)
root.title('Semi automatización de conciliaciones')

root.grid_columnconfigure(0, pad=20)
root.grid_rowconfigure(0, pad=20)

lector_erp:LectorERP = None
lector_mercadopago:LectorMercadoPago = None
lector_mercadolibre:LectorMercadoLibre = None
lector_oms:LectorOMS = None
lector_addi:LectorADDI = None
lector_pagares_sistecredito:LectorSisCredFacturas = None
lector_facturas_sistecredito:LectorSisCredFacturas = None




#mercado_pago:LecturaMercadoPago = None
#enterprise:LecturaEnterprise = None

ruta_ERP = ""
ruta_MP = ""
ruta_ML = ""
ruta_OMS = ""
ruta_ADDI = ""
ruta_pagare_SC = ""
ruta_factura_SC = ""


ruta_resultados = ""



# def copiar_archivo_a_insumos(ruta_archivo: str):
#     shutil.copy(ruta_archivo, RUTA_INSUMOS)


# def eliminar_archivo_de_insumos(ruta_archivo: str):
#     nombre_archivo = ruta_archivo.split('/')[-1]
#     os.remove(f'{RUTA_INSUMOS}\\{nombre_archivo}')


def seleccionar_archivo_erp(label):

    archivo = filedialog.askopenfilename(
        title="Seleccionar archivo ERP",
        filetypes=[("Excel", "*.xls")],
        initialdir=RUTA_INSUMOS_ERP
    )

    try:
        # Validar que se haya seleccionado un archivo
        if not archivo:
            raise ValueError("No se ha seleccionado ningún archivo para el ERP.")

        global lector_erp, ruta_ERP
        config = ConfiguracionLector(ruta_archivo=archivo)
        lector_erp = LectorERP(config)
        ruta_ERP = archivo


        # Actualizar la etiqueta con la ruta de la carpeta seleccionada
        label.configure(text=archivo)
        check_files()
    except ValueError as e:
        messagebox.showerror('Error', e)
    except Exception as e:
        messagebox.showerror('Error', e)

def seleccionar_archivo_mercadopago(label):

    archivo = filedialog.askopenfilename(
        title="Seleccionar archivo mercadopago",
        filetypes=[("Excel", "*.xlsx")],
        initialdir=RUTA_INSUMOS_MP
    )

    try:
        # Validar que se haya seleccionado un archivo
        if not archivo:
            raise ValueError("No se ha seleccionado ningún archivo de Mercadopago.")
        global lector_mercadopago, ruta_MP
        config = ConfiguracionLector(ruta_archivo=archivo)
        lector_mercadopago = LectorMercadoPago(config)
        ruta_MP = archivo

        # Actualizar la etiqueta con la ruta de la carpeta seleccionada
        label.configure(text=archivo)
        check_files()
    except ValueError as e:
        messagebox.showerror('Error', e)
    except Exception as e:
        messagebox.showerror('Error', e)

def seleccionar_archivo_mercadolibre(label):
    """
        Permite seleccionar una carpeta en lugar de un archivo para mercadolibre.

        La función abre un diálogo para que el usuario seleccione una carpeta. Si no se selecciona ninguna,
        se lanza un error. Una vez seleccionada la carpeta, se actualiza la configuración de la mercadolibre, se instancia
        el lector correspondiente y se actualiza la etiqueta con la ruta de la carpeta seleccionada.

        Args:
            label (tk.Widget): Widget de la interfaz donde se mostrará la ruta seleccionada.

        Returns:
            None
    """
    # Abrir diálogo para seleccionar carpeta únicamente
    ruta_carpeta_ml = filedialog.askdirectory(
        title="Seleccionar carpeta que contiene los archivos de Mercadolibre",
        initialdir=RUTA_INSUMOS_ML
    )

    try:
        # Validar que se haya seleccionado una carpeta
        if not ruta_carpeta_ml:
            raise ValueError("No se ha seleccionado ninguna carpeta para los archivos de Mercadolibre.")

        global lector_mercadolibre, ruta_ML
        # Inicializar la configuración utilizando la ruta de la carpeta seleccionada
        config = ConfiguracionLector(ruta_carpeta=ruta_carpeta_ml)
        lector_mercadolibre = LectorMercadoLibre(config)
        ruta_ML = ruta_carpeta_ml

        # Actualizar la etiqueta con la ruta de la carpeta seleccionada
        label.configure(text=ruta_carpeta_ml)
        check_files()
    except ValueError as error:
        messagebox.showerror('Error', error)
    except Exception as exc:
        messagebox.showerror('Error', exc)

def seleccionar_carpeta_oms(label):
    """
        Permite seleccionar una carpeta en lugar de un archivo para la OMS.

        La función abre un diálogo para que el usuario seleccione una carpeta. Si no se selecciona ninguna,
        se lanza un error. Una vez seleccionada la carpeta, se actualiza la configuración de la OMS, se instancia
        el lector correspondiente y se actualiza la etiqueta con la ruta de la carpeta seleccionada.

        Args:
            label (tk.Widget): Widget de la interfaz donde se mostrará la ruta seleccionada.

        Returns:
            None
    """
    # Abrir diálogo para seleccionar carpeta únicamente
    ruta_carpeta_oms = filedialog.askdirectory(
        title="Seleccionar carpeta que contiene los archivos OMS",
        initialdir=RUTA_INSUMOS_OMS
    )

    try:
        # Validar que se haya seleccionado una carpeta
        if not ruta_carpeta_oms:
            raise ValueError("No se ha seleccionado ninguna carpeta para los archivos de la OMS.")

        global lector_oms, ruta_OMS
        # Inicializar la configuración utilizando la ruta de la carpeta seleccionada
        config = ConfiguracionLector(ruta_carpeta=ruta_carpeta_oms)
        lector_oms = LectorOMS(config)
        ruta_OMS = ruta_carpeta_oms

        # Actualizar la etiqueta con la ruta de la carpeta seleccionada
        label.configure(text=ruta_carpeta_oms)
        check_files()
    except ValueError as error:
        messagebox.showerror('Error', error)
    except Exception as exc:
        messagebox.showerror('Error', exc)

def seleccionar_carpeta_addi(label):
    """
        Permite seleccionar una carpeta en lugar de un archivo para ADDI.

        La función abre un diálogo para que el usuario seleccione una carpeta. Si no se selecciona ninguna,
        se lanza un error. Una vez seleccionada la carpeta, se actualiza la configuración de la ADDI, se instancia
        el lector correspondiente y se actualiza la etiqueta con la ruta de la carpeta seleccionada.

        Args:
            label (tk.Widget): Widget de la interfaz donde se mostrará la ruta seleccionada.

        Returns:
            None
    """
    # Abrir diálogo para seleccionar carpeta únicamente
    ruta_carpeta_addi = filedialog.askdirectory(
        title="Seleccionar carpeta que contiene los archivos de ADDI",
        initialdir=RUTA_INSUMOS_ADDI
    )

    try:
        # Validar que se haya seleccionado una carpeta
        if not ruta_carpeta_addi:
            raise ValueError("No se ha seleccionado ninguna carpeta para los archivos de ADDI.")

        global lector_addi, ruta_ADDI
        # Inicializar la configuración utilizando la ruta de la carpeta seleccionada
        config = ConfiguracionLector(ruta_carpeta=ruta_carpeta_addi)
        lector_addi = LectorADDI(config)
        ruta_ADDI = ruta_carpeta_addi

        # Actualizar la etiqueta con la ruta de la carpeta seleccionada
        label.configure(text=ruta_carpeta_addi)
        check_files()
    except ValueError as error:
        messagebox.showerror('Error', error)
    except Exception as exc:
        messagebox.showerror('Error', exc)

def seleccionar_archivo_factura_sistecredito(label):
    # Abrir diálogo para seleccionar carpeta únicamente
    archivo = filedialog.askopenfilename(
        title="Seleccionar archivo factura de sistecredito",
        filetypes=[("Excel", "*.xlsx")],
        initialdir=RUTA_INSUMOS_FACTURA_SISTECREDITO
    )

    try:
        # Validar que se haya seleccionado un archivo
        if not archivo:
            raise ValueError("No se ha seleccionado ningún archivo de facturas sistecredito.")

        # Validar que se haya seleccionado una carpeta
        global lector_facturas_sistecredito, ruta_factura_SC
        config = ConfiguracionLector(ruta_archivo=archivo)
        lector_facturas_sistecredito = LectorSisCredFacturas(config)
        ruta_factura_SC = archivo
        # Actualizar la etiqueta con la ruta de la carpeta seleccionada
        label.configure(text=archivo)
        check_files()
    except ValueError as error:
        messagebox.showerror('Error', error)
    except Exception as exc:
        messagebox.showerror('Error', exc)

def seleccionar_archivo_pagare_sistecredito(label):
    # Abrir diálogo para seleccionar carpeta únicamente
    archivo = filedialog.askopenfilename(
        title="Seleccionar archivo pagare de sistecredito",
        filetypes=[("Excel", "*.xlsx")],
        initialdir=RUTA_INSUMOS_PAGARE_SISTECREDITO
    )

    #try:
    # Validar que se haya seleccionado un archivo
    if not archivo:
        raise ValueError("No se ha seleccionado ningún archivo de pagares sistecredito.")

    # Validar que se haya seleccionado una carpeta
    global lector_pagares_sistecredito, ruta_pagare_SC

    # Inicializar la configuración utilizando la ruta de la carpeta seleccionada
    config = ConfiguracionLector(ruta_archivo=archivo)
    lector_pagares_sistecredito = LectorSisCredPagare(config)
    ruta_pagare_SC = archivo
    # Actualizar la etiqueta con la ruta de la carpeta seleccionada
    label.configure(text=archivo)
    check_files()
    # except ValueError as error:
    #     messagebox.showerror('Error', error)
    # except Exception as exc:
    #     messagebox.showerror('Error', exc)

def check_files():
    global lector_erp, lector_mercadopago, lector_mercadolibre, lector_oms, lector_addi, lector_pagares_sistecredito, lector_facturas_sistecredito
    #try:
    if lector_erp and lector_mercadopago and  lector_oms and lector_addi and lector_pagares_sistecredito and lector_facturas_sistecredito:
        df_erp = lector_erp.dataframe()
        df_mp = lector_mercadopago.dataframe()
        df_oms = lector_oms.dataframe()
        df_addi = lector_addi.dataframe()
        df_pagare = lector_pagares_sistecredito.dataframe()
        df_factura = lector_facturas_sistecredito.dataframe()

        if df_erp.is_empty() or df_mp.is_empty() or df_oms.is_empty() or df_addi.is_empty() or df_pagare.is_empty() or df_factura.is_empty():
            boton_cruce_todos_los_archivo.configure(state='disabled')
        else:
            boton_cruce_todos_los_archivo.configure(state='normal')

    if lector_erp and lector_mercadopago and  lector_oms:
        df_erp = lector_erp.dataframe()
        df_mp = lector_mercadopago.dataframe()
        df_oms = lector_oms.dataframe()

        if df_erp.is_empty() or df_mp.is_empty() or df_oms.is_empty():
            boton_cruce_principal.configure(state='disabled')
        else:
            boton_cruce_principal.configure(state='normal')

    if lector_erp and lector_oms and lector_pagares_sistecredito and lector_facturas_sistecredito:
        df_erp = lector_erp.dataframe()
        df_oms = lector_oms.dataframe()
        df_pagare = lector_pagares_sistecredito.dataframe()
        df_factura = lector_facturas_sistecredito.dataframe()

        if df_erp.is_empty() or df_oms.is_empty() or df_pagare.is_empty() or df_factura.is_empty():
            boton_cruce_sistecredito.configure(state='disabled')
        else:
            boton_cruce_sistecredito.configure(state='normal')

    if lector_erp and lector_addi:
        df_addi = lector_addi.dataframe()
        df_erp = lector_erp.dataframe()

        if df_addi.is_empty() or df_erp.is_empty():
            boton_cruce_addi.configure(state='disabled')
        else:
            boton_cruce_addi.configure(state='normal')

    # except Exception as e:
    #     messagebox.showerror('Error', e)

def cruzar_todos_los_archivos():
    try:

        df_erp = lector_erp.dataframe()
        df_mp = lector_mercadopago.dataframe()
        df_mercadolibre = lector_mercadolibre.dataframe()
        df_oms = lector_oms.dataframe()
        df_addi = lector_addi.dataframe()
        df_pagare = lector_pagares_sistecredito.dataframe()
        df_factura = lector_facturas_sistecredito.dataframe()

        cruce_sistecredito(df_oms, df_factura, df_pagare, df_erp, RUTA_RESULTADOS)
        cruce_oms_mercado_pago_clase(df_oms, df_mercadolibre, df_mp, df_erp, RUTA_RESULTADOS)
        cruce_addi_erp(df_erp, df_addi, RUTA_RESULTADOS)

        messagebox.showinfo('Info', "Proceso finalizado con éxito. Se reiniciara la aplicación")
        reiniciar()
        root.mainloop()
    except Exception as e:
        messagebox.showerror('Error', e)

def cruzar_archivos_cruce_principal():
    try:

        df_erp = lector_erp.dataframe()
        df_mp = lector_mercadopago.dataframe()
        df_mercadolibre = lector_mercadolibre.dataframe() if lector_mercadolibre else pl.DataFrame()
        df_oms = lector_oms.dataframe()
        cruce_oms_mercado_pago_clase(df_oms, df_mercadolibre, df_mp, df_erp, RUTA_RESULTADOS)

        messagebox.showinfo('Info', "Proceso finalizado con éxito. Se reiniciara la aplicación")
        reiniciar()
        root.mainloop()
    except Exception as e:
        messagebox.showerror('Error', e)


def cruzar_archivos_cruce_sistecredito():
    try:

        df_erp = lector_erp.dataframe()
        df_oms = lector_oms.dataframe()
        df_pagare = lector_pagares_sistecredito.dataframe()
        df_factura = lector_facturas_sistecredito.dataframe()

        cruce_sistecredito(df_oms, df_factura, df_pagare, df_erp, RUTA_RESULTADOS)

        messagebox.showinfo('Info', "Proceso finalizado con éxito. Se reiniciara la aplicación")
        reiniciar()
        root.mainloop()
    except Exception as e:
        messagebox.showerror('Error', e)


def cruzar_archivos_cruce_addi():
    try:

        df_erp = lector_erp.dataframe()
        df_addi = lector_addi.dataframe()

        cruce_addi_erp(df_erp, df_addi, RUTA_RESULTADOS)

        messagebox.showinfo('Info', "Proceso finalizado con éxito. Se reiniciara la aplicación")
        reiniciar()
        root.mainloop()
    except Exception as e:
        messagebox.showerror('Error', e)


def modificar_configuracion(nuevos_valores: dict):
    import json
    import app.settings

    with open(os.path.join(RUTA_RAIZ, 'secrets.json'), 'r') as f:
        variables_ambiente = json.load(f)

    variables_ambiente.update(nuevos_valores)

    with open(os.path.join(RUTA_RAIZ, 'secrets.json'), "w") as f:
        json.dump(variables_ambiente, f, indent=4)

    app.settings.RUTA_RESULTADOS = variables_ambiente['RUTA_RESULTADOS']
    messagebox.showinfo('Info', "Se reiniciara la aplicación para guardar los cambios")
    reiniciar()
    root.mainloop()


def abrir_carpeta_resultados(label):
    global ruta_resultados

    archivo = filedialog.askdirectory(title="Seleccionar nueva carpeta")
    label.configure(text=archivo)


def abrir_explorador(ruta):
    if os.path.exists(ruta):  # Verifica que la ruta exista
        os.startfile(ruta)  # Abre la carpeta en el Explorador de Windows
    else:
        print("La ruta no existe:", ruta)  # Mensaje en consola si la ruta es inválida


def reiniciar():
    import sys
    # cierra la ventana principal
    root.destroy()
    # vuelve a ejecutar el script
    os.execl(sys.executable, sys.executable, *sys.argv)


def abrir_ventana():
    ventana = tk.Toplevel(root, height=500, width=500)
    ventana.grab_set()
    ventana.title('Ventana de configuración')

    frame = customtkinter.CTkFrame(master=ventana, border_width=0, corner_radius=0)
    ruta_resultados_label = customtkinter.CTkLabel(frame, text="RUTA DE ARCHIVOS RESULTADOS: ", anchor="w", justify="left")
    ruta_resultados_label.grid(row=0, column=0, padx=(10,0), pady=10)
    label_resultados_label = customtkinter.CTkLabel(frame, text=RUTA_RESULTADOS, fg_color="#F2F2F2", text_color="black")
    label_resultados_label.grid(row=0, column=1, pady=10)

    boton_cruce = customtkinter.CTkButton(frame, text='Seleccionar nueva carpeta', command=lambda: abrir_carpeta_resultados(label_resultados_label), cursor="hand2")
    boton_cruce.configure(height=30)
    boton_cruce.grid(row=0, column=2, padx=(0, 10))

    boton_aceptar = customtkinter.CTkButton(frame, text='Aceptar', command=lambda: modificar_configuracion({'RUTA_RESULTADOS':label_resultados_label.cget('text')}), cursor="hand2")
    boton_aceptar.configure(height=30)
    boton_aceptar.grid(row=1, column=0, columnspan=3, pady=10)

    frame.grid(row=0, column=0)

    # Aquí puedes agregar los widgets que desees en la ventana


def añadir_labels(frame, fila: int, columna: int):
    ruta_label = customtkinter.CTkLabel(frame, text=" ", anchor="w")
    ruta_label.configure(width=950,
                        height=30,
                        font=customtkinter.CTkFont(size=12, family='Arial'),
                        text_color="black",
                        corner_radius=12,
                        fg_color="#F2F2F2"
                        )
    ruta_label.grid(row=fila, column=columna, padx=(0, 15), pady=10, sticky='we')
    return ruta_label


def añadir_boton(frame, texto: str, fila: int, columna: int, label, funcion):
    boton_archivo = customtkinter.CTkButton(frame,
                                            text=texto.upper(),
                                            command=lambda: funcion(label),
                                            width=240,
                                            height=12,
                                            border_width=0,
                                            border_spacing=8,
                                            corner_radius=24,
                                            anchor="w",
                                            text_color=("gray10", "#DCE4EE"),
                                            font=customtkinter.CTkFont(size=10, weight="bold", family='Arial'),
                                            cursor="hand2")
    boton_archivo.grid(row=fila, column=columna, padx=(15, 0), pady=10, sticky='we')
    return boton_archivo


def anadir_frame_archivo(frame_contenedor, texto: str, funcion_ejectura, fila: int, ruta: str):
    frame = customtkinter.CTkFrame(master=frame_contenedor, fg_color="transparent")
    ruta_label = añadir_labels(frame, fila, 2)
    btn_ERP = añadir_boton(frame, texto, fila, 1, ruta_label, funcion_ejectura)

    # Botón para abrir el Explorador en la ruta especificada
    btn_explorador = customtkinter.CTkButton(
        master=frame,
        text="📂",  # Icono de carpeta para que sea más visual
        width=40,
        fg_color="#FFF9C4",  # Amarillo muy claro
        hover_color="#FFF176",  # Amarillo un poco más intenso al pasar el mouse
        text_color="black",
        command=lambda: abrir_explorador(ruta)  # Se pasa la ruta a la función
    )

    btn_explorador.grid(row=fila, column=3, padx=5, pady=5)  # Botón de abrir carpeta
    frame.grid(row=fila, column=1, sticky='we')



def callback(url):
    import webbrowser
    webbrowser.open_new(url)

boton = customtkinter.CTkButton(root, text='Configuración', command=abrir_ventana, cursor="hand2")
boton.grid(row=0, column=0, padx=10, pady=10, sticky='w')

# Crear botones para seleccionar los archivos ERP
frame_archivos = customtkinter.CTkFrame(master=root, height=100, width=1000, fg_color="#424242")

anadir_frame_archivo(frame_archivos, 'Seleccionar archivo ERP', seleccionar_archivo_erp, fila=0, ruta=RUTA_INSUMOS_ERP)
anadir_frame_archivo(frame_archivos, 'Seleccionar archivo Mercadopago', seleccionar_archivo_mercadopago, fila=1, ruta=RUTA_INSUMOS_MP)
anadir_frame_archivo(frame_archivos, 'Seleccionar carpeta Mercadolibre', seleccionar_archivo_mercadolibre, fila=2, ruta=RUTA_INSUMOS_ML)
anadir_frame_archivo(frame_archivos, 'Seleccionar carpeta OMS', seleccionar_carpeta_oms, fila=3, ruta=RUTA_INSUMOS_OMS)
anadir_frame_archivo(frame_archivos, 'Seleccionar carpeta ADDI', seleccionar_carpeta_addi, fila=4, ruta=RUTA_INSUMOS_ADDI)
anadir_frame_archivo(frame_archivos, 'Seleccionar factura sistecredito', seleccionar_archivo_factura_sistecredito, fila=5, ruta=RUTA_INSUMOS_FACTURA_SISTECREDITO)
anadir_frame_archivo(frame_archivos, 'Seleccionar pagare sistecredito', seleccionar_archivo_pagare_sistecredito, fila=6, ruta=RUTA_INSUMOS_PAGARE_SISTECREDITO)

frame_archivos.grid(row=3, column=0, padx=10, pady=10, sticky='we')

frame_cruce = customtkinter.CTkFrame(master=root, fg_color="#424242")

boton_cruce_todos_los_archivo = customtkinter.CTkButton(frame_cruce, text='Cruzar todos los archivos', command=cruzar_todos_los_archivos, state='disabled', cursor="hand2")
boton_cruce_todos_los_archivo.configure(height=30, font=customtkinter.CTkFont(size=12, weight="bold", family='Arial'))
boton_cruce_todos_los_archivo.grid(row=0, column=0, padx=10, pady=10)

boton_cruce_principal = customtkinter.CTkButton(frame_cruce, text='Realizar cruce OMS', command=cruzar_archivos_cruce_principal, state='disabled', cursor="hand2")
boton_cruce_principal.configure(height=30, font=customtkinter.CTkFont(size=12, weight="bold", family='Arial'))
boton_cruce_principal.grid(row=0, column=1, padx=10, pady=10)

boton_cruce_addi = customtkinter.CTkButton(frame_cruce, text='Realizar cruce ADDI', command=cruzar_archivos_cruce_addi, state='disabled', cursor="hand2")
boton_cruce_addi.configure(height=30, font=customtkinter.CTkFont(size=12, weight="bold", family='Arial'))
boton_cruce_addi.grid(row=0, column=2, padx=10, pady=10)

boton_cruce_sistecredito = customtkinter.CTkButton(frame_cruce, text='Realizar cruce sistecredito', command=cruzar_archivos_cruce_sistecredito, state='disabled', cursor="hand2")
boton_cruce_sistecredito.configure(height=30, font=customtkinter.CTkFont(size=12, weight="bold", family='Arial'))
boton_cruce_sistecredito.grid(row=0, column=3, padx=10, pady=10)

boton_carpeta_resultados = customtkinter.CTkButton(
    frame_cruce,
    text='Resultados 📂',
    command=lambda: abrir_explorador(RUTA_RESULTADOS),
    cursor="hand2",
    fg_color="#DFFFD6",
    hover_color="#B2FF99",  # Verde más intenso al pasar el mouse
    text_color="black"
)

boton_carpeta_resultados.configure(height=30, font=customtkinter.CTkFont(size=12, weight="bold", family='Arial'))
boton_carpeta_resultados.grid(row=0, column=5, padx=10, pady=10, sticky="e")

frame_cruce.grid(row=4, column=0, padx=10, pady=10, columnspan=2, sticky='we')

label_footer = customtkinter.CTkButton(root, text="Desarrollado por Danalytics SAS", command=lambda: callback('https://www.danalyticspro.co/'), hover=False, cursor="hand2")
label_footer.configure(fg_color="transparent", font=("Arial", 14), text_color="white")
label_footer.grid(row=5, column=0, padx=10, pady=10, sticky="we")

root.mainloop()










# def main():
#     """
#     Función principal que lee todos los archivos de la carpeta OMS y los une en un solo DataFrame.
#     """
#     #prueba_cruce_oms_mercado_pago_clase()

#     #prueba_cruce_addi_erp()

#     prueba_sistecredito()

# if __name__ == "__main__":
#     main()