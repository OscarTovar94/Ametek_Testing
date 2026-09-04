"""
Proyecto : Ametek Testing GTAO - GTSOC
Autor    : Oscar Tovar
"""
import tkinter as tk
from tkinter import ttk, messagebox
import sys
import os
import configparser
from datetime import datetime
import time
import win32event
import win32api
import winerror
import serial
from PIL import Image, ImageTk

mutex = win32event.CreateMutex(None, False, "Ametek_Testing")
last_error = win32api.GetLastError()

if last_error == winerror.ERROR_ALREADY_EXISTS:
    sys.exit(0)


def obtener_configuracion(seccion, clave):
    """ Lee el valor de una clave específica en una sección del archivo settings.ini """
    try:
        config = configparser.ConfigParser()
        config.read("settings.ini")

        return config[seccion][clave]

    except FileNotFoundError:
        messagebox.showerror(
            "Error",
            "El archivo de configuración 'settings.ini' no fue encontrado."
        )

    except KeyError:
        messagebox.showerror(
            "Error",
            f"La sección '{seccion}' o la clave '{clave}' no existen."
        )

    except ImportError as e:
        messagebox.showerror(
            "Error",
            f"Ocurrió un error al leer la configuración: {e}"
        )

    return None


def testspec_gtao(seccion, clave):
    """ TestSpec GTAO: Lee el valor de una clave específica en una sección del archivo TestSpec_GTAO.ini """
    try:
        config = configparser.ConfigParser()
        config.read("TestSpec_GTAO.ini")

        return config[seccion][clave]

    except FileNotFoundError:
        messagebox.showerror(
            "Error",
            "El archivo de configuración 'TestSpec_GTAO.ini' no fue encontrado."
        )

    except KeyError:
        messagebox.showerror(
            "Error",
            f"La sección '{seccion}' o la clave '{clave}' no existen."
        )

    except ImportError as e:
        messagebox.showerror(
            "Error",
            f"Ocurrió un error al leer la configuración: {e}"
        )

    return None


def testspec_gtsoc(seccion, clave):
    """ TestSpec GTSOC: Lee el valor de una clave específica en una sección del archivo TestSpec_GTSOC.ini """
    try:
        config = configparser.ConfigParser()
        config.read("TestSpec_GTSOC.ini")

        return config[seccion][clave]

    except FileNotFoundError:
        messagebox.showerror(
            "Error",
            "El archivo de configuración 'TestSpec_GTSOC.ini' no fue encontrado."
        )

    except KeyError:
        messagebox.showerror(
            "Error",
            f"La sección '{seccion}' o la clave '{clave}' no existen."
        )

    except ImportError as e:
        messagebox.showerror(
            "Error",
            f"Ocurrió un error al leer la configuración: {e}"
        )

    return None


DMM = None
PSU = None


def abrir_equipos():
    """Abrir puertos de equipos"""

    global DMM, PSU

    com_dmm = obtener_configuracion('DMM', 'COM')
    com_psu = obtener_configuracion('PSU', 'COM')

    try:
        DMM = serial.Serial(com_dmm, 9600, timeout=1)

    except serial.SerialException:
        messagebox.showerror(
            "DMM Not Found",
            f"Unable to connect to the multimeter ({com_dmm})."
        )
        return

    try:
        PSU = serial.Serial(com_psu, 9600, timeout=1)

    except serial.SerialException:
        messagebox.showerror(
            "Power Supply Not Found",
            f"Unable to connect to the power supply ({com_psu})."
        )
        return


def cerrar_puertos():
    """Cerrar puertos de equipos"""
    global DMM, PSU
    if DMM is not None and DMM.is_open:
        DMM.close()
    if PSU is not None and PSU.is_open:
        PSU.write(b"OUT0\n")
        PSU.close()


class VentanaLogin:
    """Ventana de ingreso de datos para operador y orden."""

    def __init__(self, root):
        self.root = root
        self.root.title("Ingreso de Datos")
        root.iconbitmap("icon_1.ico")

        # Etiquetas
        tk.Label(root, text="Número de empleado:", font=("Arial", 14, "bold")).grid(
            row=0, column=0, padx=10, pady=10, sticky="w"
        )

        tk.Label(root, text="Número de Orden:", font=("Arial", 14, "bold")).grid(
            row=1, column=0, padx=10, pady=10, sticky="w"
        )

        # Entradas
        self.operador = tk.Entry(root, font=("Arial", 14))
        self.operador.grid(row=0, column=1, padx=10, pady=10)
        self.operador.focus_set()
        self.operador.bind("<Return>", lambda event: self.orden.focus_set())

        self.orden = tk.Entry(root, font=("Arial", 14))
        self.orden.grid(row=1, column=1, padx=10, pady=10)
        self.orden.bind("<Return>", lambda event: self.validar_datos())

        MODEL_1 = obtener_configuracion("Part Number", "GTAO")
        MODEL_2 = obtener_configuracion("Part Number", "GTSOC")

        self.modelo = tk.StringVar()

        combo_modelo = ttk.Combobox(
            root,
            textvariable=self.modelo,
            values=[MODEL_1, MODEL_2],
            state="readonly",
            font=("Arial", 14),
            width=15)

        combo_modelo.grid(row=2, column=0, columnspan=2, padx=10, pady=10)
        combo_modelo.current(0)  # Modelo 1 por defecto

        # Botón
        tk.Button(
            root,
            text="Entrar", font=("Arial", 14, "bold"), bg="#BFBFBF", fg="black",
            command=self.validar_datos
        ).grid(row=3, column=0, columnspan=2, pady=15)

        root.update_idletasks()

        ancho = root.winfo_width()
        alto = root.winfo_height()

        x = (root.winfo_screenwidth() // 2) - (ancho // 2)
        y = (root.winfo_screenheight() // 3) - (alto // 2)

        root.geometry(f"+{x}+{y}")

    def validar_datos(self):
        """Valida que los campos de operador y orden no estén vacíos."""
        operador = self.operador.get().strip()
        orden = self.orden.get().strip()
        modelo = self.modelo.get().strip()

        if not operador or not orden:
            messagebox.showwarning(
                "Información faltante",
                "Debe capturar el número de operador y la orden."
            )
            return

        # Cerrar ventana actual
        self.root.destroy()

        # Abrir ventana principal
        ventana_principal = tk.Tk()
        if modelo == obtener_configuracion("Part Number", "GTAO"):
            TestingGTAO(ventana_principal, operador, orden, modelo)
        elif modelo == obtener_configuracion("Part Number", "GTSOC"):
            TestingGTSOC(ventana_principal, operador, orden, modelo)
        ventana_principal.mainloop()


class TestingGTAO:
    """Ventana principal que muestra el operador y la orden."""

    def __init__(self, root, operador, orden, modelo):
        self.root = root
        self.root.title("Ametek Testing GTAO")
        root.iconbitmap("icon_1.ico")
        self.root.state("zoomed")

        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_rowconfigure(2, weight=1)
        self.root.grid_rowconfigure(3, weight=1)
        self.root.grid_rowconfigure(4, weight=1)
        self.root.grid_rowconfigure(5, weight=1)
        self.root.grid_rowconfigure(6, weight=1)
        self.root.grid_rowconfigure(7, weight=1)
        self.root.grid_rowconfigure(8, weight=1)
        self.root.grid_rowconfigure(9, weight=1)
        self.root.grid_rowconfigure(10, weight=1)
        self.root.grid_rowconfigure(11, weight=1)
        self.root.grid_rowconfigure(12, weight=1)
        self.root.grid_rowconfigure(13, weight=1)
        self.root.grid_rowconfigure(14, weight=1)
        self.root.grid_rowconfigure(15, weight=1)
        self.root.grid_rowconfigure(16, weight=1)
        self.root.grid_rowconfigure(17, weight=1)
        self.root.grid_rowconfigure(18, weight=1)
        self.root.grid_rowconfigure(19, weight=1)
        self.root.grid_rowconfigure(20, weight=1)
        self.root.grid_rowconfigure(21, weight=1)
        self.root.grid_rowconfigure(22, weight=1)
        self.root.grid_rowconfigure(23, weight=1)

        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_columnconfigure(2, weight=1)
        self.root.grid_columnconfigure(3, weight=1)

        frame = tk.Frame(self.root)
        frame.grid(row=1, column=1)

        tk.Label(
            frame,
            text="Ametek Testing GTAO",
            font=("Arial", 30, "bold")
        ).grid(row=0, column=0, columnspan=3, pady=20)

        tk.Label(
            frame,
            text=f"#Empleado: {operador}",
            font=("Arial", 14)
        ).grid(row=1, column=0, padx=20, pady=10, sticky="nsew")

        tk.Label(
            frame,
            text=f"#Orden: {orden}",
            font=("Arial", 14)
        ).grid(row=1, column=1, padx=20, pady=10, sticky="nsew")

        tk.Label(
            frame,
            text=f"#Parte: {modelo}",
            font=("Arial", 14)
        ).grid(row=1, column=2, padx=20, pady=10, sticky="nsew")

        entry_id = tk.Entry(frame, font=("Arial", 20), width=30, justify="center", background="springgreen",
                            border=3)
        entry_id.grid(row=2, column=0, columnspan=3,
                      padx=20, pady=(0, 10), sticky="nsew")
        entry_id.focus_set()

        Instrucciones = tk.Label(frame,
                                 text="Escanea el ID de la PCBA para comenzar",
                                 font=("Arial", 14, "italic"),
                                 fg="blue",
                                 width=130,      # ancho fijo en caracteres
                                 height=2,       # alto fijo en líneas
                                 wraplength=1000,  # salto de línea automático
                                 anchor="center",
                                 justify="center")
        Instrucciones.grid(row=3, column=0, columnspan=3, pady=10, padx=0)

        label_id = tk.Label(frame, text="", font=(
            "Arial", 12, "bold"), fg="black")
        label_id.grid(row=4, column=0, columnspan=2,
                      pady=0, padx=0, sticky="w")

        test_1_name = testspec_gtao("+5V_Ohm", "Name")
        test_1_unit = testspec_gtao("+5V_Ohm", "Unit")
        test_1_min = float(testspec_gtao("+5V_Ohm", "Min"))
        test_1_max = float(testspec_gtao("+5V_Ohm", "Max"))

        test_2_name = testspec_gtao("+3V3_Ohm", "Name")
        test_2_unit = testspec_gtao("+3V3_Ohm", "Unit")
        test_2_min = float(testspec_gtao("+3V3_Ohm", "Min"))
        test_2_max = float(testspec_gtao("+3V3_Ohm", "Max"))

        test_3_name = testspec_gtao("FPGA_VccInt_Ohm", "Name")
        test_3_unit = testspec_gtao("FPGA_VccInt_Ohm", "Unit")
        test_3_min = float(testspec_gtao("FPGA_VccInt_Ohm", "Min"))
        test_3_max = float(testspec_gtao("FPGA_VccInt_Ohm", "Max"))

        test_4_name = testspec_gtao("FPGA_VccAux_Ohm", "Name")
        test_4_unit = testspec_gtao("FPGA_VccAux_Ohm", "Unit")
        test_4_min = float(testspec_gtao("FPGA_VccAux_Ohm", "Min"))
        test_4_max = float(testspec_gtao("FPGA_VccAux_Ohm", "Max"))

        test_5_name = testspec_gtao("FPGA_MgtVccAux_Ohm", "Name")
        test_5_unit = testspec_gtao("FPGA_MgtVccAux_Ohm", "Unit")
        test_5_min = float(testspec_gtao("FPGA_MgtVccAux_Ohm", "Min"))
        test_5_max = float(testspec_gtao("FPGA_MgtVccAux_Ohm", "Max"))

        test_6_name = testspec_gtao("FPGA_MgtAVtt_Ohm", "Name")
        test_6_unit = testspec_gtao("FPGA_MgtAVtt_Ohm", "Unit")
        test_6_min = float(testspec_gtao("FPGA_MgtAVtt_Ohm", "Min"))
        test_6_max = float(testspec_gtao("FPGA_MgtAVtt_Ohm", "Max"))

        test_7_name = testspec_gtao("FPGA_MgtAVcc_Ohm", "Name")
        test_7_unit = testspec_gtao("FPGA_MgtAVcc_Ohm", "Unit")
        test_7_min = float(testspec_gtao("FPGA_MgtAVcc_Ohm", "Min"))
        test_7_max = float(testspec_gtao("FPGA_MgtAVcc_Ohm", "Max"))

        test_8_name = testspec_gtao("+12V_Ohm", "Name")
        test_8_unit = testspec_gtao("+12V_Ohm", "Unit")
        test_8_min = float(testspec_gtao("+12V_Ohm", "Min"))
        test_8_max = float(testspec_gtao("+12V_Ohm", "Max"))

        test_9_name = testspec_gtao("-12V_Ohm", "Name")
        test_9_unit = testspec_gtao("-12V_Ohm", "Unit")
        test_9_min = float(testspec_gtao("-12V_Ohm", "Min"))
        test_9_max = float(testspec_gtao("-12V_Ohm", "Max"))

        test_10_name = testspec_gtao("+5V_Voltage", "Name")
        test_10_unit = testspec_gtao("+5V_Voltage", "Unit")
        test_10_min = float(testspec_gtao("+5V_Voltage", "Min"))
        test_10_max = float(testspec_gtao("+5V_Voltage", "Max"))

        test_11_name = testspec_gtao("+3V3_Voltage", "Name")
        test_11_unit = testspec_gtao("+3V3_Voltage", "Unit")
        test_11_min = float(testspec_gtao("+3V3_Voltage", "Min"))
        test_11_max = float(testspec_gtao("+3V3_Voltage", "Max"))

        test_12_name = testspec_gtao("FPGA_VccInt_Voltage", "Name")
        test_12_unit = testspec_gtao("FPGA_VccInt_Voltage", "Unit")
        test_12_min = float(testspec_gtao("FPGA_VccInt_Voltage", "Min"))
        test_12_max = float(testspec_gtao("FPGA_VccInt_Voltage", "Max"))

        test_13_name = testspec_gtao("FPGA_VccAux_Voltage", "Name")
        test_13_unit = testspec_gtao("FPGA_VccAux_Voltage", "Unit")
        test_13_min = float(testspec_gtao("FPGA_VccAux_Voltage", "Min"))
        test_13_max = float(testspec_gtao("FPGA_VccAux_Voltage", "Max"))

        test_14_name = testspec_gtao("FPGA_MgtVccAux_Voltage", "Name")
        test_14_unit = testspec_gtao("FPGA_MgtVccAux_Voltage", "Unit")
        test_14_min = float(testspec_gtao("FPGA_MgtVccAux_Voltage", "Min"))
        test_14_max = float(testspec_gtao("FPGA_MgtVccAux_Voltage", "Max"))

        test_15_name = testspec_gtao("FPGA_MgtAVtt_Voltage", "Name")
        test_15_unit = testspec_gtao("FPGA_MgtAVtt_Voltage", "Unit")
        test_15_min = float(testspec_gtao("FPGA_MgtAVtt_Voltage", "Min"))
        test_15_max = float(testspec_gtao("FPGA_MgtAVtt_Voltage", "Max"))

        test_16_name = testspec_gtao("FPGA_MgtAVcc_Voltage", "Name")
        test_16_unit = testspec_gtao("FPGA_MgtAVcc_Voltage", "Unit")
        test_16_min = float(testspec_gtao("FPGA_MgtAVcc_Voltage", "Min"))
        test_16_max = float(testspec_gtao("FPGA_MgtAVcc_Voltage", "Max"))

        label_test1 = tk.Label(frame, text=f"Test 1: {test_1_name} - Min: {test_1_min} {test_1_unit}, Max: {test_1_max} {test_1_unit}", font=(
            "Arial", 12), fg="black")
        label_test1.grid(row=5, column=0, columnspan=2,
                         pady=0, padx=0, sticky="w")

        label_test2 = tk.Label(frame, text=f"Test 2: {test_2_name} - Min: {test_2_min} {test_2_unit}, Max: {test_2_max} {test_2_unit}", font=(
            "Arial", 12), fg="black")
        label_test2.grid(row=6, column=0, columnspan=2,
                         pady=0, padx=0, sticky="w")

        label_test3 = tk.Label(frame, text=f"Test 3: {test_3_name} - Min: {test_3_min} {test_3_unit}, Max: {test_3_max} {test_3_unit}", font=(
            "Arial", 12), fg="black")
        label_test3.grid(row=7, column=0, columnspan=2,
                         pady=0, padx=0, sticky="w")

        label_test4 = tk.Label(frame, text=f"Test 4: {test_4_name} - Min: {test_4_min} {test_4_unit}, Max: {test_4_max} {test_4_unit}", font=(
            "Arial", 12), fg="black")
        label_test4.grid(row=8, column=0, columnspan=2,
                         pady=0, padx=0, sticky="w")

        label_test5 = tk.Label(frame, text=f"Test 5: {test_5_name} - Min: {test_5_min} {test_5_unit}, Max: {test_5_max} {test_5_unit}", font=(
            "Arial", 12), fg="black")
        label_test5.grid(row=9, column=0, columnspan=2,
                         pady=0, padx=0, sticky="w")

        label_test6 = tk.Label(frame, text=f"Test 6: {test_6_name} - Min: {test_6_min} {test_6_unit}, Max: {test_6_max} {test_6_unit}", font=(
            "Arial", 12), fg="black")
        label_test6.grid(row=10, column=0, columnspan=2,
                         pady=0, padx=0, sticky="w")

        label_test7 = tk.Label(frame, text=f"Test 7: {test_7_name} - Min: {test_7_min} {test_7_unit}, Max: {test_7_max} {test_7_unit}", font=(
            "Arial", 12), fg="black")
        label_test7.grid(row=11, column=0, columnspan=3,
                         pady=0, padx=0, sticky="w")

        label_test8 = tk.Label(frame, text=f"Test 8: {test_8_name} - Min: {test_8_min} {test_8_unit}, Max: {test_8_max} {test_8_unit}", font=(
            "Arial", 12), fg="black")
        label_test8.grid(row=12, column=0, columnspan=2,
                         pady=0, padx=0, sticky="w")

        label_test9 = tk.Label(frame, text=f"Test 9: {test_9_name} - Min: {test_9_min} {test_9_unit}, Max: {test_9_max} {test_9_unit}", font=(
            "Arial", 12), fg="black")
        label_test9.grid(row=13, column=0, columnspan=2,
                         pady=0, padx=0, sticky="w")

        label_test10 = tk.Label(frame, text=f"Test 10: {test_10_name} - Min: {test_10_min} {test_10_unit}, Max: {test_10_max} {test_10_unit}", font=(
            "Arial", 12), fg="black")
        label_test10.grid(row=14, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_test11 = tk.Label(frame, text=f"Test 11: {test_11_name} - Min: {test_11_min} {test_11_unit}, Max: {test_11_max} {test_11_unit}", font=(
            "Arial", 12), fg="black")
        label_test11.grid(row=15, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_test12 = tk.Label(frame, text=f"Test 12: {test_12_name} - Min: {test_12_min} {test_12_unit}, Max: {test_12_max} {test_12_unit}", font=(
            "Arial", 12), fg="black")
        label_test12.grid(row=16, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_test13 = tk.Label(frame, text=f"Test 13: {test_13_name} - Min: {test_13_min} {test_13_unit}, Max: {test_13_max} {test_13_unit}", font=(
            "Arial", 12), fg="black")
        label_test13.grid(row=17, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_test14 = tk.Label(frame, text=f"Test 14: {test_14_name} - Min: {test_14_min} {test_14_unit}, Max: {test_14_max} {test_14_unit}", font=(
            "Arial", 12), fg="black")
        label_test14.grid(row=18, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_test15 = tk.Label(frame, text=f"Test 15: {test_15_name} - Min: {test_15_min} {test_15_unit}, Max: {test_15_max} {test_15_unit}", font=(
            "Arial", 12), fg="black")
        label_test15.grid(row=19, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_test16 = tk.Label(frame, text=f"Test 16: {test_16_name} - Min: {test_16_min} {test_16_unit}, Max: {test_16_max} {test_16_unit}", font=(
            "Arial", 12), fg="black")
        label_test16.grid(row=20, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_resultado = tk.Label(frame, text="", font=(
            "Arial", 20), fg="black")
        label_resultado.grid(row=21, column=0, columnspan=2,
                             pady=10, padx=0, sticky="nsew")

        imagen = Image.open("GTAO.png")
        imagen = imagen.resize((600, 500))
        foto = ImageTk.PhotoImage(imagen)

        label_image_gtao = tk.Label(frame, image=foto)
        label_image_gtao.image = foto
        label_image_gtao.grid(row=4, column=2, columnspan=2, rowspan=19,
                              pady=0, padx=0, sticky="e")

        def reiniciar_prueba():
            """Reinicia la prueba"""
            cerrar_puertos()
            root.unbind("<space>")
            entry_id.config(state="normal")
            # entry_id.delete(0, tk.END)
            entry_id.focus_set()
            Instrucciones.config(
                text="Escanea el ID de la PCBA para comenzar", bg="SystemButtonFace", fg="blue")
            label_id.config(text="", bg="SystemButtonFace", fg="black")
            label_test1.config(
                text=f"Test 1: {test_1_name} - Min: {test_1_min} {test_1_unit}, Max: {test_1_max} {test_1_unit}", bg="SystemButtonFace", fg="black")
            label_test2.config(
                text=f"Test 2: {test_2_name} - Min: {test_2_min} {test_2_unit}, Max: {test_2_max} {test_2_unit}", bg="SystemButtonFace", fg="black")
            label_test3.config(
                text=f"Test 3: {test_3_name} - Min: {test_3_min} {test_3_unit}, Max: {test_3_max} {test_3_unit}", bg="SystemButtonFace", fg="black")
            label_test4.config(
                text=f"Test 4: {test_4_name} - Min: {test_4_min} {test_4_unit}, Max: {test_4_max} {test_4_unit}", bg="SystemButtonFace", fg="black")
            label_test5.config(
                text=f"Test 5: {test_5_name} - Min: {test_5_min} {test_5_unit}, Max: {test_5_max} {test_5_unit}", bg="SystemButtonFace", fg="black")
            label_test6.config(
                text=f"Test 6: {test_6_name} - Min: {test_6_min} {test_6_unit}, Max: {test_6_max} {test_6_unit}", bg="SystemButtonFace", fg="black")
            label_test7.config(
                text=f"Test 7: {test_7_name} - Min: {test_7_min} {test_7_unit}, Max: {test_7_max} {test_7_unit}", bg="SystemButtonFace", fg="black")
            label_test8.config(
                text=f"Test 8: {test_8_name} - Min: {test_8_min} {test_8_unit}, Max: {test_8_max} {test_8_unit}", bg="SystemButtonFace", fg="black")
            label_test9.config(
                text=f"Test 9: {test_9_name} - Min: {test_9_min} {test_9_unit}, Max: {test_9_max} {test_9_unit}", bg="SystemButtonFace", fg="black")
            label_test10.config(
                text=f"Test 10: {test_10_name} - Min: {test_10_min} {test_10_unit}, Max: {test_10_max} {test_10_unit}", bg="SystemButtonFace", fg="black")
            label_test11.config(
                text=f"Test 11: {test_11_name} - Min: {test_11_min} {test_11_unit}, Max: {test_11_max} {test_11_unit}", bg="SystemButtonFace", fg="black")
            label_test12.config(
                text=f"Test 12: {test_12_name} - Min: {test_12_min} {test_12_unit}, Max: {test_12_max} {test_12_unit}", bg="SystemButtonFace", fg="black")
            label_test13.config(
                text=f"Test 13: {test_13_name} - Min: {test_13_min} {test_13_unit}, Max: {test_13_max} {test_13_unit}", bg="SystemButtonFace", fg="black")
            label_test14.config(
                text=f"Test 14: {test_14_name} - Min: {test_14_min} {test_14_unit}, Max: {test_14_max} {test_14_unit}", bg="SystemButtonFace", fg="black")
            label_test15.config(
                text=f"Test 15: {test_15_name} - Min: {test_15_min} {test_15_unit}, Max: {test_15_max} {test_15_unit}", bg="SystemButtonFace", fg="black")
            label_test16.config(
                text=f"Test 16: {test_16_name} - Min: {test_16_min} {test_16_unit}, Max: {test_16_max} {test_16_unit}", bg="SystemButtonFace", fg="black")
            label_resultado.config(text="", bg="SystemButtonFace", fg="black")

        tk.Button(
            root,
            text="Reiniciar Prueba", font=("Arial", 14, "bold"), bg="#BFBFBF", fg="black",
            command=reiniciar_prueba
        ).grid(row=20, column=0, columnspan=3,
               pady=(0, 10), padx=50, sticky="nsew")

        def validar_id(event=None):
            reiniciar_prueba()
            id_value = entry_id.get().strip()
            if not id_value:
                messagebox.showwarning(
                    "Información faltante",
                    "Debe capturar el ID."
                )
                return

            if len(id_value) == 16 and id_value[:6] == str(modelo).strip():
                label_id.config(
                    text=f"{id_value}", bg="#C6EFCE", fg="green")
                entry_id.delete(0, tk.END)
                entry_id.config(state="disabled")
                test_1_gtao()
            else:
                messagebox.showerror(
                    "ID Inválido", f"El ID ingresado no es válido: {id_value}")
                entry_id.delete(0, tk.END)
                Instrucciones.config(
                    text="Escanea el ID de la PCBA para comenzar")

        def conexion_pcba():
            """Instrucciones para la conexión de la PCBA"""
            root.focus_set()
            Instrucciones.config(
                text="Precaución\nConecta los cables/arneses en la PCBA según la WI y presiona la barra espaciadora para iniciar.", bg="#FFC7CE")

            root.bind("<space>", test_10_gtao)

        def test_1_gtao(event=None):
            abrir_equipos()
            label_resultado.config(text="En proceso...",
                                   bg="#FFEB9C", fg="#9C5700")
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_1_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_1(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_1_name} {test_1_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro)

            def configurar_multimetro():
                try:
                    DMM.write(b"CONF:RES\n")
                    delay = int(testspec_gtao("Delay_Ohm", "delay"))
                    root.after(delay, leer_resultado)
                except Exception:
                    messagebox.showerror(
                        "No encontrado",
                        f"No se puede conectar al Multímetro y/o Fuente de Alimentación."
                    )
                    self.root.destroy()
                    return

            def leer_resultado():
                DMM.write(b"MEAS:RES?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_1_min <= resultado <= test_1_max:
                    label_test1.config(
                        text=f"Test 1: {test_1_name} - Min: {test_1_min} {test_1_unit}, Max: {test_1_max} {test_1_unit} - Result: PASS ({resultado:.4f} {test_1_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    test_2_gtao()
                else:
                    label_test1.config(
                        text=f"Test 1: {test_1_name} - Min: {test_1_min} {test_1_unit}, Max: {test_1_max} {test_1_unit} - Result: FAIL ({resultado:.4f} {test_1_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_1)

        def test_2_gtao(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_2_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_2(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_2_name} {test_2_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro)

            def configurar_multimetro():
                # DMM.write(b"CONF:RES\n")
                delay = int(testspec_gtao("Delay_Ohm", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"MEAS:RES?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_2_min <= resultado <= test_2_max:
                    label_test2.config(
                        text=f"Test 2: {test_2_name} - Min: {test_2_min} {test_2_unit}, Max: {test_2_max} {test_2_unit} - Result: PASS ({resultado:.4f} {test_2_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    test_3_gtao()
                else:
                    label_test2.config(
                        text=f"Test 2: {test_2_name} - Min: {test_2_min} {test_2_unit}, Max: {test_2_max} {test_2_unit} - Result: FAIL ({resultado:.4f} {test_2_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_2)

        def test_3_gtao(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_3_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_3(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_3_name} {test_3_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro)

            def configurar_multimetro():
                # DMM.write(b"CONF:RES\n")
                delay = int(testspec_gtao("Delay_Ohm", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"MEAS:RES?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_3_min <= resultado <= test_3_max:
                    label_test3.config(
                        text=f"Test 3: {test_3_name} - Min: {test_3_min} {test_3_unit}, Max: {test_3_max} {test_3_unit} - Result: PASS ({resultado:.4f} {test_3_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    test_4_gtao()
                else:
                    label_test3.config(
                        text=f"Test 3: {test_3_name} - Min: {test_3_min} {test_3_unit}, Max: {test_3_max} {test_3_unit} - Result: FAIL ({resultado:.4f} {test_3_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_3)

        def test_4_gtao(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_4_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_4(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_4_name} {test_4_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro)

            def configurar_multimetro():
                # DMM.write(b"CONF:RES\n")
                delay = int(testspec_gtao("Delay_Ohm", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"MEAS:RES?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_4_min <= resultado <= test_4_max:
                    label_test4.config(
                        text=f"Test 4: {test_4_name} - Min: {test_4_min} {test_4_unit}, Max: {test_4_max} {test_4_unit} - Result: PASS ({resultado:.4f} {test_4_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    test_5_gtao()
                else:
                    label_test4.config(
                        text=f"Test 4: {test_4_name} - Min: {test_4_min} {test_4_unit}, Max: {test_4_max} {test_4_unit} - Result: FAIL ({resultado:.4f} {test_4_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_4)

        def test_5_gtao(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_5_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_5(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_5_name} {test_5_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro)

            def configurar_multimetro():
                # DMM.write(b"CONF:RES\n")
                delay = int(testspec_gtao("Delay_Ohm", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"MEAS:RES?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_5_min <= resultado <= test_5_max:
                    label_test5.config(
                        text=f"Test 5: {test_5_name} - Min: {test_5_min} {test_5_unit}, Max: {test_5_max} {test_5_unit} - Result: PASS ({resultado:.4f} {test_5_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    test_6_gtao()
                else:
                    label_test5.config(
                        text=f"Test 5: {test_5_name} - Min: {test_5_min} {test_5_unit}, Max: {test_5_max} {test_5_unit} - Result: FAIL ({resultado:.4f} {test_5_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_5)

        def test_6_gtao(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_6_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_6(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_6_name} {test_6_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro)

            def configurar_multimetro():
                # DMM.write(b"CONF:RES\n")
                delay = int(testspec_gtao("Delay_Ohm", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"MEAS:RES?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_6_min <= resultado <= test_6_max:
                    label_test6.config(
                        text=f"Test 6: {test_6_name} - Min: {test_6_min} {test_6_unit}, Max: {test_6_max} {test_6_unit} - Result: PASS ({resultado:.4f} {test_6_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    test_7_gtao()
                else:
                    label_test6.config(
                        text=f"Test 6: {test_6_name} - Min: {test_6_min} {test_6_unit}, Max: {test_6_max} {test_6_unit} - Result: FAIL ({resultado:.4f} {test_6_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_6)

        def test_7_gtao(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_7_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_7(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_7_name} {test_7_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro)

            def configurar_multimetro():
                # DMM.write(b"CONF:RES\n")
                delay = int(testspec_gtao("Delay_Ohm", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"MEAS:RES?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_7_min <= resultado <= test_7_max:
                    label_test7.config(
                        text=f"Test 7: {test_7_name} - Min: {test_7_min} {test_7_unit}, Max: {test_7_max} {test_7_unit} - Result: PASS ({resultado:.4f} {test_7_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    test_8_gtao()
                else:
                    label_test7.config(
                        text=f"Test 7: {test_7_name} - Min: {test_7_min} {test_7_unit}, Max: {test_7_max} {test_7_unit} - Result: FAIL ({resultado:.4f} {test_7_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_7)

        def test_8_gtao(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_8_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_8(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_8_name} {test_8_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro)

            def configurar_multimetro():
                # DMM.write(b"CONF:RES\n")
                delay = int(testspec_gtao("Delay_Ohm", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"MEAS:RES?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_8_min <= resultado <= test_8_max:
                    label_test8.config(
                        text=f"Test 8: {test_8_name} - Min: {test_8_min} {test_8_unit}, Max: {test_8_max} {test_8_unit} - Result: PASS ({resultado:.4f} {test_8_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    test_9_gtao()
                else:
                    label_test8.config(
                        text=f"Test 8: {test_8_name} - Min: {test_8_min} {test_8_unit}, Max: {test_8_max} {test_8_unit} - Result: FAIL ({resultado:.4f} {test_8_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_8)

        def test_9_gtao(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_9_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_9(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_9_name} {test_9_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro)

            def configurar_multimetro():
                # DMM.write(b"CONF:RES\n")
                delay = int(testspec_gtao("Delay_Ohm", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"MEAS:RES?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_9_min <= resultado <= test_9_max:
                    label_test9.config(
                        text=f"Test 9: {test_9_name} - Min: {test_9_min} {test_9_unit}, Max: {test_9_max} {test_9_unit} - Result: PASS ({resultado:.4f} {test_9_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    conexion_pcba()
                else:
                    label_test9.config(
                        text=f"Test 9: {test_9_name} - Min: {test_9_min} {test_9_unit}, Max: {test_9_max} {test_9_unit} - Result: FAIL ({resultado:.4f} {test_9_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_9)

        def test_10_gtao(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_10_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_10(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_10_name} {test_10_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro_psu)

            volatage_gtao = float(testspec_gtao("Voltage", "Voltage"))
            current_gtao = float(testspec_gtao("Current", "Current"))

            def configurar_multimetro_psu():
                DMM.write(b"CONF:VOLT:DC\n")
                time.sleep(0.1)
                PSU.write(b"VSET2:0\n")
                time.sleep(0.1)
                PSU.write(b"ISET2:0\n")
                time.sleep(0.1)
                PSU.write(f"VSET1:{volatage_gtao}\n".encode())
                time.sleep(0.1)
                PSU.write(f"ISET1:{current_gtao}\n".encode())
                time.sleep(0.1)
                PSU.write(b"OUT1\n")

                delay = int(testspec_gtao("Delay_Voltage", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"READ?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_10_min <= resultado <= test_10_max:
                    label_test10.config(
                        text=f"Test 10: {test_10_name} - Min: {test_10_min} {test_10_unit}, Max: {test_10_max} {test_10_unit} - Result: PASS ({resultado:.4f} {test_10_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    PSU.write(b"OUT0\n")
                    test_11_gtao()
                else:
                    label_test10.config(
                        text=f"Test 10: {test_10_name} - Min: {test_10_min} {test_10_unit}, Max: {test_10_max} {test_10_unit} - Result: FAIL ({resultado:.4f} {test_10_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_10)

        def test_11_gtao(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_11_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_11(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_11_name} {test_11_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro_psu)

            def configurar_multimetro_psu():
                PSU.write(b"OUT1\n")
                delay = int(testspec_gtao("Delay_Voltage", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"READ?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_11_min <= resultado <= test_11_max:
                    label_test11.config(
                        text=f"Test 11: {test_11_name} - Min: {test_11_min} {test_11_unit}, Max: {test_11_max} {test_11_unit} - Result: PASS ({resultado:.4f} {test_11_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    PSU.write(b"OUT0\n")
                    test_12_gtao()
                else:
                    label_test11.config(
                        text=f"Test 11: {test_11_name} - Min: {test_11_min} {test_11_unit}, Max: {test_11_max} {test_11_unit} - Result: FAIL ({resultado:.4f} {test_11_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_11)

        def test_12_gtao(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_12_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_12(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_12_name} {test_12_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro_psu)

            def configurar_multimetro_psu():
                PSU.write(b"OUT1\n")
                delay = int(testspec_gtao("Delay_Voltage", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"READ?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_12_min <= resultado <= test_12_max:
                    label_test12.config(
                        text=f"Test 12: {test_12_name} - Min: {test_12_min} {test_12_unit}, Max: {test_12_max} {test_12_unit} - Result: PASS ({resultado:.4f} {test_12_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    PSU.write(b"OUT0\n")
                    test_13_gtao()
                else:
                    label_test12.config(
                        text=f"Test 12: {test_12_name} - Min: {test_12_min} {test_12_unit}, Max: {test_12_max} {test_12_unit} - Result: FAIL ({resultado:.4f} {test_12_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_12)

        def test_13_gtao(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_13_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_13(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_13_name} {test_13_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro_psu)

            def configurar_multimetro_psu():
                PSU.write(b"OUT1\n")
                delay = int(testspec_gtao("Delay_Voltage", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"READ?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_13_min <= resultado <= test_13_max:
                    label_test13.config(
                        text=f"Test 13: {test_13_name} - Min: {test_13_min} {test_13_unit}, Max: {test_13_max} {test_13_unit} - Result: PASS ({resultado:.4f} {test_13_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    test_14_gtao()
                    PSU.write(b"OUT0\n")
                else:
                    label_test13.config(
                        text=f"Test 13: {test_13_name} - Min: {test_13_min} {test_13_unit}, Max: {test_13_max} {test_13_unit} - Result: FAIL ({resultado:.4f} {test_13_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_13)

        def test_14_gtao(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_14_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_14(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_14_name} {test_14_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro_psu)

            def configurar_multimetro_psu():
                PSU.write(b"OUT1\n")
                delay = int(testspec_gtao("Delay_Voltage", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"READ?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_14_min <= resultado <= test_14_max:
                    label_test14.config(
                        text=f"Test 14: {test_14_name} - Min: {test_14_min} {test_14_unit}, Max: {test_14_max} {test_14_unit} - Result: PASS ({resultado:.4f} {test_14_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    PSU.write(b"OUT0\n")
                    test_15_gtao()
                else:
                    label_test14.config(
                        text=f"Test 14: {test_14_name} - Min: {test_14_min} {test_14_unit}, Max: {test_14_max} {test_14_unit} - Result: FAIL ({resultado:.4f} {test_14_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_14)

        def test_15_gtao(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_15_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_15(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_15_name} {test_15_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro_psu)

            def configurar_multimetro_psu():
                PSU.write(b"OUT1\n")
                delay = int(testspec_gtao("Delay_Voltage", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"READ?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_15_min <= resultado <= test_15_max:
                    label_test15.config(
                        text=f"Test 15: {test_15_name} - Min: {test_15_min} {test_15_unit}, Max: {test_15_max} {test_15_unit} - Result: PASS ({resultado:.4f} {test_15_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    PSU.write(b"OUT0\n")
                    test_16_gtao()
                else:
                    label_test15.config(
                        text=f"Test 15: {test_15_name} - Min: {test_15_min} {test_15_unit}, Max: {test_15_max} {test_15_unit} - Result: FAIL ({resultado:.4f} {test_15_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_15)

        def test_16_gtao(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_16_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_16(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_16_name} {test_16_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro_psu)

            def configurar_multimetro_psu():
                PSU.write(b"OUT1\n")
                delay = int(testspec_gtao("Delay_Voltage", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"READ?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_16_min <= resultado <= test_16_max:
                    label_test16.config(
                        text=f"Test 16: {test_16_name} - Min: {test_16_min} {test_16_unit}, Max: {test_16_max} {test_16_unit} - Result: PASS ({resultado:.4f} {test_16_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    PSU.write(b"OUT0\n")
                    test_pass()
                else:
                    label_test16.config(
                        text=f"Test 16: {test_16_name} - Min: {test_16_min} {test_16_unit}, Max: {test_16_max} {test_16_unit} - Result: FAIL ({resultado:.4f} {test_16_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_16)

        def test_fail(event=None):
            """En caso de falla en algún test, esta función permitirá reiniciar la prueba"""
            label_resultado.config(text="TEST FAIL", bg="#FFC7CE", fg="red")
            cerrar_puertos()
            Instrucciones.config(
                text="Escanea el ID de la PCBA para comenzar", bg="SystemButtonFace", fg="blue")
            entry_id.config(state="normal")
            entry_id.delete(0, tk.END)
            root.unbind("<space>")
            guardar_resultados()
            entry_id.focus_set()

        def test_pass(event=None):
            """En caso de falla en algún test, esta función permitirá reiniciar la prueba"""
            label_resultado.config(text="TEST PASS", bg="#C6EFCE", fg="green")
            cerrar_puertos()
            Instrucciones.config(
                text="Escanea el ID de la PCBA para comenzar", bg="SystemButtonFace", fg="blue")
            entry_id.config(state="normal")
            entry_id.delete(0, tk.END)
            root.unbind("<space>")
            guardar_resultados()
            entry_id.focus_set()

        def guardar_resultados():
            ahora = datetime.now()

            fecha = ahora.strftime("%Y%m%d")
            hora = ahora.strftime("%H%M%S")

            pieza = label_id.cget("text")
            resultado = label_resultado.cget("text")
            num_operador = operador
            num_orden = orden

            nombre_archivo = f"{pieza}_{fecha}_{hora}.txt"

            carpeta = "Result_GTAO"

            os.makedirs(carpeta, exist_ok=True)

            ruta = os.path.join(carpeta, nombre_archivo)

            with open(ruta, "w", encoding="utf-8") as archivo:

                archivo.write(f"ID: {pieza}\n")
                archivo.write(
                    f"Date: {ahora.strftime('%Y-%m-%d %H:%M:%S')}\n")
                archivo.write(
                    f"Employee: {num_operador}\n")
                archivo.write(
                    f"Work order: {num_orden}\n")
                archivo.write(
                    f"Result: {resultado}\n\n")

                archivo.write(label_test1.cget("text") + "\n")
                archivo.write(label_test2.cget("text") + "\n")
                archivo.write(label_test3.cget("text") + "\n")
                archivo.write(label_test4.cget("text") + "\n")
                archivo.write(label_test5.cget("text") + "\n")
                archivo.write(label_test6.cget("text") + "\n")
                archivo.write(label_test7.cget("text") + "\n")
                archivo.write(label_test8.cget("text") + "\n")
                archivo.write(label_test9.cget("text") + "\n")
                archivo.write(label_test10.cget("text") + "\n")
                archivo.write(label_test11.cget("text") + "\n")
                archivo.write(label_test12.cget("text") + "\n")
                archivo.write(label_test13.cget("text") + "\n")
                archivo.write(label_test14.cget("text") + "\n")
                archivo.write(label_test15.cget("text") + "\n")
                archivo.write(label_test16.cget("text") + "\n")

        entry_id.bind("<Return>", validar_id)


class TestingGTSOC:
    """Ventana principal que muestra el operador y la orden."""

    def __init__(self, root, operador, orden, modelo):
        self.root = root
        self.root.title("Ametek Testing GTSOC")
        root.iconbitmap("icon_1.ico")
        self.root.state("zoomed")

        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_rowconfigure(2, weight=1)
        self.root.grid_rowconfigure(3, weight=1)
        self.root.grid_rowconfigure(4, weight=1)
        self.root.grid_rowconfigure(5, weight=1)
        self.root.grid_rowconfigure(6, weight=1)
        self.root.grid_rowconfigure(7, weight=1)
        self.root.grid_rowconfigure(8, weight=1)
        self.root.grid_rowconfigure(9, weight=1)
        self.root.grid_rowconfigure(10, weight=1)
        self.root.grid_rowconfigure(11, weight=1)
        self.root.grid_rowconfigure(12, weight=1)
        self.root.grid_rowconfigure(13, weight=1)
        self.root.grid_rowconfigure(14, weight=1)
        self.root.grid_rowconfigure(15, weight=1)
        self.root.grid_rowconfigure(16, weight=1)
        self.root.grid_rowconfigure(17, weight=1)
        self.root.grid_rowconfigure(18, weight=1)
        self.root.grid_rowconfigure(19, weight=1)
        self.root.grid_rowconfigure(20, weight=1)
        self.root.grid_rowconfigure(21, weight=1)
        self.root.grid_rowconfigure(22, weight=1)
        self.root.grid_rowconfigure(23, weight=1)
        self.root.grid_rowconfigure(24, weight=1)
        self.root.grid_rowconfigure(25, weight=1)
        self.root.grid_rowconfigure(26, weight=1)
        self.root.grid_rowconfigure(27, weight=1)
        self.root.grid_rowconfigure(28, weight=1)
        self.root.grid_rowconfigure(29, weight=1)
        self.root.grid_rowconfigure(30, weight=1)
        self.root.grid_rowconfigure(31, weight=1)
        self.root.grid_rowconfigure(32, weight=1)
        self.root.grid_rowconfigure(33, weight=1)

        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_columnconfigure(2, weight=1)
        self.root.grid_columnconfigure(2, weight=1)

        frame = tk.Frame(self.root)
        frame.grid(row=1, column=1)

        tk.Label(
            frame,
            text="Ametek Testing GTSOC",
            font=("Arial", 30, "bold")
        ).grid(row=0, column=0, columnspan=3, pady=20)

        tk.Label(
            frame,
            text=f"#Empleado: {operador}",
            font=("Arial", 14)
        ).grid(row=1, column=0, padx=20, pady=10, sticky="nsew")

        tk.Label(
            frame,
            text=f"#Orden: {orden}",
            font=("Arial", 14)
        ).grid(row=1, column=1, padx=20, pady=10, sticky="nsew")

        tk.Label(
            frame,
            text=f"#Parte: {modelo}",
            font=("Arial", 14)
        ).grid(row=1, column=2, padx=20, pady=10, sticky="nsew")

        entry_id = tk.Entry(frame, font=("Arial", 20), width=30, justify="center", background="springgreen",
                            border=3)
        entry_id.grid(row=2, column=0, columnspan=3,
                      padx=20, pady=(0, 10), sticky="nsew")
        entry_id.focus_set()

        Instrucciones = tk.Label(frame,
                                 text="Escanea el ID de la PCBA para comenzar",
                                 font=("Arial", 14, "italic"),
                                 fg="blue",
                                 width=130,      # ancho fijo en caracteres
                                 height=2,       # alto fijo en líneas
                                 wraplength=1000,  # salto de línea automático
                                 anchor="center",
                                 justify="center")
        Instrucciones.grid(row=3, column=0, columnspan=3, pady=10, padx=0)

        label_id = tk.Label(frame, text="", font=(
            "Arial", 12, "bold"), fg="black")
        label_id.grid(row=4, column=0, columnspan=2,
                      pady=0, padx=0, sticky="w")

        test_1_name = testspec_gtsoc("UTIL_3V3_Ohm", "Name")
        test_1_unit = testspec_gtsoc("UTIL_3V3_Ohm", "Unit")
        test_1_min = float(testspec_gtsoc("UTIL_3V3_Ohm", "Min"))
        test_1_max = float(testspec_gtsoc("UTIL_3V3_Ohm", "Max"))

        test_2_name = testspec_gtsoc("UTIL_1V13_Ohm", "Name")
        test_2_unit = testspec_gtsoc("UTIL_1V13_Ohm", "Unit")
        test_2_min = float(testspec_gtsoc("UTIL_1V13_Ohm", "Min"))
        test_2_max = float(testspec_gtsoc("UTIL_1V13_Ohm", "Max"))

        test_3_name = testspec_gtsoc("UTIL_2V5_Ohm", "Name")
        test_3_unit = testspec_gtsoc("UTIL_2V5_Ohm", "Unit")
        test_3_min = float(testspec_gtsoc("UTIL_2V5_Ohm", "Min"))
        test_3_max = float(testspec_gtsoc("UTIL_2V5_Ohm", "Max"))

        test_4_name = testspec_gtsoc("VCCINT_Ohm", "Name")
        test_4_unit = testspec_gtsoc("VCCINT_Ohm", "Unit")
        test_4_min = float(testspec_gtsoc("VCCINT_Ohm", "Min"))
        test_4_max = float(testspec_gtsoc("VCCINT_Ohm", "Max"))

        test_5_name = testspec_gtsoc("VCCPSINT_Ohm", "Name")
        test_5_unit = testspec_gtsoc("VCCPSINT_Ohm", "Unit")
        test_5_min = float(testspec_gtsoc("VCCPSINT_Ohm", "Min"))
        test_5_max = float(testspec_gtsoc("VCCPSINT_Ohm", "Max"))

        test_6_name = testspec_gtsoc("VCCAUX_Ohm", "Name")
        test_6_unit = testspec_gtsoc("VCCAUX_Ohm", "Unit")
        test_6_min = float(testspec_gtsoc("VCCAUX_Ohm", "Min"))
        test_6_max = float(testspec_gtsoc("VCCAUX_Ohm", "Max"))

        test_7_name = testspec_gtsoc("MGTAVCC_Ohm", "Name")
        test_7_unit = testspec_gtsoc("MGTAVCC_Ohm", "Unit")
        test_7_min = float(testspec_gtsoc("MGTAVCC_Ohm", "Min"))
        test_7_max = float(testspec_gtsoc("MGTAVCC_Ohm", "Max"))

        test_8_name = testspec_gtsoc("MGTAVTT_Ohm", "Name")
        test_8_unit = testspec_gtsoc("MGTAVTT_Ohm", "Unit")
        test_8_min = float(testspec_gtsoc("MGTAVTT_Ohm", "Min"))
        test_8_max = float(testspec_gtsoc("MGTAVTT_Ohm", "Max"))

        test_9_name = testspec_gtsoc("MGTVCCAUX_Ohm", "Name")
        test_9_unit = testspec_gtsoc("MGTVCCAUX_Ohm", "Unit")
        test_9_min = float(testspec_gtsoc("MGTVCCAUX_Ohm", "Min"))
        test_9_max = float(testspec_gtsoc("MGTVCCAUX_Ohm", "Max"))

        test_10_name = testspec_gtsoc("VCCOPS_Ohm", "Name")
        test_10_unit = testspec_gtsoc("VCCOPS_Ohm", "Unit")
        test_10_min = float(testspec_gtsoc("VCCOPS_Ohm", "Min"))
        test_10_max = float(testspec_gtsoc("VCCOPS_Ohm", "Max"))

        test_11_name = testspec_gtsoc("VCC3V3_Ohm", "Name")
        test_11_unit = testspec_gtsoc("VCC3V3_Ohm", "Unit")
        test_11_min = float(testspec_gtsoc("VCC3V3_Ohm", "Min"))
        test_11_max = float(testspec_gtsoc("VCC3V3_Ohm", "Max"))

        test_12_name = testspec_gtsoc("DDR4_DIMM_VDDQ_Ohm", "Name")
        test_12_unit = testspec_gtsoc("DDR4_DIMM_VDDQ_Ohm", "Unit")
        test_12_min = float(testspec_gtsoc("DDR4_DIMM_VDDQ_Ohm", "Min"))
        test_12_max = float(testspec_gtsoc("DDR4_DIMM_VDDQ_Ohm", "Max"))

        test_13_name = testspec_gtsoc("DDR4_VTT_Ohm", "Name")
        test_13_unit = testspec_gtsoc("DDR4_VTT_Ohm", "Unit")
        test_13_min = float(testspec_gtsoc("DDR4_VTT_Ohm", "Min"))
        test_13_max = float(testspec_gtsoc("DDR4_VTT_Ohm", "Max"))

        test_14_name = testspec_gtsoc("UTIL_3V3_Voltage", "Name")
        test_14_unit = testspec_gtsoc("UTIL_3V3_Voltage", "Unit")
        test_14_min = float(testspec_gtsoc("UTIL_3V3_Voltage", "Min"))
        test_14_max = float(testspec_gtsoc("UTIL_3V3_Voltage", "Max"))

        test_15_name = testspec_gtsoc("UTIL_1V13_Voltage", "Name")
        test_15_unit = testspec_gtsoc("UTIL_1V13_Voltage", "Unit")
        test_15_min = float(testspec_gtsoc("UTIL_1V13_Voltage", "Min"))
        test_15_max = float(testspec_gtsoc("UTIL_1V13_Voltage", "Max"))

        test_16_name = testspec_gtsoc("UTIL_2V5_Voltage", "Name")
        test_16_unit = testspec_gtsoc("UTIL_2V5_Voltage", "Unit")
        test_16_min = float(testspec_gtsoc("UTIL_2V5_Voltage", "Min"))
        test_16_max = float(testspec_gtsoc("UTIL_2V5_Voltage", "Max"))

        test_17_name = testspec_gtsoc("VCCINT_Voltage", "Name")
        test_17_unit = testspec_gtsoc("VCCINT_Voltage", "Unit")
        test_17_min = float(testspec_gtsoc("VCCINT_Voltage", "Min"))
        test_17_max = float(testspec_gtsoc("VCCINT_Voltage", "Max"))

        test_18_name = testspec_gtsoc("VCCPSINT_Voltage", "Name")
        test_18_unit = testspec_gtsoc("VCCPSINT_Voltage", "Unit")
        test_18_min = float(testspec_gtsoc("VCCPSINT_Voltage", "Min"))
        test_18_max = float(testspec_gtsoc("VCCPSINT_Voltage", "Max"))

        test_19_name = testspec_gtsoc("VCCAUX_Voltage", "Name")
        test_19_unit = testspec_gtsoc("VCCAUX_Voltage", "Unit")
        test_19_min = float(testspec_gtsoc("VCCAUX_Voltage", "Min"))
        test_19_max = float(testspec_gtsoc("VCCAUX_Voltage", "Max"))

        test_20_name = testspec_gtsoc("MGTAVCC_Voltage", "Name")
        test_20_unit = testspec_gtsoc("MGTAVCC_Voltage", "Unit")
        test_20_min = float(testspec_gtsoc("MGTAVCC_Voltage", "Min"))
        test_20_max = float(testspec_gtsoc("MGTAVCC_Voltage", "Max"))

        test_21_name = testspec_gtsoc("MGTAVTT_Voltage", "Name")
        test_21_unit = testspec_gtsoc("MGTAVTT_Voltage", "Unit")
        test_21_min = float(testspec_gtsoc("MGTAVTT_Voltage", "Min"))
        test_21_max = float(testspec_gtsoc("MGTAVTT_Voltage", "Max"))

        test_22_name = testspec_gtsoc("MGTVCCAUX_Voltage", "Name")
        test_22_unit = testspec_gtsoc("MGTVCCAUX_Voltage", "Unit")
        test_22_min = float(testspec_gtsoc("MGTVCCAUX_Voltage", "Min"))
        test_22_max = float(testspec_gtsoc("MGTVCCAUX_Voltage", "Max"))

        test_23_name = testspec_gtsoc("VCCOPS_Voltage", "Name")
        test_23_unit = testspec_gtsoc("VCCOPS_Voltage", "Unit")
        test_23_min = float(testspec_gtsoc("VCCOPS_Voltage", "Min"))
        test_23_max = float(testspec_gtsoc("VCCOPS_Voltage", "Max"))

        test_24_name = testspec_gtsoc("VCC3V3_Voltage", "Name")
        test_24_unit = testspec_gtsoc("VCC3V3_Voltage", "Unit")
        test_24_min = float(testspec_gtsoc("VCC3V3_Voltage", "Min"))
        test_24_max = float(testspec_gtsoc("VCC3V3_Voltage", "Max"))

        test_25_name = testspec_gtsoc("DDR4_DIMM_VDDQ_Voltage", "Name")
        test_25_unit = testspec_gtsoc("DDR4_DIMM_VDDQ_Voltage", "Unit")
        test_25_min = float(testspec_gtsoc("DDR4_DIMM_VDDQ_Voltage", "Min"))
        test_25_max = float(testspec_gtsoc("DDR4_DIMM_VDDQ_Voltage", "Max"))

        test_26_name = testspec_gtsoc("DDR4_VTT_Voltage", "Name")
        test_26_unit = testspec_gtsoc("DDR4_VTT_Voltage", "Unit")
        test_26_min = float(testspec_gtsoc("DDR4_VTT_Voltage", "Min"))
        test_26_max = float(testspec_gtsoc("DDR4_VTT_Voltage", "Max"))

        label_test1 = tk.Label(frame, text=f"Test 1: {test_1_name} - Min: {test_1_min} {test_1_unit}, Max: {test_1_max} {test_1_unit}", font=(
            "Arial", 12), fg="black")
        label_test1.grid(row=5, column=0, columnspan=2,
                         pady=0, padx=0, sticky="w")

        label_test2 = tk.Label(frame, text=f"Test 2: {test_2_name} - Min: {test_2_min} {test_2_unit}, Max: {test_2_max} {test_2_unit}", font=(
            "Arial", 12), fg="black")
        label_test2.grid(row=6, column=0, columnspan=2,
                         pady=0, padx=0, sticky="w")

        label_test3 = tk.Label(frame, text=f"Test 3: {test_3_name} - Min: {test_3_min} {test_3_unit}, Max: {test_3_max} {test_3_unit}", font=(
            "Arial", 12), fg="black")
        label_test3.grid(row=7, column=0, columnspan=2,
                         pady=0, padx=0, sticky="w")

        label_test4 = tk.Label(frame, text=f"Test 4: {test_4_name} - Min: {test_4_min} {test_4_unit}, Max: {test_4_max} {test_4_unit}", font=(
            "Arial", 12), fg="black")
        label_test4.grid(row=8, column=0, columnspan=2,
                         pady=0, padx=0, sticky="w")

        label_test5 = tk.Label(frame, text=f"Test 5: {test_5_name} - Min: {test_5_min} {test_5_unit}, Max: {test_5_max} {test_5_unit}", font=(
            "Arial", 12), fg="black")
        label_test5.grid(row=9, column=0, columnspan=2,
                         pady=0, padx=0, sticky="w")

        label_test6 = tk.Label(frame, text=f"Test 6: {test_6_name} - Min: {test_6_min} {test_6_unit}, Max: {test_6_max} {test_6_unit}", font=(
            "Arial", 12), fg="black")
        label_test6.grid(row=10, column=0, columnspan=2,
                         pady=0, padx=0, sticky="w")

        label_test7 = tk.Label(frame, text=f"Test 7: {test_7_name} - Min: {test_7_min} {test_7_unit}, Max: {test_7_max} {test_7_unit}", font=(
            "Arial", 12), fg="black")
        label_test7.grid(row=11, column=0, columnspan=2,
                         pady=0, padx=0, sticky="w")

        label_test8 = tk.Label(frame, text=f"Test 8: {test_8_name} - Min: {test_8_min} {test_8_unit}, Max: {test_8_max} {test_8_unit}", font=(
            "Arial", 12), fg="black")
        label_test8.grid(row=12, column=0, columnspan=2,
                         pady=0, padx=0, sticky="w")

        label_test9 = tk.Label(frame, text=f"Test 9: {test_9_name} - Min: {test_9_min} {test_9_unit}, Max: {test_9_max} {test_9_unit}", font=(
            "Arial", 12), fg="black")
        label_test9.grid(row=13, column=0, columnspan=2,
                         pady=0, padx=0, sticky="w")

        label_test10 = tk.Label(frame, text=f"Test 10: {test_10_name} - Min: {test_10_min} {test_10_unit}, Max: {test_10_max} {test_10_unit}", font=(
            "Arial", 12), fg="black")
        label_test10.grid(row=14, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_test11 = tk.Label(frame, text=f"Test 11: {test_11_name} - Min: {test_11_min} {test_11_unit}, Max: {test_11_max} {test_11_unit}", font=(
            "Arial", 12), fg="black")
        label_test11.grid(row=15, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_test12 = tk.Label(frame, text=f"Test 12: {test_12_name} - Min: {test_12_min} {test_12_unit}, Max: {test_12_max} {test_12_unit}", font=(
            "Arial", 12), fg="black")
        label_test12.grid(row=16, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_test13 = tk.Label(frame, text=f"Test 13: {test_13_name} - Min: {test_13_min} {test_13_unit}, Max: {test_13_max} {test_13_unit}", font=(
            "Arial", 12), fg="black")
        label_test13.grid(row=17, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_test14 = tk.Label(frame, text=f"Test 14: {test_14_name} - Min: {test_14_min} {test_14_unit}, Max: {test_14_max} {test_14_unit}", font=(
            "Arial", 12), fg="black")
        label_test14.grid(row=18, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_test15 = tk.Label(frame, text=f"Test 15: {test_15_name} - Min: {test_15_min} {test_15_unit}, Max: {test_15_max} {test_15_unit}", font=(
            "Arial", 12), fg="black")
        label_test15.grid(row=19, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_test16 = tk.Label(frame, text=f"Test 16: {test_16_name} - Min: {test_16_min} {test_16_unit}, Max: {test_16_max} {test_16_unit}", font=(
            "Arial", 12), fg="black")
        label_test16.grid(row=20, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_test17 = tk.Label(frame, text=f"Test 17: {test_17_name} - Min: {test_17_min} {test_17_unit}, Max: {test_17_max} {test_17_unit}", font=(
            "Arial", 12), fg="black")
        label_test17.grid(row=21, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_test18 = tk.Label(frame, text=f"Test 18: {test_18_name} - Min: {test_18_min} {test_18_unit}, Max: {test_18_max} {test_18_unit}", font=(
            "Arial", 12), fg="black")
        label_test18.grid(row=22, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_test19 = tk.Label(frame, text=f"Test 19: {test_19_name} - Min: {test_19_min} {test_19_unit}, Max: {test_19_max} {test_19_unit}", font=(
            "Arial", 12), fg="black")
        label_test19.grid(row=23, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_test20 = tk.Label(frame, text=f"Test 20: {test_20_name} - Min: {test_20_min} {test_20_unit}, Max: {test_20_max} {test_20_unit}", font=(
            "Arial", 12), fg="black")
        label_test20.grid(row=24, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_test21 = tk.Label(frame, text=f"Test 21: {test_21_name} - Min: {test_21_min} {test_21_unit}, Max: {test_21_max} {test_21_unit}", font=(
            "Arial", 12), fg="black")
        label_test21.grid(row=25, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_test22 = tk.Label(frame, text=f"Test 22: {test_22_name} - Min: {test_22_min} {test_22_unit}, Max: {test_22_max} {test_22_unit}", font=(
            "Arial", 12), fg="black")
        label_test22.grid(row=26, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_test23 = tk.Label(frame, text=f"Test 23: {test_23_name} - Min: {test_23_min} {test_23_unit}, Max: {test_23_max} {test_23_unit}", font=(
            "Arial", 12), fg="black")
        label_test23.grid(row=27, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_test24 = tk.Label(frame, text=f"Test 24: {test_24_name} - Min: {test_24_min} {test_24_unit}, Max: {test_24_max} {test_24_unit}", font=(
            "Arial", 12), fg="black")
        label_test24.grid(row=28, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_test25 = tk.Label(frame, text=f"Test 25: {test_25_name} - Min: {test_25_min} {test_25_unit}, Max: {test_25_max} {test_25_unit}", font=(
            "Arial", 12), fg="black")
        label_test25.grid(row=29, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_test26 = tk.Label(frame, text=f"Test 26: {test_26_name} - Min: {test_26_min} {test_26_unit}, Max: {test_26_max} {test_26_unit}", font=(
            "Arial", 12), fg="black")
        label_test26.grid(row=30, column=0, columnspan=2,
                          pady=0, padx=0, sticky="w")

        label_resultado = tk.Label(frame, text="", font=(
            "Arial", 20), fg="black")
        label_resultado.grid(row=31, column=0, columnspan=2,
                             pady=10, padx=0, sticky="nsew")

        imagen = Image.open("GTSOC.png")
        imagen = imagen.resize((600, 500))
        foto = ImageTk.PhotoImage(imagen)

        label_image_gtsoc = tk.Label(frame, image=foto)
        label_image_gtsoc.image = foto
        label_image_gtsoc.grid(row=4, column=2, columnspan=2, rowspan=33,
                               pady=0, padx=0, sticky="e")

        def reiniciar_prueba():
            """Reinicia la prueba"""
            cerrar_puertos()
            root.unbind("<space>")
            entry_id.config(state="normal")
            # entry_id.delete(0, tk.END)
            entry_id.focus_set()
            Instrucciones.config(
                text="Escanea el ID de la PCBA para comenzar", bg="SystemButtonFace", fg="blue")
            label_id.config(text="", bg="SystemButtonFace", fg="black")
            label_test1.config(
                text=f"Test 1: {test_1_name} - Min: {test_1_min} {test_1_unit}, Max: {test_1_max} {test_1_unit}", bg="SystemButtonFace", fg="black")
            label_test2.config(
                text=f"Test 2: {test_2_name} - Min: {test_2_min} {test_2_unit}, Max: {test_2_max} {test_2_unit}", bg="SystemButtonFace", fg="black")
            label_test3.config(
                text=f"Test 3: {test_3_name} - Min: {test_3_min} {test_3_unit}, Max: {test_3_max} {test_3_unit}", bg="SystemButtonFace", fg="black")
            label_test4.config(
                text=f"Test 4: {test_4_name} - Min: {test_4_min} {test_4_unit}, Max: {test_4_max} {test_4_unit}", bg="SystemButtonFace", fg="black")
            label_test5.config(
                text=f"Test 5: {test_5_name} - Min: {test_5_min} {test_5_unit}, Max: {test_5_max} {test_5_unit}", bg="SystemButtonFace", fg="black")
            label_test6.config(
                text=f"Test 6: {test_6_name} - Min: {test_6_min} {test_6_unit}, Max: {test_6_max} {test_6_unit}", bg="SystemButtonFace", fg="black")
            label_test7.config(
                text=f"Test 7: {test_7_name} - Min: {test_7_min} {test_7_unit}, Max: {test_7_max} {test_7_unit}", bg="SystemButtonFace", fg="black")
            label_test8.config(
                text=f"Test 8: {test_8_name} - Min: {test_8_min} {test_8_unit}, Max: {test_8_max} {test_8_unit}", bg="SystemButtonFace", fg="black")
            label_test9.config(
                text=f"Test 9: {test_9_name} - Min: {test_9_min} {test_9_unit}, Max: {test_9_max} {test_9_unit}", bg="SystemButtonFace", fg="black")
            label_test10.config(
                text=f"Test 10: {test_10_name} - Min: {test_10_min} {test_10_unit}, Max: {test_10_max} {test_10_unit}", bg="SystemButtonFace", fg="black")
            label_test11.config(
                text=f"Test 11: {test_11_name} - Min: {test_11_min} {test_11_unit}, Max: {test_11_max} {test_11_unit}", bg="SystemButtonFace", fg="black")
            label_test12.config(
                text=f"Test 12: {test_12_name} - Min: {test_12_min} {test_12_unit}, Max: {test_12_max} {test_12_unit}", bg="SystemButtonFace", fg="black")
            label_test13.config(
                text=f"Test 13: {test_13_name} - Min: {test_13_min} {test_13_unit}, Max: {test_13_max} {test_13_unit}", bg="SystemButtonFace", fg="black")
            label_test14.config(
                text=f"Test 14: {test_14_name} - Min: {test_14_min} {test_14_unit}, Max: {test_14_max} {test_14_unit}", bg="SystemButtonFace", fg="black")
            label_test15.config(
                text=f"Test 15: {test_15_name} - Min: {test_15_min} {test_15_unit}, Max: {test_15_max} {test_15_unit}", bg="SystemButtonFace", fg="black")
            label_test16.config(
                text=f"Test 16: {test_16_name} - Min: {test_16_min} {test_16_unit}, Max: {test_16_max} {test_16_unit}", bg="SystemButtonFace", fg="black")
            label_test17.config(
                text=f"Test 17: {test_17_name} - Min: {test_17_min} {test_17_unit}, Max: {test_17_max} {test_17_unit}", bg="SystemButtonFace", fg="black")
            label_test18.config(
                text=f"Test 18: {test_18_name} - Min: {test_18_min} {test_18_unit}, Max: {test_18_max} {test_18_unit}",
                bg="SystemButtonFace", fg="black")
            label_test19.config(
                text=f"Test 19: {test_19_name} - Min: {test_19_min} {test_19_unit}, Max: {test_19_max} {test_19_unit}",
                bg="SystemButtonFace", fg="black")
            label_test20.config(
                text=f"Test 20: {test_20_name} - Min: {test_20_min} {test_20_unit}, Max: {test_20_max} {test_20_unit}",
                bg="SystemButtonFace",
                fg="black"
            )
            label_test21.config(
                text=f"Test 21: {test_21_name} - Min: {test_21_min} {test_21_unit}, Max: {test_21_max} {test_21_unit}",
                bg="SystemButtonFace",
                fg="black"
            )
            label_test22.config(
                text=f"Test 22: {test_22_name} - Min: {test_22_min} {test_22_unit}, Max: {test_22_max} {test_22_unit}",
                bg="SystemButtonFace",
                fg="black"
            )
            label_test23.config(
                text=f"Test 23: {test_23_name} - Min: {test_23_min} {test_23_unit}, Max: {test_23_max} {test_23_unit}",
                bg="SystemButtonFace",
                fg="black"
            )
            label_test24.config(
                text=f"Test 24: {test_24_name} - Min: {test_24_min} {test_24_unit}, Max: {test_24_max} {test_24_unit}",
                bg="SystemButtonFace",
                fg="black"
            )
            label_test25.config(
                text=f"Test 25: {test_25_name} - Min: {test_25_min} {test_25_unit}, Max: {test_25_max} {test_25_unit}",
                bg="SystemButtonFace",
                fg="black"
            )
            label_test26.config(
                text=f"Test 26: {test_26_name} - Min: {test_26_min} {test_26_unit}, Max: {test_26_max} {test_26_unit}",
                bg="SystemButtonFace",
                fg="black"
            )
            label_resultado.config(text="", bg="SystemButtonFace", fg="black")

        tk.Button(
            root,
            text="Reiniciar Prueba", font=("Arial", 12, "bold"), bg="#BFBFBF", fg="black",
            command=reiniciar_prueba
        ).grid(row=33, column=0, columnspan=3,
               pady=(0, 10), padx=50, sticky="nsew")

        def validar_id(event=None):
            reiniciar_prueba()
            id_value = entry_id.get().strip()
            if not id_value:
                messagebox.showwarning(
                    "Información faltante",
                    "Debe capturar el ID."
                )
                return

            if len(id_value) == 16 and id_value[:6] == str(modelo).strip():
                label_id.config(
                    text=f"{id_value}", bg="#C6EFCE", fg="green")
                entry_id.delete(0, tk.END)
                entry_id.config(state="disabled")
                test_1_gtsoc()
            else:
                messagebox.showerror(
                    "ID Inválido", f"El ID ingresado no es válido: {id_value}")
                entry_id.delete(0, tk.END)
                Instrucciones.config(
                    text="Escanea el ID de la PCBA para comenzar")

        def conexion_pcba():
            """Instrucciones para la conexión de la PCBA"""
            root.focus_set()
            Instrucciones.config(
                text="Precaución\nConecta los cables/arneses en la PCBA según la WI y presiona la barra espaciadora para iniciar.", bg="#FFC7CE")

            root.bind("<space>", test_14_gtsoc)

        def test_1_gtsoc(event=None):
            abrir_equipos()
            label_resultado.config(text="En proceso...",
                                   bg="#FFEB9C", fg="#9C5700")
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_1_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_1(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_1_name} {test_1_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro)

            def configurar_multimetro():
                try:
                    DMM.write(b"CONF:RES\n")
                    delay = int(testspec_gtsoc("Delay_Ohm", "delay"))
                    root.after(delay, leer_resultado)
                except Exception:
                    messagebox.showerror(
                        "No encontrado",
                        f"No se puede conectar al Multímetro y/o Fuente de Alimentación."
                    )
                    self.root.destroy()
                    return

            def leer_resultado():
                DMM.write(b"MEAS:RES?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_1_min <= resultado <= test_1_max:
                    label_test1.config(
                        text=f"Test 1: {test_1_name} - Min: {test_1_min} {test_1_unit}, Max: {test_1_max} {test_1_unit} - Result: PASS ({resultado:.4f} {test_1_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    test_2_gtsoc()
                else:
                    label_test1.config(
                        text=f"Test 1: {test_1_name} - Min: {test_1_min} {test_1_unit}, Max: {test_1_max} {test_1_unit} - Result: FAIL ({resultado:.4f} {test_1_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_1)

        def test_2_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_2_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_2(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_2_name} {test_2_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro)

            def configurar_multimetro():
                # DMM.write(b"CONF:RES\n")
                delay = int(testspec_gtsoc("Delay_Ohm", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"MEAS:RES?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_2_min <= resultado <= test_2_max:
                    label_test2.config(
                        text=f"Test 2: {test_2_name} - Min: {test_2_min} {test_2_unit}, Max: {test_2_max} {test_2_unit} - Result: PASS ({resultado:.4f} {test_2_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    test_3_gtsoc()
                else:
                    label_test2.config(
                        text=f"Test 2: {test_2_name} - Min: {test_2_min} {test_2_unit}, Max: {test_2_max} {test_2_unit} - Result: FAIL ({resultado:.4f} {test_2_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_2)

        def test_3_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_3_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_3(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_3_name} {test_3_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro)

            def configurar_multimetro():
                # DMM.write(b"CONF:RES\n")
                delay = int(testspec_gtsoc("Delay_Ohm", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"MEAS:RES?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_3_min <= resultado <= test_3_max:
                    label_test3.config(
                        text=f"Test 3: {test_3_name} - Min: {test_3_min} {test_3_unit}, Max: {test_3_max} {test_3_unit} - Result: PASS ({resultado:.4f} {test_3_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    test_4_gtsoc()
                else:
                    label_test3.config(
                        text=f"Test 3: {test_3_name} - Min: {test_3_min} {test_3_unit}, Max: {test_3_max} {test_3_unit} - Result: FAIL ({resultado:.4f} {test_3_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_3)

        def test_4_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_4_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_4(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_4_name} {test_4_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro)

            def configurar_multimetro():
                # DMM.write(b"CONF:RES\n")
                delay = int(testspec_gtsoc("Delay_Ohm", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"MEAS:RES?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_4_min <= resultado <= test_4_max:
                    label_test4.config(
                        text=f"Test 4: {test_4_name} - Min: {test_4_min} {test_4_unit}, Max: {test_4_max} {test_4_unit} - Result: PASS ({resultado:.4f} {test_4_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    test_5_gtsoc()
                else:
                    label_test4.config(
                        text=f"Test 4: {test_4_name} - Min: {test_4_min} {test_4_unit}, Max: {test_4_max} {test_4_unit} - Result: FAIL ({resultado:.4f} {test_4_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_4)

        def test_5_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_5_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_5(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_5_name} {test_5_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro)

            def configurar_multimetro():
                # DMM.write(b"CONF:RES\n")
                delay = int(testspec_gtsoc("Delay_Ohm", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"MEAS:RES?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_5_min <= resultado <= test_5_max:
                    label_test5.config(
                        text=f"Test 5: {test_5_name} - Min: {test_5_min} {test_5_unit}, Max: {test_5_max} {test_5_unit} - Result: PASS ({resultado:.4f} {test_5_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    test_6_gtsoc()
                else:
                    label_test5.config(
                        text=f"Test 5: {test_5_name} - Min: {test_5_min} {test_5_unit}, Max: {test_5_max} {test_5_unit} - Result: FAIL ({resultado:.4f} {test_5_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_5)

        def test_6_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_6_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_6(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_6_name} {test_6_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro)

            def configurar_multimetro():
                # DMM.write(b"CONF:RES\n")
                delay = int(testspec_gtsoc("Delay_Ohm", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"MEAS:RES?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_6_min <= resultado <= test_6_max:
                    label_test6.config(
                        text=f"Test 6: {test_6_name} - Min: {test_6_min} {test_6_unit}, Max: {test_6_max} {test_6_unit} - Result: PASS ({resultado:.4f} {test_6_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    test_7_gtsoc()
                else:
                    label_test6.config(
                        text=f"Test 6: {test_6_name} - Min: {test_6_min} {test_6_unit}, Max: {test_6_max} {test_6_unit} - Result: FAIL ({resultado:.4f} {test_6_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_6)

        def test_7_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_7_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_7(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_7_name} {test_7_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro)

            def configurar_multimetro():
                # DMM.write(b"CONF:RES\n")
                delay = int(testspec_gtsoc("Delay_Ohm", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"MEAS:RES?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_7_min <= resultado <= test_7_max:
                    label_test7.config(
                        text=f"Test 7: {test_7_name} - Min: {test_7_min} {test_7_unit}, Max: {test_7_max} {test_7_unit} - Result: PASS ({resultado:.4f} {test_7_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    test_8_gtsoc()
                else:
                    label_test7.config(
                        text=f"Test 7: {test_7_name} - Min: {test_7_min} {test_7_unit}, Max: {test_7_max} {test_7_unit} - Result: FAIL ({resultado:.4f} {test_7_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_7)

        def test_8_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_8_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_8(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_8_name} {test_8_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro)

            def configurar_multimetro():
                # DMM.write(b"CONF:RES\n")
                delay = int(testspec_gtsoc("Delay_Ohm", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"MEAS:RES?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_8_min <= resultado <= test_8_max:
                    label_test8.config(
                        text=f"Test 8: {test_8_name} - Min: {test_8_min} {test_8_unit}, Max: {test_8_max} {test_8_unit} - Result: PASS ({resultado:.4f} {test_8_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    test_9_gtsoc()
                else:
                    label_test8.config(
                        text=f"Test 8: {test_8_name} - Min: {test_8_min} {test_8_unit}, Max: {test_8_max} {test_8_unit} - Result: FAIL ({resultado:.4f} {test_8_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_8)

        def test_9_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_9_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_9(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_9_name} {test_9_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro)

            def configurar_multimetro():
                # DMM.write(b"CONF:RES\n")
                delay = int(testspec_gtsoc("Delay_Ohm", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"MEAS:RES?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_9_min <= resultado <= test_9_max:
                    label_test9.config(
                        text=f"Test 9: {test_9_name} - Min: {test_9_min} {test_9_unit}, Max: {test_9_max} {test_9_unit} - Result: PASS ({resultado:.4f} {test_9_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    test_10_gtsoc()
                else:
                    label_test9.config(
                        text=f"Test 9: {test_9_name} - Min: {test_9_min} {test_9_unit}, Max: {test_9_max} {test_9_unit} - Result: FAIL ({resultado:.4f} {test_9_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_9)

        def test_10_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_10_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_10(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_10_name} {test_10_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro)

            def configurar_multimetro():
                # DMM.write(b"CONF:RES\n")
                delay = int(testspec_gtsoc("Delay_Ohm", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"MEAS:RES?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_10_min <= resultado <= test_10_max:
                    label_test10.config(
                        text=f"Test 10: {test_10_name} - Min: {test_10_min} {test_10_unit}, Max: {test_10_max} {test_10_unit} - Result: PASS ({resultado:.4f} {test_10_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    test_11_gtsoc()
                else:
                    label_test10.config(
                        text=f"Test 10: {test_10_name} - Min: {test_10_min} {test_10_unit}, Max: {test_10_max} {test_10_unit} - Result: FAIL ({resultado:.4f} {test_10_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_10)

        def test_11_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_11_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_11(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_11_name} {test_11_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro)

            def configurar_multimetro():
                # DMM.write(b"CONF:RES\n")
                delay = int(testspec_gtsoc("Delay_Ohm", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"MEAS:RES?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_11_min <= resultado <= test_11_max:
                    label_test11.config(
                        text=f"Test 11: {test_11_name} - Min: {test_11_min} {test_11_unit}, Max: {test_11_max} {test_11_unit} - Result: PASS ({resultado:.4f} {test_11_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    test_12_gtsoc()
                else:
                    label_test11.config(
                        text=f"Test 11: {test_11_name} - Min: {test_11_min} {test_11_unit}, Max: {test_11_max} {test_11_unit} - Result: FAIL ({resultado:.4f} {test_11_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_11)

        def test_12_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_12_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_12(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_12_name} {test_12_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro)

            def configurar_multimetro():
                # DMM.write(b"CONF:RES\n")
                delay = int(testspec_gtsoc("Delay_Ohm", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"MEAS:RES?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_12_min <= resultado <= test_12_max:
                    label_test12.config(
                        text=f"Test 12: {test_12_name} - Min: {test_12_min} {test_12_unit}, Max: {test_12_max} {test_12_unit} - Result: PASS ({resultado:.4f} {test_12_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    test_13_gtsoc()
                else:
                    label_test12.config(
                        text=f"Test 12: {test_12_name} - Min: {test_12_min} {test_12_unit}, Max: {test_12_max} {test_12_unit} - Result: FAIL ({resultado:.4f} {test_12_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_12)

        def test_13_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_13_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_13(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_13_name} {test_13_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro)

            def configurar_multimetro():
                # DMM.write(b"CONF:RES\n")
                delay = int(testspec_gtsoc("Delay_Ohm", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"MEAS:RES?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_13_min <= resultado <= test_13_max:
                    label_test13.config(
                        text=f"Test 13: {test_13_name} - Min: {test_13_min} {test_13_unit}, Max: {test_13_max} {test_13_unit} - Result: PASS ({resultado:.4f} {test_13_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    conexion_pcba()
                else:
                    label_test13.config(
                        text=f"Test 13: {test_13_name} - Min: {test_13_min} {test_13_unit}, Max: {test_13_max} {test_13_unit} - Result: FAIL ({resultado:.4f} {test_13_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_13)

        def test_14_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_14_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_14(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_14_name} {test_14_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro_psu)

            volatage_gtsoc = float(testspec_gtsoc("Voltage", "Voltage"))
            current_gtsoc = float(testspec_gtsoc("Current", "Current"))

            def configurar_multimetro_psu():
                DMM.write(b"CONF:VOLT:DC\n")
                time.sleep(0.1)
                PSU.write(b"VSET2:0\n")
                time.sleep(0.1)
                PSU.write(b"ISET2:0\n")
                time.sleep(0.1)
                PSU.write(f"VSET1:{volatage_gtsoc}\n".encode())
                time.sleep(0.1)
                PSU.write(f"ISET1:{current_gtsoc}\n".encode())
                time.sleep(0.1)
                PSU.write(b"OUT1\n")

                delay = int(testspec_gtsoc("Delay_Voltage", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"READ?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_14_min <= resultado <= test_14_max:
                    label_test14.config(
                        text=f"Test 14: {test_14_name} - Min: {test_14_min} {test_14_unit}, Max: {test_14_max} {test_14_unit} - Result: PASS ({resultado:.4f} {test_14_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    PSU.write(b"OUT0\n")
                    test_15_gtsoc()
                else:
                    label_test14.config(
                        text=f"Test 14: {test_14_name} - Min: {test_14_min} {test_14_unit}, Max: {test_14_max} {test_14_unit} - Result: FAIL ({resultado:.4f} {test_14_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_14)

        def test_15_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_15_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_15(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_15_name} {test_15_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro_psu)

            def configurar_multimetro_psu():
                PSU.write(b"OUT1\n")
                delay = int(testspec_gtsoc("Delay_Voltage", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"READ?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_15_min <= resultado <= test_15_max:
                    label_test15.config(
                        text=f"Test 15: {test_15_name} - Min: {test_15_min} {test_15_unit}, Max: {test_15_max} {test_15_unit} - Result: PASS ({resultado:.4f} {test_15_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    PSU.write(b"OUT0\n")
                    test_16_gtsoc()
                else:
                    label_test15.config(
                        text=f"Test 15: {test_15_name} - Min: {test_15_min} {test_15_unit}, Max: {test_15_max} {test_15_unit} - Result: FAIL ({resultado:.4f} {test_15_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_15)

        def test_16_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_16_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_16(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_16_name} {test_16_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro_psu)

            def configurar_multimetro_psu():
                PSU.write(b"OUT1\n")
                delay = int(testspec_gtsoc("Delay_Voltage", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"READ?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_16_min <= resultado <= test_16_max:
                    label_test16.config(
                        text=f"Test 16: {test_16_name} - Min: {test_16_min} {test_16_unit}, Max: {test_16_max} {test_16_unit} - Result: PASS ({resultado:.4f} {test_16_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    PSU.write(b"OUT0\n")
                    test_17_gtsoc()
                else:
                    label_test16.config(
                        text=f"Test 16: {test_16_name} - Min: {test_16_min} {test_16_unit}, Max: {test_16_max} {test_16_unit} - Result: FAIL ({resultado:.4f} {test_16_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_16)

        def test_17_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_17_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_17(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_17_name} {test_17_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro_psu)

            def configurar_multimetro_psu():
                PSU.write(b"OUT1\n")
                delay = int(testspec_gtsoc("Delay_Voltage", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"READ?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_17_min <= resultado <= test_17_max:
                    label_test17.config(
                        text=f"Test 17: {test_17_name} - Min: {test_17_min} {test_17_unit}, Max: {test_17_max} {test_17_unit} - Result: PASS ({resultado:.4f} {test_17_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    PSU.write(b"OUT0\n")
                    test_18_gtsoc()
                else:
                    label_test17.config(
                        text=f"Test 17: {test_17_name} - Min: {test_17_min} {test_17_unit}, Max: {test_17_max} {test_17_unit} - Result: FAIL ({resultado:.4f} {test_17_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_17)

        def test_18_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_18_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_18(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_18_name} {test_18_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro_psu)

            def configurar_multimetro_psu():
                PSU.write(b"OUT1\n")
                delay = int(testspec_gtsoc("Delay_Voltage", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"READ?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_18_min <= resultado <= test_18_max:
                    label_test18.config(
                        text=f"Test 18: {test_18_name} - Min: {test_18_min} {test_18_unit}, Max: {test_18_max} {test_18_unit} - Result: PASS ({resultado:.4f} {test_18_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    PSU.write(b"OUT0\n")
                    test_19_gtsoc()
                else:
                    label_test18.config(
                        text=f"Test 18: {test_18_name} - Min: {test_18_min} {test_18_unit}, Max: {test_18_max} {test_18_unit} - Result: FAIL ({resultado:.4f} {test_18_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_18)

        def test_19_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_19_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_19(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_19_name} {test_19_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro_psu)

            def configurar_multimetro_psu():
                PSU.write(b"OUT1\n")
                delay = int(testspec_gtsoc("Delay_Voltage", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"READ?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_19_min <= resultado <= test_19_max:
                    label_test19.config(
                        text=f"Test 19: {test_19_name} - Min: {test_19_min} {test_19_unit}, Max: {test_19_max} {test_19_unit} - Result: PASS ({resultado:.4f} {test_19_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    PSU.write(b"OUT0\n")
                    test_20_gtsoc()
                else:
                    label_test19.config(
                        text=f"Test 19: {test_19_name} - Min: {test_19_min} {test_19_unit}, Max: {test_19_max} {test_19_unit} - Result: FAIL ({resultado:.4f} {test_19_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_19)

        def test_20_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_20_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_20(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_20_name} {test_20_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro_psu)

            def configurar_multimetro_psu():
                PSU.write(b"OUT1\n")
                delay = int(testspec_gtsoc("Delay_Voltage", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"READ?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_20_min <= resultado <= test_20_max:
                    label_test20.config(
                        text=f"Test 20: {test_20_name} - Min: {test_20_min} {test_20_unit}, Max: {test_20_max} {test_20_unit} - Result: PASS ({resultado:.4f} {test_20_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    PSU.write(b"OUT0\n")
                    test_21_gtsoc()
                else:
                    label_test20.config(
                        text=f"Test 20: {test_20_name} - Min: {test_20_min} {test_20_unit}, Max: {test_20_max} {test_20_unit} - Result: FAIL ({resultado:.4f} {test_20_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_20)

        def test_21_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_21_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_21(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_21_name} {test_21_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro_psu)

            def configurar_multimetro_psu():
                PSU.write(b"OUT1\n")
                delay = int(testspec_gtsoc("Delay_Voltage", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"READ?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_21_min <= resultado <= test_21_max:
                    label_test21.config(
                        text=f"Test 21: {test_21_name} - Min: {test_21_min} {test_21_unit}, Max: {test_21_max} {test_21_unit} - Result: PASS ({resultado:.4f} {test_21_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    PSU.write(b"OUT0\n")
                    test_22_gtsoc()
                else:
                    label_test21.config(
                        text=f"Test 21: {test_21_name} - Min: {test_21_min} {test_21_unit}, Max: {test_21_max} {test_21_unit} - Result: FAIL ({resultado:.4f} {test_21_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_21)

        def test_22_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_22_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_22(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_22_name} {test_22_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro_psu)

            def configurar_multimetro_psu():
                PSU.write(b"OUT1\n")
                delay = int(testspec_gtsoc("Delay_Voltage", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"READ?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_22_min <= resultado <= test_22_max:
                    label_test22.config(
                        text=f"Test 22: {test_22_name} - Min: {test_22_min} {test_22_unit}, Max: {test_22_max} {test_22_unit} - Result: PASS ({resultado:.4f} {test_22_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    PSU.write(b"OUT0\n")
                    test_23_gtsoc()
                else:
                    label_test22.config(
                        text=f"Test 22: {test_22_name} - Min: {test_22_min} {test_22_unit}, Max: {test_22_max} {test_22_unit} - Result: FAIL ({resultado:.4f} {test_22_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_22)

        def test_23_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_23_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_23(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_23_name} {test_23_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro_psu)

            def configurar_multimetro_psu():
                PSU.write(b"OUT1\n")
                delay = int(testspec_gtsoc("Delay_Voltage", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"READ?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_23_min <= resultado <= test_23_max:
                    label_test23.config(
                        text=f"Test 23: {test_23_name} - Min: {test_23_min} {test_23_unit}, Max: {test_23_max} {test_23_unit} - Result: PASS ({resultado:.4f} {test_23_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    PSU.write(b"OUT0\n")
                    test_24_gtsoc()
                else:
                    label_test23.config(
                        text=f"Test 23: {test_23_name} - Min: {test_23_min} {test_23_unit}, Max: {test_23_max} {test_23_unit} - Result: FAIL ({resultado:.4f} {test_23_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_23)

        def test_24_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_24_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_24(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_24_name} {test_24_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro_psu)

            def configurar_multimetro_psu():
                PSU.write(b"OUT1\n")
                delay = int(testspec_gtsoc("Delay_Voltage", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"READ?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_24_min <= resultado <= test_24_max:
                    label_test24.config(
                        text=f"Test 24: {test_24_name} - Min: {test_24_min} {test_24_unit}, Max: {test_24_max} {test_24_unit} - Result: PASS ({resultado:.4f} {test_24_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    PSU.write(b"OUT0\n")
                    test_25_gtsoc()
                else:
                    label_test24.config(
                        text=f"Test 24: {test_24_name} - Min: {test_24_min} {test_24_unit}, Max: {test_24_max} {test_24_unit} - Result: FAIL ({resultado:.4f} {test_24_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_24)

        def test_25_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_25_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_25(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_25_name} {test_25_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro_psu)

            def configurar_multimetro_psu():
                PSU.write(b"OUT1\n")
                delay = int(testspec_gtsoc("Delay_Voltage", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"READ?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_25_min <= resultado <= test_25_max:
                    label_test25.config(
                        text=f"Test 25: {test_25_name} - Min: {test_25_min} {test_25_unit}, Max: {test_25_max} {test_25_unit} - Result: PASS ({resultado:.4f} {test_25_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    PSU.write(b"OUT0\n")
                    test_26_gtsoc()
                else:
                    label_test25.config(
                        text=f"Test 25: {test_25_name} - Min: {test_25_min} {test_25_unit}, Max: {test_25_max} {test_25_unit} - Result: FAIL ({resultado:.4f} {test_25_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_25)

        def test_26_gtsoc(event=None):
            Instrucciones.config(
                text=f"Coloque la punta roja del multímetro en el pin de {test_26_name} y la punta negra a tierra según la WI y presiona la barra espaciadora para iniciar", bg="SystemButtonFace", fg="blue")

            def inicio_test_26(event=None):
                Instrucciones.config(
                    text=f"Prueba: {test_26_name} {test_26_unit} - En proceso...",
                    bg="#FFEB9C", fg="#9C5700"
                )
                root.after(100, configurar_multimetro_psu)

            def configurar_multimetro_psu():
                PSU.write(b"OUT1\n")
                delay = int(testspec_gtsoc("Delay_Voltage", "delay"))
                root.after(delay, leer_resultado)

            def leer_resultado():
                DMM.write(b"READ?\n")

                respuesta = DMM.readline().decode().strip()
                resultado = float(respuesta) if respuesta else 0.0

                if test_26_min <= resultado <= test_26_max:
                    label_test26.config(
                        text=f"Test 26: {test_26_name} - Min: {test_26_min} {test_26_unit}, Max: {test_26_max} {test_26_unit} - Result: PASS ({resultado:.4f} {test_26_unit})",
                        bg="#C6EFCE", fg="green"
                    )
                    PSU.write(b"OUT0\n")
                    test_pass()
                else:
                    label_test26.config(
                        text=f"Test 26: {test_26_name} - Min: {test_26_min} {test_26_unit}, Max: {test_26_max} {test_26_unit} - Result: FAIL ({resultado:.4f} {test_26_unit})",
                        bg="#FFC7CE", fg="red"
                    )
                    test_fail()

            root.bind("<space>", inicio_test_26)

        def test_fail(event=None):
            """En caso de falla en algún test, esta función permitirá reiniciar la prueba"""
            label_resultado.config(text="TEST FAIL", bg="#FFC7CE", fg="red")
            cerrar_puertos()
            Instrucciones.config(
                text="Escanea el ID de la PCBA para comenzar", bg="SystemButtonFace", fg="blue")
            entry_id.config(state="normal")
            entry_id.delete(0, tk.END)
            root.unbind("<space>")
            guardar_resultados()
            entry_id.focus_set()

        def test_pass(event=None):
            """En caso de falla en algún test, esta función permitirá reiniciar la prueba"""
            label_resultado.config(text="TEST PASS", bg="#C6EFCE", fg="green")
            cerrar_puertos()
            Instrucciones.config(
                text="Escanea el ID de la PCBA para comenzar", bg="SystemButtonFace", fg="blue")
            entry_id.config(state="normal")
            entry_id.delete(0, tk.END)
            root.unbind("<space>")
            guardar_resultados()
            entry_id.focus_set()

        def guardar_resultados():
            ahora = datetime.now()

            fecha = ahora.strftime("%Y%m%d")
            hora = ahora.strftime("%H%M%S")

            pieza = label_id.cget("text")
            resultado = label_resultado.cget("text")
            num_operador = operador
            num_orden = orden

            nombre_archivo = f"{pieza}_{fecha}_{hora}.txt"

            carpeta = "Result_GTSOC"

            os.makedirs(carpeta, exist_ok=True)

            ruta = os.path.join(carpeta, nombre_archivo)

            with open(ruta, "w", encoding="utf-8") as archivo:

                archivo.write(f"ID: {pieza}\n")
                archivo.write(
                    f"Date: {ahora.strftime('%Y-%m-%d %H:%M:%S')}\n")
                archivo.write(
                    f"Employee: {num_operador}\n")
                archivo.write(
                    f"Work order: {num_orden}\n")
                archivo.write(
                    f"Result: {resultado}\n\n")

                archivo.write(label_test1.cget("text") + "\n")
                archivo.write(label_test2.cget("text") + "\n")
                archivo.write(label_test3.cget("text") + "\n")
                archivo.write(label_test4.cget("text") + "\n")
                archivo.write(label_test5.cget("text") + "\n")
                archivo.write(label_test6.cget("text") + "\n")
                archivo.write(label_test7.cget("text") + "\n")
                archivo.write(label_test8.cget("text") + "\n")
                archivo.write(label_test9.cget("text") + "\n")
                archivo.write(label_test10.cget("text") + "\n")
                archivo.write(label_test11.cget("text") + "\n")
                archivo.write(label_test12.cget("text") + "\n")
                archivo.write(label_test13.cget("text") + "\n")
                archivo.write(label_test14.cget("text") + "\n")
                archivo.write(label_test15.cget("text") + "\n")
                archivo.write(label_test16.cget("text") + "\n")
                archivo.write(label_test17.cget("text") + "\n")
                archivo.write(label_test18.cget("text") + "\n")
                archivo.write(label_test19.cget("text") + "\n")
                archivo.write(label_test20.cget("text") + "\n")
                archivo.write(label_test21.cget("text") + "\n")
                archivo.write(label_test22.cget("text") + "\n")
                archivo.write(label_test23.cget("text") + "\n")
                archivo.write(label_test24.cget("text") + "\n")
                archivo.write(label_test25.cget("text") + "\n")
                archivo.write(label_test26.cget("text") + "\n")

        entry_id.bind("<Return>", validar_id)


if __name__ == "__main__":
    root = tk.Tk()
    app = VentanaLogin(root)
    # root.protocol("WM_DELETE_WINDOW", cerrar_puertos)
    root.mainloop()
