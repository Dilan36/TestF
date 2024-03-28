import tkinter as tk
from tkinter import ttk
import pyperclip
from selenium import webdriver
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import tkinter.simpledialog as sd
import datetime
import re
import openpyxl
from operator import itemgetter


# Declare a global variable to store the data
global saved_data
saved_data = {}
#saved_data = [['Falabella Empresas', '338', 'LU.MA.MI.JU.VI.-.-', '8#18'], ['BOGOTA ICBS(ATH)', '232', 'LU.MA.MI.JU.VI.SA.DO', '4#23'], ['AV.VILLAS ICBS(ATH)', '215', 'LU.MA.MI.JU.VI.SA.DO', '4#23'], ['OCCIDENTE ICBS(ATH)', '226', 'LU.MA.MI.JU.VI.-.-', '4#23'], ['POPULAR ICBS(ATH)', '154', 'LU.MA.MI.JU.VI.-.-', '6#20'], ['POPULAR ICBS(BPOP)', '307', 'LU.MA.MI.JU.VI.-.-', '6#20'], ['TRANSFIYA', '234', 'LU.MA.MI.JU.VI.-.-', '7#17'], ['SUCURSAL NY', '154', 'LU.MA.MI.JU.VI.-.-', '5.75#19.75'], ['SUCURSAL NY', '344', 'LU.MA.MI.JU.VI.-.-', '5.75#19.75'], ['Transferencia internacionales PER', '361', 'LU.MA.MI.JU.VI.SA.-', '9#18'], ['Transferencia cuentas propias PER', '362', 'LU.MA.MI.JU.VI.SA.-', '9#18'], ['Transferencia cuentas propias PER 2', '369', 'LU.MA.MI.JU.VI.SA.-', '9#18'], ['Transferencia cuentas propias PER 2', '369', 'LU.MA.MI.JU.VI.SA.-', '9#18'], ['DaviPlata Desarrollo (8 am - 7 pm)', '529', 'LU.MA.MI.JU.VI.-.-', '8#19']]
saved_data = [['Falabella Empresas', '338', 'LU.MA.MI.JU.VI.-.-', '8#18'], ['BOGOTA ICBS(ATH)', '232', 'LU.MA.MI.JU.VI.SA.DO', '4#23'], ['AV.VILLAS ICBS(ATH)', '215', 'LU.MA.MI.JU.VI.SA.DO', '4#23'], ['OCCIDENTE ICBS(ATH)', '226', 'LU.MA.MI.JU.VI.-.-', '4#23'], ['POPULAR ICBS(ATH)', '154', 'LU.MA.MI.JU.VI.-.-', '6#20'], ['POPULAR ICBS(BPOP)', '307', 'LU.MA.MI.JU.VI.-.-', '6#20'], ['TRANSFIYA', '234', 'LU.MA.MI.JU.VI.-.-', '7#17'], ['SUCURSAL NY', '154', 'LU.MA.MI.JU.VI.-.-', '5.75#19.75'], ['SUCURSAL NY', '344', 'LU.MA.MI.JU.VI.-.-', '5.75#19.75'], ['Transferencia internacionales PER', '361', 'LU.MA.MI.JU.VI.SA.-', '9#18'], ['Transferencia cuentas propias PER', '362', 'LU.MA.MI.JU.VI.SA.-', '9#18'], ['Transferencia cuentas propias PER 2', '369', 'LU.MA.MI.JU.VI.SA.-', '9#18'], ['Transferencia cuentas propias PER 2', '369', 'LU.MA.MI.JU.VI.SA.-', '9#18'], ['DaviPlata Desarrollo (8 am - 7 pm)', '529', 'LU.MA.MI.JU.VI.-.-', '8#19'], [' Mi CLARO B2B2C - RECARGAS PASATIEMPO SMS 1', '435', ' LU.MA.MI.JU.VI.-.-', '5#9'], [' Mi CLARO B2B2C - RECARGAS PASATIEMPO SMS 1', '435', ' LU.MA.MI.JU.VI.-.-', '14#21'], [' Mi CLARO B2B2C - RECARGAS PASATIEMPO MAIL 1', '436', ' LU.MA.MI.JU.VI.-.-', '5#9'], [' Mi CLARO B2B2C - RECARGAS PASATIEMPO MAIL 1', '436', ' LU.MA.MI.JU.VI.-.-', '14#21'], [' Mi CLARO B2B2C - RECARGAS PASATIEMPO SMS 2', '437', ' -.-.-.-.-.SA.DO', '7#19'], [' Mi CLARO B2B2C - RECARGAS PASATIEMPO MAIL 2', '438', ' -.-.-.-.-.SA.DO', '7#19']]


def open_menu():
    global saved_data
    # Crear una ventana secundaria para el menú
    menu_window = tk.Toplevel(root)
    menu_window.title("Excepciones:")

    # Crear etiqueta y campo de texto para excepciones
    ttk.Label(menu_window, text="Ingrese Excepciones (separadas por comas):").grid(row=0, column=0, sticky="w", padx=10, pady=5)

    # Usar tk.Text en lugar de scrolledtext.ScrolledText
    excepciones_text = tk.Text(menu_window, wrap=tk.WORD, width=100, height=3)
    excepciones_text.grid(row=1, columnspan=2, column=0, padx=10, pady=5, sticky="w")
    
    # Crear un widget Treeview para mostrar los datos ingresados
    data_tree = ttk.Treeview(menu_window, columns=("Campaña", "Número Campaña", "Días de Operación", "Horario de Operación"), show="headings")
    data_tree.heading("Campaña", text="Campaña")
    data_tree.heading("Número Campaña", text="Número Campaña")
    data_tree.heading("Días de Operación", text="Días de Operación")
    data_tree.heading("Horario de Operación", text="Horario de Operación")
    data_tree.grid(row=3, column=0, padx=10, pady=5, columnspan=2, sticky="w")

    # Función para obtener y mostrar los datos ingresados
    def show_data():
        # Borrar el contenido actual en la tabla
        data_tree.delete(*data_tree.get_children())
        for item in saved_data:
            data_tree.insert("", "end", values=item)

    # Botón para mostrar los datos
    ttk.Button(menu_window, text="Mostrar Datos", command=show_data).grid(row=4, column=1, padx=10, pady=5, sticky="w")
    
    # Función para obtener y guardar los datos ingresados
    def save_data():
        global saved_data
        excepciones = excepciones_text.get("1.0", "end-1c").split("//")
        # Puedes adaptar esto según tus necesidades específicas
        # Aquí, estoy guardando las excepciones y el texto adicional en un diccionario
        data = [elemento.split(',') for elemento in excepciones]
        
        saved_data= data

    # Botón para guardar los datos
    ttk.Button(menu_window, text="Guardar", command=save_data).grid(row=2, column=1, padx=10, pady=5, sticky="w")

    # Nota debajo del campo de texto
    ttk.Label(menu_window, text="NOTA: Tener por defecto siempre lo siguiente:").grid(row=4, column=0, sticky="w", padx=10)

    # Nota de texto
    nota_text = "Falabella Empresas,338,LU.MA.MI.JU.VI.-.-,8#18//BOGOTA ICBS(ATH),232,LU.MA.MI.JU.VI.SA.DO,4#23//AV.VILLAS ICBS(ATH),215,LU.MA.MI.JU.VI.SA.DO,4#23//OCCIDENTE ICBS(ATH),226,LU.MA.MI.JU.VI.-.-,4#23//POPULAR ICBS(ATH),154,LU.MA.MI.JU.VI.-.-,6#20//POPULAR ICBS(BPOP),307,LU.MA.MI.JU.VI.-.-,6#20//TRANSFIYA,234,LU.MA.MI.JU.VI.-.-,7#17//SUCURSAL NY,154,LU.MA.MI.JU.VI.-.-,5.75#19.75//SUCURSAL NY,344,LU.MA.MI.JU.VI.-.-,5.75#19.75//Transferencia internacionales PER,361,LU.MA.MI.JU.VI.SA.-,9#18//Transferencia cuentas propias PER,362,LU.MA.MI.JU.VI.SA.-,9#18//Transferencia cuentas propias PER 2,369,LU.MA.MI.JU.VI.SA.-,9#18//Transferencia cuentas propias PER 2,369,LU.MA.MI.JU.VI.SA.-,9#18//DaviPlata Desarrollo (8 am - 7 pm),529,LU.MA.MI.JU.VI.-.-,8#19//Mi CLARO B2B2C - RECARGAS PASATIEMPO SMS 1,435,LU.MA.MI.JU.VI.-.-,5#9// Mi CLARO B2B2C - RECARGAS PASATIEMPO SMS 1,435,LU.MA.MI.JU.VI.-.-,14#21//Mi CLARO B2B2C - RECARGAS PASATIEMPO MAIL 1,436,LU.MA.MI.JU.VI.-.-,5#9//Mi CLARO B2B2C - RECARGAS PASATIEMPO MAIL 1,436,LU.MA.MI.JU.VI.-.-,14#21// Mi CLARO B2B2C - RECARGAS PASATIEMPO SMS 2,437,-.-.-.-.-.SA.DO,7#19//Mi CLARO B2B2C - RECARGAS PASATIEMPO MAIL 2,438,-.-.-.-.-.SA.DO,7#19"
    nota_label = ttk.Label(menu_window, text=nota_text, justify="left", anchor="w")
    nota_label.config(wraplength=810)  # Ajusta 300 según tus necesidades
    nota_label.grid(row=5,column=0, columnspan=2, sticky="w", padx=10, pady=5)

    # Función para copiar el texto del Label al portapapeles
    def copy_label_text():
        texto = nota_label.cget("text")
        root.clipboard_clear()
        root.clipboard_append(texto)
        root.update()

    # Botón para copiar el texto del Label
    ttk.Button(menu_window, text="Copia texto por defecto", command=copy_label_text).grid(row=6, column=0, padx=10, pady=5, sticky="w")


def start_process():
    lim_menor = float(lim_menor_entry.get())
    lim_mayor = float(lim_mayor_entry.get())
    datiris=saved_data
    # Función para copiar la información de las filas seleccionadas
    def copy_selected():
        selected_items = root.result_table.selection()
        if selected_items:
            selected_data = []
            for item in selected_items:
                values = root.result_table.item(item, "values")
                selected_data.append("\t".join(values))  # Puedes personalizar el separador si es necesario
            copied_data = "\n".join(selected_data)
            pyperclip.copy(copied_data)
            print("Información copiada al portapapeles.")
        else:
            print("Selecciona al menos una fila primero.")
                
    def copy_all():
        all_data = []
        for item in root.result_table.get_children():
            values = root.result_table.item(item, "values")
            all_data.append("\t".join(values))  # Puedes personalizar el separador si es necesario
        copied_data = "\n".join(all_data)
        pyperclip.copy(copied_data)
        print("Toda la tabla copiada al portapapeles.")

    def destroy_table():
        if hasattr(root, 'result_table') and root.result_table.winfo_exists():
            root.result_table.destroy()
            
    
    def veridia(texto):
        # Obten el día actual
        dia_actual = datetime.datetime.now().strftime("%A")

        # Define un diccionario para mapear abreviaciones de días a días completos
        mapeo_dias = {
            "LU": "Monday",
            "MA": "Tuesday",
            "MI": "Wednesday",
            "JU": "Thursday",
            "VI": "Friday",
            "SA": "Saturday",
            "DO": "Sunday"
        }

        # Divide el texto en días usando coma como separador
        dias = texto.split(',')

        # Verifica si el día actual coincide con alguno de los días del texto
        for dia in dias:
            if mapeo_dias.get(dia) == dia_actual:
                return True

        return False

    def evalhoras(texto):
        # Extrae la hora de inicio y la hora final del texto
        hora_inicio_texto, hora_final_texto = map(float, texto.split('#'))  # Convierte las cifras en el texto a enteros

        # Obtiene la hora actual, incluyendo minutos y segundos
        hora_actual = datetime.datetime.now().hour + datetime.datetime.now().minute / 60 + datetime.datetime.now().second / 3600

        # Verifica si la hora actual está dentro del rango especificado
        if hora_inicio_texto <= hora_actual <= hora_final_texto:
            return True
        else:
            return False

    # Función para ordenar la tabla por tiempo caído de mayor a menor
    def sort_table():
        items = root.result_table.get_children()  # Obtiene las filas de la tabla
        data = []  # Almacena los datos de las filas

        for item in items:
            values = root.result_table.item(item, "values")
            data.append(values)

        # Ordena los datos en función del tiempo caído (columna 6, índice 5 en Python)
        sorted_data = sorted(data, key=lambda x: float(x[5]), reverse=True)

        # Borra todas las filas actuales en la tabla
        for item in items:
            root.result_table.delete(item)

        # Inserta las filas ordenadas en la tabla
        for row in sorted_data:
            root.result_table.insert("", "end", values=row)
    
    destroy_table()
        
    # Crear una tabla para mostrar los resultados
    root.result_table = ttk.Treeview(root, columns=("Campaña", "IMEI", "Clave", "Estado", "Operador", "Valor", "Tiempo Caido (horas)"), selectmode="extended")  # "extended" permite la selección múltiple
    
    root.result_table.heading("#1", text="Campaña")
    root.result_table.heading("#2", text="IMEI")
    root.result_table.heading("#3", text="Clave")
    root.result_table.heading("#4", text="Estado")
    root.result_table.heading("#5", text="Operador")
    root.result_table.heading("#6", text="Valor")
    root.result_table.heading("#7", text="Tiempo Caido (horas)")
    
    # Establecer el ancho de cada columna
    root.result_table.column("#1", width=300)  # Ancho de la primera columna
    root.result_table.column("#2", width=100)  # Ancho de la segunda columna
    root.result_table.column("#3", width=80)  # Ancho de la tercera columna
    root.result_table.column("#4", width=150)  # Ancho de la cuarta columna
    root.result_table.column("#5", width=100)  # Ancho de la quinta columna
    root.result_table.column("#6", width=80)  # Ancho de la quinta columna
    root.result_table.column("#7", width=80)  # Ancho de la quinta columna
    
    root.result_table.grid(row=5, column=0, columnspan=7, rowspan=7)  # Colocar la tabla en la interfaz
    
        
    #fila vacia
    empty_row = tk.Frame(root, height=10)
    empty_row.grid(row=12)    
    
    copy_button = ttk.Button(root, text="Copiar Selección", command=copy_selected)
    copy_button.grid(row=13, column=1)
       
    copy_all_button = ttk.Button(root, text="Copiar Toda la Tabla", command=copy_all)
    copy_all_button.grid(row=13, column=4)
           
    # Agrega un botón para ordenar las filas
    sort_button = ttk.Button(root, text="Ordenar por Tiempo Caido (Mayor a Menor)", command=sort_table)
    sort_button.grid(row=13, column=2)
    
    # Resto de tu código aquí
    opts = Options()
    opts.add_argument("--incognito")
    #...................ocultar ventana..........................................................
    #opts.add_argument("--headless")
    #opts.add_argument("--disable-gpu")
    #opts.add_argument("--window-size=1920x1080")  # Puedes ajustar el tamaño de la ventana
    #--------------------------------------------------------------------------------------------
    opts.add_argument("user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36")
    driver = webdriver.Chrome(service = Service(ChromeDriverManager().install()), options=opts)
    driver.get('https://mantenedor.movizzon.com/appMonitors')

    user = "felipe.zapata@movizzon.com"
    password = "12345678"
    cant_ant=0
    
    # Me doy cuenta que la pagina carga el formulario dinamicamente luego de que la carga incial ha sido completada
    # Por eso tengo que esperar que aparezca 
    #input_user = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div/div[2]/form/div[1]/input')))
    #input_user.send_keys(user)

    # Obtengo el boton next y lo presiono para poder poner la pass
    #next_button = driver.find_element(By.XPATH, '//div[@class="css-901oao r-1awozwy r-6koalj r-18u37iz r-16y2uox r-37j5jr r-a023e6 r-b88u0q r-1777fci r-rjixqe r-bcqeeo r-q4m81j r-qvutc0"]')
    #next_button.click()

    # Obtengo los inputs de usuario (linea 28) y password
    #input_pass = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div/div[2]/form/div[2]/input')))
    #input_pass.send_keys(password)

    # Obtengo el boton de login
    #login_button = driver.find_element(By.XPATH, '/html/body/div/div/div[2]/form/div[4]/button')
    # Le doy click
    #login_button.click()

    campaña = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div/main/div/div/div/div/div/h2')))
    imeis = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div/main/div/div/div/div/div/table/tbody/tr/td[1]')))
    operadores = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div/main/div/div/div/div/div/table/tbody/tr/td[2]')))
    valores = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div/main/div/div/div/div/div/table/tbody/tr/td[5]')))
    fallahora = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div/main/div/div/div/div/div/table/tbody/tr/td[7]/b')))
    actividad = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div/main/div/div/div/div/div/table/tbody/tr/td[3]')))
    # Función para convertir días, horas y minutos a horas
    def convert_to_hours(text):
        # Dividir el texto en días, horas y minutos
        parts = text.split(" - ")
        #dias, horas, minutos = map(int, [p.strip() for p in parts[0].split() if p.strip().isdigit()])
        dias= int(parts[0].replace('D', ''))
        horas= int(parts[1].replace('H', ''))
        minutos= int(parts[2].replace('M', ''))
        horas_totales = dias * 24 + horas + minutos/ 60
        return horas_totales

    # Procesar los textos, realizar la conversión y suma
    horas_totales = []

    for valor in valores:
        texto = valor.text
        total_horas = convert_to_hours(texto)
        horas_totales.append(total_horas)
    
    # Imprime el texto de los elementos
    for x in range(len(campaña)):
    
        xpath = f"/html/body/div/main/div/div/div/div/div[{x+2}]/table/tbody/tr/td[5]"
        imeis_campaña = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
        cantidadimeis=len(imeis_campaña)
        accion_realizada= False
            
        for i in range(cant_ant, cantidadimeis + cant_ant):
            while not accion_realizada:
                cant_ant = len(imeis_campaña)+ cant_ant
                accion_realizada = True  # Establecemos la variable a True para que no se repita la acción

            campañae= campaña[x].text
            numcampaña= campañae.split()[-1]
            imei = imeis[i].text
            operador = operadores[i].text
            horas = horas_totales[i]
            valor = valores[i].text
            activ = actividad[i].text
            falla = fallahora[i].text
            falla= float(falla.replace('%', ''))
            
            
            if  (lim_menor <= horas <= lim_mayor) or (horas==2476.65):
                xx2=0
                xx1=0
                if saved_data[0][0]!='':
                    for elemento in range(len(saved_data)):
                        diasop = saved_data[elemento][2]
                        horaseva = saved_data[elemento][3]
                        resultado = veridia(diasop)
                        resultadohoras = evalhoras(horaseva)
            
                        if saved_data[elemento][1] == numcampaña and resultado== True and resultadohoras== True:
                            #print(f"CAMPANA:{campañae}, IMEI: {imei}, Operador: {operador}, Valor: {valor}, tiempo caido (horas): {horas:.2f} ") 
                            #root.result_table.insert("", "end", values=(campañae, imei, activ ,operador, valor, f"{horas:.2f}"))
                            xx1=1
                        if saved_data[elemento][1] != numcampaña:
                            xx2=xx2+1
                        
                    if (xx1==0 and xx2==len(saved_data)) or (xx1==1 and xx2!=len(saved_data)):   
                        root.result_table.insert("", "end", values=(campañae, imei, activ ,operador, valor, f"{horas:.2f}"))
                elif saved_data[0][0]=='':
                    root.result_table.insert("", "end", values=(campañae, imei, activ ,operador, valor, f"{horas:.2f}"))
                    
                    

def start_process2():
    porc_menor = float(Porc_menor_entry.get())
    porc_mayor = float(Porc_mayor_entry.get())
                
    # Función para copiar la información de las filas seleccionadas
    def copy_selected2():
        selected_items = root.result_table.selection()
        if selected_items:
            selected_data = []
            for item in selected_items:
                values = root.result_table.item(item, "values")
                selected_data.append("\t".join(values))  # Puedes personalizar el separador si es necesario
            copied_data = "\n".join(selected_data)
            pyperclip.copy(copied_data)
            print("Información copiada al portapapeles.")
        else:
            print("Selecciona al menos una fila primero.")
                
    def copy_all2():
        all_data = []
        for item in root.result_table.get_children():
            values = root.result_table.item(item, "values")
            all_data.append("\t".join(values))  # Puedes personalizar el separador si es necesario
        copied_data = "\n".join(all_data)
        pyperclip.copy(copied_data)
        print("Toda la tabla copiada al portapapeles.")

    def destroy_table():
        if hasattr(root, 'result_table') and root.result_table.winfo_exists():
            root.result_table.destroy()
            
    # Establecer el tamaño de la ventana principal
    #root.geometry("1000x600")  # Ancho x Alto en píxeles

    destroy_table()
    
    # Crear una tabla para mostrar los resultados
    root.result_table = ttk.Treeview(root, columns=("Campaña", "IMEI", "Estado", "Operador", "Valor", "Tiempo Caido (horas)"), selectmode="extended")  # "extended" permite la selección múltiple
    
    
    # Crear una segunda tabla para mostrar la información adicional
    root.result_table = ttk.Treeview(root, columns=("Campaña", "IMEI", "Estado", "Operador", "Valor", "Tiempo Caido (horas)", "Porcentaje de Falla por Hora"))
    root.result_table.heading("#1", text="Campaña")
    root.result_table.heading("#2", text="IMEI")
    root.result_table.heading("#3", text="Estado")
    root.result_table.heading("#4", text="Operador")
    root.result_table.heading("#5", text="Valor")
    root.result_table.heading("#6", text="Tiempo Caido (horas)")
    root.result_table.heading("#7", text="Porcentaje de Falla por Hora")
    # Establecer el ancho de cada columna
    root.result_table.column("#1", width=300)  # Ancho de la primera columna
    root.result_table.column("#2", width=100)  # Ancho de la segunda columna
    root.result_table.column("#3", width=80)  # Ancho de la tercera columna
    root.result_table.column("#4", width=150)  # Ancho de la cuarta columna
    root.result_table.column("#5", width=100)  # Ancho de la quinta columna
    root.result_table.column("#6", width=80)  # Ancho de la sexta columna
    root.result_table.column("#7", width=200)  # Ancho de la sexta columna
    root.result_table.grid(row=5, column=0, columnspan=6, rowspan=6)  # Colocar la segunda tabla en la interfaz
    
    #fila vacia
    empty_row = tk.Frame(root, height=10)
    empty_row.grid(row=12)
    
    copy_button = ttk.Button(root, text="Copiar Selección", command=copy_selected2)
    copy_button.grid(row=13, column=1)
       
    copy_all_button = ttk.Button(root, text="Copiar Toda la Tabla", command=copy_all2)
    copy_all_button.grid(row=13, column=4)
       
    # Resto de tu código aquí
    opts = Options()
    opts.add_argument("--incognito")
    #...................ocultar ventana..........................................................
    opts.add_argument("--headless")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1920x1080")  # Puedes ajustar el tamaño de la ventana
    #--------------------------------------------------------------------------------------------
    opts.add_argument("user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36")

    driver = webdriver.Chrome(service = Service(ChromeDriverManager().install()), options=opts)
    driver.get('https://mantenedor.movizzon.com/appMonitors')
    
    user = "felipe.zapata@movizzon.com"
    password = "12345678"
    cant_ant=0
    
    # Me doy cuenta que la pagina carga el formulario dinamicamente luego de que la carga incial ha sido completada
    # Por eso tengo que esperar que aparezca 
    #input_user = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div/div[2]/form/div[1]/input')))
    #input_user.send_keys(user)

    # Obtengo el boton next y lo presiono para poder poner la pass
    #next_button = driver.find_element(By.XPATH, '//div[@class="css-901oao r-1awozwy r-6koalj r-18u37iz r-16y2uox r-37j5jr r-a023e6 r-b88u0q r-1777fci r-rjixqe r-bcqeeo r-q4m81j r-qvutc0"]')
    #next_button.click()

    # Obtengo los inputs de usuario (linea 28) y password
    #input_pass = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div/div[2]/form/div[2]/input')))
    #input_pass.send_keys(password)

    # Obtengo el boton de login
    #login_button = driver.find_element(By.XPATH, '/html/body/div/div/div[2]/form/div[4]/button')
    # Le doy click
    #login_button.click()

    campaña = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div/main/div/div/div/div/div/h2')))
    imeis = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div/main/div/div/div/div/div/table/tbody/tr/td[1]')))
    operadores = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div/main/div/div/div/div/div/table/tbody/tr/td[2]')))
    valores = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div/main/div/div/div/div/div/table/tbody/tr/td[5]')))
    fallahora = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div/main/div/div/div/div/div/table/tbody/tr/td[7]/b')))
    actividad = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '/html/body/div/main/div/div/div/div/div/table/tbody/tr/td[3]')))
    
    # Función para convertir días, horas y minutos a horas
    def convert_to_hours(text):
        # Dividir el texto en días, horas y minutos
        parts = text.split(" - ")
        #dias, horas, minutos = map(int, [p.strip() for p in parts[0].split() if p.strip().isdigit()])
        dias= int(parts[0].replace('D', ''))
        horas= int(parts[1].replace('H', ''))
        minutos= int(parts[2].replace('M', ''))
        horas_totales = dias * 24 + horas + minutos/ 60
        return horas_totales

    # Procesar los textos, realizar la conversión y suma
    horas_totales = []
    
    for valor in valores:
        texto = valor.text
        total_horas = convert_to_hours(texto)
        horas_totales.append(total_horas)

    # Imprime el texto de los elementos
    for x in range(len(campaña)):
    
        xpath = f"/html/body/div/main/div/div/div/div/div[{x+2}]/table/tbody/tr/td[5]"
        imeis_campaña = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, xpath)))
        cantidadimeis=len(imeis_campaña)
        accion_realizada= False
            
        for i in range(cant_ant, cantidadimeis + cant_ant):
            while not accion_realizada:
                cant_ant = len(imeis_campaña)+ cant_ant
                accion_realizada = True  # Establecemos la variable a True para que no se repita la acción

            campañae= campaña[x].text
            imei = imeis[i].text
            operador = operadores[i].text
            horas = horas_totales[i]
            valor = valores[i].text
            activ = actividad[i].text
            falla = fallahora[i].text
            falla= float(falla.replace('%', ''))

            if porc_menor <= falla <= porc_mayor:
                #print(f"CAMPANA:{campañae}, IMEI: {imei}, Operador: {operador}, Valor: {valor}, tiempo caido (horas): {horas:.2f}, porcentaje falla por hora : {falla}") 
                root.result_table.insert("", "end", values=(campañae, imei, activ, operador, valor, f"{horas:.2f}", falla))                  

# Función para buscar el número ingresado en la columna C del archivo de Excel
def buscar_numero(event):
    try:
        numero_buscar = entry_numero.get()  # Obtener el número ingresado en el Entry como una cadena
        archivo_excel = openpyxl.load_workbook('ACCESOS.xlsx')  # Reemplaza 'archivo.xlsx' con el nombre de tu archivo Excel
        hoja = archivo_excel.active
        
        # Limpiar la tabla antes de mostrar nuevos resultados
        for row in resultados_treeview.get_children():
            resultados_treeview.delete(row)
        
        for fila in hoja.iter_rows(min_row=2, values_only=True):  # Comienza desde la segunda fila (suponiendo que la primera fila son encabezados)
            valores_celda = fila[2]  # Columna C (índice 2)
            if valores_celda is not None and numero_buscar in str(int(valores_celda)):
                resultados_treeview.insert("", "end", values=(fila[0],fila[1],int(fila[2]),fila[3],fila[4],int(fila[5])))
    except Exception as e:
        print(f"Ocurrió un error: {str(e)}")


# Crear la ventana principal
root = tk.Tk()
root.title("Proceso Automatizado")

# Diccionario para almacenar la configuración de cada campaña
campañas_config = {}

# Agregar un menú con la opción "Configurar Campaña"
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)
menu_campaña = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Configuración", menu=menu_campaña)
menu_campaña.add_command(label="excepciones", command=open_menu)

# Utiliza columnconfigure y rowconfigure para ajustar la geometría de la ventana principal
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=1)
root.columnconfigure(2, weight=1)
root.columnconfigure(3, weight=1)
root.columnconfigure(4, weight=1)
root.columnconfigure(5, weight=1)

root.rowconfigure(0, weight=1)
root.rowconfigure(1, weight=1)
root.rowconfigure(2, weight=1)
root.rowconfigure(3, weight=1)
root.rowconfigure(4, weight=1)
root.rowconfigure(5, weight=1)
root.rowconfigure(6, weight=1)
root.rowconfigure(7, weight=1)
root.rowconfigure(8, weight=1)
root.rowconfigure(9, weight=1)


#fila vacia
empty_row = tk.Frame(root, height=10)
empty_row.grid(row=0)

# Crear etiquetas y campos de entrada para lim_menor y lim_mayor
lim_menor_label = ttk.Label(root, text="Intervalo limite Menor de tiempo:")
lim_menor_label.grid(row=1, column=1)
lim_menor_entry = ttk.Entry(root)
lim_menor_entry.grid(row=1, column=2)

lim_mayor_label = ttk.Label(root, text="Intervalo Limite Mayor de tiempo:")
lim_mayor_label.grid(row=2, column=1)
lim_mayor_entry = ttk.Entry(root)
lim_mayor_entry.grid(row=2, column=2)

# Crear etiquetas y campos de entrada para lim_menor y lim_mayor
Porc_menor_label = ttk.Label(root, text="Intervalo Limite Porcentaje de falla Menor:")
Porc_menor_label.grid(row=1, column=4)
Porc_menor_entry = ttk.Entry(root)
Porc_menor_entry.grid(row=1, column=5)

Porc_mayor_label = ttk.Label(root, text="Intervalo Limite Porcentaje de falla Mayor:")
Porc_mayor_label.grid(row=2, column=4)
Porc_mayor_entry = ttk.Entry(root)
Porc_mayor_entry.grid(row=2, column=5)

#fila vacia
empty_row = tk.Frame(root, height=10)
empty_row.grid(row=3)

# Botón para iniciar el proceso
start_button = ttk.Button(root, text="Ver procesos caidos", command=start_process)
start_button.grid(row=4, column=2, columnspan=1, sticky="nsew", pady=10, padx=10)

# Botón para iniciar el proceso
start_button = ttk.Button(root, text="Ver procesos con falla", command=start_process2)
start_button.grid(row=4, column=5, columnspan=1, sticky="nsew",pady=10,padx=10)

##----------------------------------------------------------------------------------------------
# Crear un Entry para ingresar el número
label_numero = tk.Label(root, text="Busqueda de acceso remoto ingresa el imei:")
label_numero.grid(row=15, column=3, columnspan=1, sticky="nsew", padx=0)
entry_numero = tk.Entry(root)
entry_numero.grid(row=16, column=3, columnspan=1, sticky="nsew", padx=0)

# Vincular la función buscar_numero al evento KeyRelease del Entry
entry_numero.bind('<KeyRelease>', buscar_numero)
#fila vacia
empty_row = tk.Frame(root, height=10)
empty_row.grid(row=17)
# Crear la tabla (Treeview) para mostrar los resultados

    
# Establecer el ancho de cada columna

    
##----------------------------------------------------------------------------------------------

# Iniciar la interfaz gráfica
root.mainloop()
