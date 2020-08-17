from time import sleep
import os
import calendar
from enum import Enum
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
import pandas as pd


USUARIOS = []
EXCEL = "/home/panchi/Public/Claves-IIBB.xls"
RESULTADOS = "resultados.txt"
ERRORES = 0
INTENTOS = 3
MES = 7


class AuthError(Exception):
    pass

# Me paro donde tengo el codigo
os.chdir(os.path.dirname(os.path.abspath(__file__)))
if not os.path.exists('Temporales'):
    os.makedirs('Temporales')
CARPETA_DESCARGAS = os.path.join(os.getcwd(), 'Temporales')


class Usuario:
    def __init__(self, nombre, apellido, cuit, clave):
        self.nombre = nombre
        self.apellido = apellido
        self.cuit = str(cuit)
        self.clave = clave

    def __repr__(self):
        return f"<Usuario: {self.nombre} {self.apellido} CUIT: {self.cuit} Clave: {self.clave}>"

    def __str__(self):
        return f"<Usuario: {self.nombre} {self.apellido} CUIT: {self.cuit} Clave: {self.clave}>"

class Comprobantes(Enum):
    EMITIDOS = "Emitidos"
    RECIBIDOS = "Recibidos"

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    
def cprint(color, msg):
    print(f"{color}{msg}{bcolors.ENDC}")
    
def separador():
    print("\n=^..^=   =^..^=   =^..^=    =^..^=    =^..^=    =^..^=    =^..^=\n")


def afip_bot(driver, usuario, meses, tipo_comprobante):
    
    # Voy a la pagina del Afip
    print("Abro la pagina de Afip")
    driver.get("https://auth.afip.gob.ar/contribuyente_/login.xhtml")
    sleep(2)
    
    # Cargar CUIT/CUIL
    driver.find_element_by_xpath(
        "//input[@name=\"F1:username\"]").send_keys(usuario.cuit)
    driver.find_element_by_xpath("//input[@type=\"submit\"]").click()
    sleep(3)
    
    # Cargar clave
    driver.find_element_by_xpath("//input[@name=\"F1:password\"]")\
        .send_keys(usuario.clave)
    driver.find_element_by_xpath("//input[@type=\"submit\"]").click()
    sleep(1)
    try:
        mensaje = driver.find_element_by_xpath("//span[@id=\"F1:msg\"]")
    except:
        pass
    else:
        if "Clave o usuario incorrecto" in mensaje.text:
            raise AuthError("Clave o usuario incorrecto")
    
    try:
        mensaje = driver.find_element_by_xpath("//*[@id=\"contenido\"]/div/div/div[2]/div/div/h4")
    except:
        pass
    else:
        if "CAMBIAR CLAVE FISCAL" in mensaje.text:
            raise AuthError("Pide cambiar la clave fiscal")
    sleep(10)
    
    # Mis comprobantes
    try:
        # Pag vieja
        elems = driver.find_elements_by_css_selector(".azul.bold.font-size-14")
        elems[0].text
        for elem in elems:
            if "mis comprobantes" in elem.text.lower():
                elem.click()
                break
    except:
        # Pag nueva
        elems = driver.find_elements_by_tag_name("h4")
        for elem in elems:
            if "mis comprobantes" in elem.text.lower():
                elem.click()
                break
    sleep(5)

    # Cambio de pestaña
    driver.switch_to.window(driver.window_handles[1])
    sleep(2)
    
    # Tipo de Comprobantes
    if tipo_comprobante == Comprobantes.EMITIDOS:
        driver.find_element_by_xpath("//a[@id=\"btnEmitidos\"]").click()
    elif tipo_comprobante == Comprobantes.RECIBIDOS:
        driver.find_element_by_xpath("//a[@id=\"btnRecibidos\"]").click()
    else:
        raise ValueError("Tipo de comprobante desconocido") 
    
    for mes in meses:
        sleep(2)
        # Cargar Fechas
        fechas = crear_rango_fechas(mes)
        print(f"Cargo el rango de fehas: {fechas}")
        driver.find_element_by_xpath("//input[@id=\"fechaEmision\"]").clear()
        driver.find_element_by_xpath("//input[@id=\"fechaEmision\"]").send_keys(fechas)
        driver.find_element_by_xpath("//button[@type=\"submit\"]").click()
        sleep(5)
        
        # Bajar como CSV
        try:
            boton = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((
                    By.XPATH, "/html/body/main/div/section/div[1]/div/div[2]/div[2]/div[2]/div[1]/div[1]/div/button[1]")))
            boton.click()
            print("Se descargo el archivo CSV")
        except TimeoutException:
            print("Loading took too much time!")
        sleep(5)
        
        # Renombro el archivo
        archivo_csv = renombrar_temporal(usuario, mes, tipo_comprobante.value)
        print(f"Se renombró el archivo a {archivo_csv}")
        
        # Sumo el importe total
        total = sum_csv(archivo_csv, "Imp. Total")
        cprint(bcolors.OKGREEN, f"Total del mes {mes}, {usuario.apellido}: {total}")
        
        # Lo escribo en el archivo
        f = open(RESULTADOS, "a")
        f.write(f"Mes {mes} - {usuario.apellido} {usuario.nombre}\t{total}\n")
        f.close()
        
        driver.find_element_by_link_text("Consulta").click()


def sum_csv(archivo, columna):
    df = pd.read_csv(archivo)
    
    df_credito = df.loc[df['Tipo'].str.contains("Nota de Crédito")]
    suma_credito = df_credito[f'{columna}'].sum()
    
    df_debito = df[~df['Tipo'].str.contains("Nota de Crédito")]
    suma_debito = df_debito[f'{columna}'].sum()   
    
    return suma_debito - suma_credito


def crear_rango_fechas(un_mes):
    #anio = datetime.now().year
    anio = 2019
    dias = calendar.monthrange(anio, un_mes)[1]
    mes = f"{un_mes:02d}" # Leading zeroes
    fecha = f"01/{mes}/{anio} - {dias}/{mes}/{anio}"
    return fecha


def cargar_usuarios():
    xl = pd.read_excel(EXCEL, sheet_name="Clientes", skiprows=[0])
    df = pd.DataFrame(xl, columns=['Nombre', 'Apellido', 'CUIT', 'Clave AFIP'])\
        .rename(columns={'Clave AFIP': 'Clave'})
    df = df[df.Clave.notnull()] # Filtro los que no tienen la clave
    for row in df.itertuples():
        nuevo_usuario = Usuario(row.Nombre, row.Apellido, row.CUIT, row.Clave)
        USUARIOS.append(nuevo_usuario)


def renombrar_temporal(usuario, mes, tipo_comprobante):
    archivos = os.listdir(CARPETA_DESCARGAS)
    for f in archivos:
        if usuario.cuit in f:
            nombre = f"{usuario.apellido} {usuario.nombre} {tipo_comprobante} {mes}.csv"
            viejo = os.path.join(CARPETA_DESCARGAS, f)
            nuevo = os.path.join(CARPETA_DESCARGAS, nombre)
            os.rename(viejo, nuevo)
            return nuevo

def get_cliente(nombre):
    return next(x for x in iter(USUARIOS) if nombre.lower() in f"{x.nombre} {x.apellido} {x.nombre}".lower())

def cargar_error(usuario, msg=None):
    f = open("errores.txt", "a")
    f.write(f"{usuario.apellido} {usuario.nombre}")
    if msg:
        f.write(f": {msg}")
    f.write("\n")
    f.close()
    global ERRORES
    ERRORES += 1


cargar_usuarios()
#for u in USUARIOS: print(u.nombre)

# user = get_cliente("mariano pere")
# USUARIOS = [user]
# print(user)

# Downloads Folder
chrome_options = webdriver.ChromeOptions()
prefs = {'download.default_directory': CARPETA_DESCARGAS}
chrome_options.add_experimental_option('prefs', prefs)
# chrome_options.add_argument("--headless")

meses = list(range(1, 13))

for index, user in enumerate(USUARIOS):
    separador()
    for i in range(INTENTOS):
        driver = webdriver.Chrome(options=chrome_options)
        try:
            cprint(bcolors.OKBLUE, f"Usuario {index+1}/{len(USUARIOS)}: {user.apellido} {user.nombre}")
            afip_bot(driver, user, meses, Comprobantes.RECIBIDOS)
            driver.quit()
            break  # Salgo del for de intentos
        except AuthError as e:
            cprint(bcolors.FAIL, f"Error: {e}\n{user}")
            driver.quit()
            cargar_error(user, e)
            break
        except Exception as e:
            cprint(bcolors.WARNING, f"Intento {i+1}: Hubo un error con {user}:")
            cprint(bcolors.WARNING, f"Error: {e}")
            driver.quit()
            if i == INTENTOS-1:
                cprint(bcolors.FAIL ,"Limite de intentos")
                cargar_error(user)
