import time
import csv
import pandas as pd
import unicodedata
import os
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException

# ==========================================
# 0. CONFIGURACIÓN DE LOGS
# ==========================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(levelname)-8s | %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[
        logging.FileHandler("automatizacion_rayuela.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# --- CONFIGURACIÓN DE RUTAS ---
ARCHIVO_CSV = 'ausencias.csv'
ARCHIVO_ERRORES = 'errores.csv'
COLUMNA_ESTADO = 'Registrado'  # Nombre de la columna de control
URL_RAYUELA = "https://rayuela.educarex.es/modulo_acceso/"

# ==========================================
# 1. PREPARACIÓN DEL CSV Y DATOS
# ==========================================
logging.info(f"📂 Leyendo archivo: {ARCHIVO_CSV}")

try:
    # IMPORTANTE: utf-8-sig evita caracteres ocultos (BOM) de Excel
    df = pd.read_csv(ARCHIVO_CSV, sep=',', dtype=str, encoding='utf-8-sig').fillna("")
except UnicodeDecodeError:
    logging.warning("⚠️ Error UTF-8. Cambiando a latin-1.")
    df = pd.read_csv(ARCHIVO_CSV, sep=',', dtype=str, encoding='latin-1').fillna("")
except FileNotFoundError:
    logging.error(f"❌ No se encuentra el archivo {ARCHIVO_CSV}.")
    exit()

# Limpiamos espacios en las cabeceras por seguridad
df.columns = df.columns.str.strip()

if 'Nombre' not in df.columns:
    logging.error(f"❌ No se encontró la columna 'Nombre'. Columnas detectadas: {df.columns.tolist()}")
    exit()

if COLUMNA_ESTADO not in df.columns:
    logging.info(f"ℹ️ Creando columna de control '{COLUMNA_ESTADO}'...")
    df[COLUMNA_ESTADO] = ""

f_err = open(ARCHIVO_ERRORES, 'w', newline='', encoding='utf-8')
writer = csv.writer(f_err, delimiter=';')
writer.writerow(df.columns) 

ok_count = 0
err_count = 0
omitidos_count = 0

# ==========================================
# 2. INICIO DEL NAVEGADOR
# ==========================================
logging.info("🚀 Iniciando el navegador Chrome...")
driver = webdriver.Chrome()
driver.maximize_window()
logging.info(f"🌐 Navegando a: {URL_RAYUELA}")
driver.get(URL_RAYUELA)

logging.info("🛑 PASO 1: Inicia sesión en Rayuela.")
logging.info("🛑 PASO 2: Navega hasta el listado de profesores (Página 1).")
input("🟢 Pulsa ENTER en esta consola cuando estés listo para empezar...")
logging.info("▶️ Usuario confirmó inicio. Arrancando bucle de procesamiento...")

ventana_principal = driver.current_window_handle

# ==========================================
# 3. FUNCIONES AUXILIARES
# ==========================================

def ir_a_cuerpo():
    driver.switch_to.default_content()
    wait = WebDriverWait(driver, 2)
    try:
        wait.until(EC.frame_to_be_available_and_switch_to_it("inferior"))
        wait.until(EC.frame_to_be_available_and_switch_to_it("principal"))
        wait.until(EC.frame_to_be_available_and_switch_to_it("centro"))
        wait.until(EC.frame_to_be_available_and_switch_to_it("cuerpo"))
    except Exception as e: 
        logging.debug(f"Aviso frame cuerpo: {e}")

def ir_a_botonera():
    driver.switch_to.default_content()
    wait = WebDriverWait(driver, 2)
    try:
        wait.until(EC.frame_to_be_available_and_switch_to_it("inferior"))
        wait.until(EC.frame_to_be_available_and_switch_to_it("principal"))
        wait.until(EC.frame_to_be_available_and_switch_to_it("centro"))
        wait.until(EC.frame_to_be_available_and_switch_to_it("botoneraTitulo"))
    except Exception as e: 
        logging.debug(f"Aviso frame botonera: {e}")

def click_seguro(by_tipo, selector, tiempo_espera=5):
    intentos = 0
    while intentos < 3:
        try:
            elemento = WebDriverWait(driver, tiempo_espera).until(
                EC.element_to_be_clickable((by_tipo, selector))
            )
            elemento.click()
            return True
        except StaleElementReferenceException:
            time.sleep(1)
            intentos += 1
        except Exception:
            return False
    return False

def normalizar_nombre(texto):
    if not isinstance(texto, str): return ""
    texto = texto.upper().replace('Ñ', 'N')
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

def buscar_con_paginacion(nombre_original):
    paginas = 0
    MAX_PAGINAS = 50 
    nombre_buscado = normalizar_nombre(nombre_original)
    logging.info(f"   🔎 Buscando en listado a: '{nombre_buscado}'")
    
    origen  = "abcdefghijklmnopqrstuvwxyzáéíóúñÁÉÍÓÚÑ"
    destino = "ABCDEFGHIJKLMNOPQRSTUVWXYZAEIOUNAEIOUN"
    
    while paginas < MAX_PAGINAS:
        ir_a_cuerpo()
        xpath_nombre = f"//a[contains(translate(text(), '{origen}', '{destino}'), '{nombre_buscado}')]"
        
        if click_seguro(By.XPATH, xpath_nombre, tiempo_espera=1):
            logging.info(f"   ✅ Profesor encontrado en la página {paginas + 1}")
            return True
        
        try:
            xpath_siguiente = "//a[contains(text(), 'Siguiente >')]"
            btn_siguiente = driver.find_element(By.XPATH, xpath_siguiente)
            btn_siguiente.click()
            paginas += 1
            logging.info(f"   ➡️ Pasando a la página {paginas + 1}...")
            time.sleep(3) 
        except NoSuchElementException:
            logging.warning("   ⚠️ No hay más páginas y no se encontró al profesor.")
            return False 
            
    return False

# ==========================================
# 4. BUCLE PRINCIPAL
# ==========================================
total_filas = len(df)
for index, fila in df.iterrows():
    nombre = fila['Nombre']
    
    if str(fila[COLUMNA_ESTADO]).strip().upper() == 'OK':
        logging.info(f"[{index+1}/{total_filas}] ⏭️ Saltando: {nombre} (Ya marcado como OK).")
        omitidos_count += 1
        continue

    f_ini = fila['Fecha Inicio']
    f_fin = fila['Fecha Fin']
    horas = fila['Horas Lectivas']
    motivo = fila['Motivo']

    logging.info("-" * 50)
    logging.info(f"[{index+1}/{total_filas}] 🔄 PROCESANDO: {nombre}")
    logging.info(f"   📋 Datos a introducir -> Inicio: {f_ini} | Fin: {f_fin} | Horas: '{horas}' | Motivo: '{motivo}'")

    try:
        # A. BÚSQUEDA DEL PROFESOR
        encontrado = buscar_con_paginacion(nombre)
        if not encontrado:
            raise Exception(f"No se encontró a '{nombre}' en el listado.")

        # B. ABRIR MENÚ "Nueva Ausencia"
        logging.info("   🖱️ Intentando abrir menú contextual y hacer clic en 'Nueva Ausencia'...")
        time.sleep(1)
        exito_menu = False
        try: exito_menu = click_seguro(By.ID, "menuItemText0", tiempo_espera=2)
        except: pass
        
        if not exito_menu:
            driver.switch_to.default_content()
            exito_menu = click_seguro(By.ID, "menuItemText0", tiempo_espera=2)
            
        if not exito_menu:
            raise Exception("Falló el clic en 'Nueva Ausencia'.")
        logging.info("   ✅ Menú 'Nueva Ausencia' abierto correctamente.")

        # C. RELLENAR FORMULARIO
        logging.info("   📝 Esperando a que cargue el formulario...")
        time.sleep(2) 
        if len(driver.window_handles) > 1: 
            driver.switch_to.window(driver.window_handles[-1])
            logging.info("   🪟 Cambiado a nueva ventana emergente.")
        else: 
            ir_a_cuerpo()
            logging.info("   🪟 Usando frame del cuerpo en la misma ventana.")

        wait = WebDriverWait(driver, 5)
        
        # --- Escribir Fechas ---
        try:
            c_ini = wait.until(EC.presence_of_element_located((By.NAME, "F_INICIO")))
            logging.info(f"   ⌨️ Escribiendo Fecha Inicio: {f_ini}")
            c_ini.clear(); c_ini.send_keys(f_ini); c_ini.send_keys(Keys.TAB)
            
            c_fin = driver.find_element(By.NAME, "F_FINAL")
            logging.info(f"   ⌨️ Escribiendo Fecha Fin: {f_fin}")
            c_fin.clear(); c_fin.send_keys(f_fin); c_fin.send_keys(Keys.TAB)
        except TimeoutException:
            raise Exception("El formulario no cargó a tiempo (No se encontró F_INICIO).")

        # --- Escribir Horas ---
        time.sleep(1)
        if horas != "":
            try:
                logging.info(f"   ⌨️ Escribiendo Horas Lectivas: {horas}")
                c_horas = driver.find_element(By.NAME, "N_HORASL")
                c_horas.clear(); c_horas.send_keys(horas); c_horas.send_keys(Keys.ENTER)
                
                logging.info(f"   ⌨️ Estableciendo Horas Complementarias a '00:00'")
                c_horas_comp = driver.find_element(By.NAME, "N_HORASC")
                c_horas_comp.clear(); c_horas_comp.send_keys("00:00"); c_horas_comp.send_keys(Keys.ENTER)
            except Exception as e:
                logging.warning(f"   ⚠️ Problema menor al rellenar horas: {e}")
        else:
            logging.info("   ⏩ Campo 'Horas Lectivas' vacío en CSV. Se omite.")

        # --- Seleccionar Motivo ---
        try:
            logging.info(f"   🖱️ Desplegando combo de Motivos buscando: '{motivo}'")
            xpath_trigger = "//img[@id='img_id_comboC_MOTIVO' and contains(@class, 'search')]"
            if not click_seguro(By.XPATH, xpath_trigger, tiempo_espera=2):
                driver.find_element(By.ID, "special_input_id_comboC_MOTIVO").click()
            time.sleep(1)
            xpath_opcion = f"//div[@id='autocomplete_id_comboC_MOTIVO']/div[contains(text(), '{motivo}')]"
            wait.until(EC.element_to_be_clickable((By.XPATH, xpath_opcion))).click()
            logging.info("   ✅ Motivo seleccionado con éxito.")
        except Exception as e:
            raise Exception(f"Error seleccionando el motivo '{motivo}': {e}")

        # D. GUARDAR Y ACTUALIZAR CSV
        logging.info("   💾 Navegando a la botonera para guardar...")
        ir_a_cuerpo() # A veces hay que volver a centrar
        ir_a_botonera()
        time.sleep(0.5)
        
        logging.info("   🖱️ Pulsando botón 'Aceptar' (i_ACEPTAR)...")
        if click_seguro(By.ID, "i_ACEPTAR", tiempo_espera=3):
            logging.info(f"   🎉 ¡EXITO! Ausencia de {nombre} registrada en Rayuela.")
            ok_count += 1
            
            df.at[index, COLUMNA_ESTADO] = 'OK'
            df.to_csv(ARCHIVO_CSV, sep=';', index=False, encoding='utf-8-sig')
            logging.info("   📝 Archivo CSV actualizado (Registrado = OK).")
            
            logging.info("   ⏳ Esperando a que la web vuelva al listado principal (4s)...")
            time.sleep(4)
        else:
            raise Exception("No se pudo hacer clic en el botón 'Aceptar'.")

    except Exception as e:
        logging.error(f"   ❌ FALLO CRÍTICO procesando a {nombre}: {e}")
        err_count += 1
        
        writer.writerow(fila)
        f_err.flush()
        
        logging.info("   🔄 Iniciando protocolo de recuperación para el siguiente profesor...")
        try:
            ir_a_botonera()
            logging.info("   🖱️ Intentando pulsar 'Volver' (i_VOLVER)...")
            click_seguro(By.ID, "i_VOLVER", tiempo_espera=3)
        except Exception as recovery_error: 
            logging.error(f"   ⚠️ Falló la recuperación automática: {recovery_error}")
            # Intento de fuerza bruta para volver
            try:
                if len(driver.window_handles) > 1: driver.close()
                driver.switch_to.window(ventana_principal)
            except: pass

# CIERRE
f_err.close()
driver.quit() 

logging.info("="*50)
logging.info("📊 RESUMEN FINAL DE LA EJECUCIÓN:")
logging.info(f"   ⏭️ Filas omitidas (Ya estaban OK): {omitidos_count}")
logging.info(f"   ✅ Nuevas ausencias registradas:   {ok_count}")
logging.info(f"   ❌ Errores encontrados:            {err_count}")
logging.info(f"   📁 Archivo de errores actualizado: {ARCHIVO_ERRORES}")
logging.info("="*50)