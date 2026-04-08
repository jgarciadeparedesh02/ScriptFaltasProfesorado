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
# Configura el log para que escriba en consola y en un archivo .log
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(levelname)-8s | %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[
        logging.FileHandler("automatizacion_rayuela.log", encoding='utf-8'),
        logging.StreamHandler() # Sigue mostrando la info en la consola
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
logging.info(f"📂 Leyendo {ARCHIVO_CSV}...")

# Intentamos leer con diferentes codificaciones para evitar errores
try:
    df = pd.read_csv(ARCHIVO_CSV, sep=';', dtype=str, encoding='utf-8').fillna("")
except UnicodeDecodeError:
    logging.warning("⚠️ Error de codificación UTF-8 detectado. Cambiando a latin-1.")
    df = pd.read_csv(ARCHIVO_CSV, sep=';', dtype=str, encoding='latin-1').fillna("")

# Verificar si existe la columna 'Registrado'. Si no, la creamos vacía.
if COLUMNA_ESTADO not in df.columns:
    logging.info(f"ℹ️ La columna '{COLUMNA_ESTADO}' no existía. Creándola...")
    df[COLUMNA_ESTADO] = ""

# Preparamos el archivo de errores (se sobrescribe en cada ejecución nueva)
f_err = open(ARCHIVO_ERRORES, 'w', newline='', encoding='utf-8')
writer = csv.writer(f_err, delimiter=';')
writer.writerow(df.columns) # Copiamos cabeceras

ok_count = 0
err_count = 0
omitidos_count = 0

# ==========================================
# 2. INICIO DEL NAVEGADOR
# ==========================================
logging.info("🚀 Iniciando el navegador Chrome...")
driver = webdriver.Chrome()
driver.maximize_window()
driver.get(URL_RAYUELA)

logging.info("🛑 PASO 1: Inicia sesión en Rayuela.")
logging.info("🛑 PASO 2: Navega hasta el listado de profesores (Página 1).")
input("🟢 Pulsa ENTER en esta consola cuando estés listo para empezar...")
logging.info("▶️ Usuario confirmó inicio. Arrancando procesamiento...")

ventana_principal = driver.current_window_handle

# ==========================================
# 3. FUNCIONES AUXILIARES
# ==========================================

def ir_a_cuerpo():
    """Navega por los frames hasta llegar al listado/formulario."""
    driver.switch_to.default_content()
    wait = WebDriverWait(driver, 2)
    try:
        wait.until(EC.frame_to_be_available_and_switch_to_it("inferior"))
        wait.until(EC.frame_to_be_available_and_switch_to_it("principal"))
        wait.until(EC.frame_to_be_available_and_switch_to_it("centro"))
        wait.until(EC.frame_to_be_available_and_switch_to_it("cuerpo"))
    except Exception as e: 
        logging.debug(f"Aviso al navegar al frame cuerpo: {e}")

def ir_a_botonera():
    """Navega por los frames hasta llegar al botón Aceptar."""
    driver.switch_to.default_content()
    wait = WebDriverWait(driver, 2)
    try:
        wait.until(EC.frame_to_be_available_and_switch_to_it("inferior"))
        wait.until(EC.frame_to_be_available_and_switch_to_it("principal"))
        wait.until(EC.frame_to_be_available_and_switch_to_it("centro"))
        wait.until(EC.frame_to_be_available_and_switch_to_it("botoneraTitulo"))
    except Exception as e: 
        logging.debug(f"Aviso al navegar al frame botonera: {e}")

def click_seguro(by_tipo, selector, tiempo_espera=5):
    """Intenta hacer clic resistiendo errores de elementos caducados (Stale)."""
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
    """Limpia el texto: MAYÚSCULAS, SIN TILDES y Ñ -> N."""
    if not isinstance(texto, str): return ""
    texto = texto.upper().replace('Ñ', 'N')
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

def buscar_con_paginacion(nombre_original):
    """Busca nombre pasando páginas si es necesario."""
    paginas = 0
    MAX_PAGINAS = 50 
    nombre_buscado = normalizar_nombre(nombre_original)
    
    origen  = "abcdefghijklmnopqrstuvwxyzáéíóúñÁÉÍÓÚÑ"
    destino = "ABCDEFGHIJKLMNOPQRSTUVWXYZAEIOUNAEIOUN"
    
    while paginas < MAX_PAGINAS:
        ir_a_cuerpo()
        xpath_nombre = f"//a[contains(translate(text(), '{origen}', '{destino}'), '{nombre_buscado}')]"
        
        if click_seguro(By.XPATH, xpath_nombre, tiempo_espera=1):
            logging.info(f"✅ Encontrado en página {paginas + 1}")
            return True
        
        try:
            xpath_siguiente = "//a[contains(text(), 'Siguiente >')]"
            btn_siguiente = driver.find_element(By.XPATH, xpath_siguiente)
            btn_siguiente.click()
            paginas += 1
            logging.debug(f"Avanzando a página {paginas + 1}...")
            time.sleep(3) 
        except NoSuchElementException:
            return False 
            
    return False

# ==========================================
# 4. BUCLE PRINCIPAL
# ==========================================
total_filas = len(df)
for index, fila in df.iterrows():
    nombre = fila['Nombre']
    
    if str(fila[COLUMNA_ESTADO]).strip().upper() == 'OK':
        logging.info(f"[{index+1}/{total_filas}] ⏭️ Saltando: {nombre} (Ya registrado).")
        omitidos_count += 1
        continue

    f_ini = fila['Fecha Inicio']
    f_fin = fila['Fecha Fin']
    horas = fila['Horas Lectivas']
    motivo = fila['Motivo']

    logging.info(f"[{index+1}/{total_filas}] 🔄 Procesando: {nombre}")

    try:
        # A. BÚSQUEDA DEL PROFESOR
        encontrado = buscar_con_paginacion(nombre)
        if not encontrado:
            raise Exception("No encontrado en ninguna página del listado.")

        # B. ABRIR MENÚ "Nueva Ausencia"
        time.sleep(1)
        exito_menu = False
        try: exito_menu = click_seguro(By.ID, "menuItemText0", tiempo_espera=2)
        except: pass
        
        if not exito_menu:
            driver.switch_to.default_content()
            exito_menu = click_seguro(By.ID, "menuItemText0", tiempo_espera=2)
            
        if not exito_menu:
            raise Exception("No se pudo hacer clic en 'Nueva Ausencia'.")

        # C. RELLENAR FORMULARIO
        time.sleep(2) 
        if len(driver.window_handles) > 1: driver.switch_to.window(driver.window_handles[-1])
        else: ir_a_cuerpo()

        logging.info("📝 Rellenando formulario de ausencia...")
        wait = WebDriverWait(driver, 5)
        
        try:
            c_ini = wait.until(EC.presence_of_element_located((By.NAME, "F_INICIO")))
            c_ini.clear(); c_ini.send_keys(f_ini); c_ini.send_keys(Keys.TAB)
            
            c_fin = driver.find_element(By.NAME, "F_FINAL")
            c_fin.clear(); c_fin.send_keys(f_fin); c_fin.send_keys(Keys.TAB)
        except TimeoutException:
            raise Exception("El formulario no cargó (Fallo al encontrar F_INICIO).")

        time.sleep(1)
        if horas != "":
            try:
                c_horas = driver.find_element(By.NAME, "N_HORASL")
                c_horas.clear(); c_horas.send_keys(horas); c_horas.send_keys(Keys.ENTER)
                c_horas_comp = driver.find_element(By.NAME, "N_HORASC")
                c_horas_comp.clear(); c_horas_comp.send_keys("00:00"); c_horas_comp.send_keys(Keys.ENTER)
            except Exception as e:
                logging.warning(f"Aviso al rellenar horas para {nombre}: {e}")

        try:
            xpath_trigger = "//img[@id='img_id_comboC_MOTIVO' and contains(@class, 'search')]"
            if not click_seguro(By.XPATH, xpath_trigger, tiempo_espera=2):
                driver.find_element(By.ID, "special_input_id_comboC_MOTIVO").click()
            time.sleep(1)
            xpath_opcion = f"//div[@id='autocomplete_id_comboC_MOTIVO']/div[contains(text(), '{motivo}')]"
            wait.until(EC.element_to_be_clickable((By.XPATH, xpath_opcion))).click()
        except Exception as e:
            raise Exception(f"Error seleccionando Motivo '{motivo}': {e}")

        # D. GUARDAR Y ACTUALIZAR CSV
        logging.info("💾 Guardando ausencia...")
        ir_a_botonera()
        time.sleep(0.5)
        
        if click_seguro(By.ID, "i_ACEPTAR", tiempo_espera=3):
            logging.info(f"✅ ¡{nombre} guardado en Rayuela correctamente!")
            ok_count += 1
            
            df.at[index, COLUMNA_ESTADO] = 'OK'
            df.to_csv(ARCHIVO_CSV, sep=';', index=False, encoding='utf-8-sig')
            logging.info("📝 CSV actualizado: Registrado = OK")
            
            time.sleep(4)
        else:
            raise Exception("No se encontró o no se pudo pulsar el botón 'Aceptar'.")

    except Exception as e:
        logging.error(f"❌ FALLO con {nombre}: {e}")
        err_count += 1
        
        writer.writerow(fila)
        f_err.flush()
        
        try:
            ir_a_botonera()
            click_seguro(By.ID, "i_VOLVER", tiempo_espera=3)
        except Exception as recovery_error: 
            logging.error(f"Error al intentar recuperar la pestaña principal: {recovery_error}")
            

# CIERRE
f_err.close()
driver.quit() # Es buena práctica cerrar el driver al final

logging.info("="*40)
logging.info("📊 RESUMEN FINAL:")
logging.info(f"   ⏭️ Omitidos (Ya OK): {omitidos_count}")
logging.info(f"   ✅ Procesados OK:    {ok_count}")
logging.info(f"   ❌ Errores:          {err_count}")
logging.info(f"   📁 Archivo errores:  {ARCHIVO_ERRORES}")
logging.info("="*40)