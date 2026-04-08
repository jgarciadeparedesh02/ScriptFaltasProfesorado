import time
import csv
import pandas as pd
import unicodedata
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException

# --- CONFIGURACIÓN ---
ARCHIVO_CSV = 'ausencias.csv'
ARCHIVO_ERRORES = 'errores.csv'
COLUMNA_ESTADO = 'Registrado'  # Nombre de la columna de control
URL_RAYUELA = "https://rayuela.educarex.es/modulo_acceso/"

# ==========================================
# 1. PREPARACIÓN DEL CSV Y DATOS
# ==========================================
print(f"📂 Leyendo {ARCHIVO_CSV}...")

# Intentamos leer con diferentes codificaciones para evitar errores
try:
    df = pd.read_csv(ARCHIVO_CSV, sep=',', dtype=str, encoding='utf-8').fillna("")
except UnicodeDecodeError:
    df = pd.read_csv(ARCHIVO_CSV, sep=',', dtype=str, encoding='latin-1').fillna("")

# Verificar si existe la columna 'Registrado'. Si no, la creamos vacía.
if COLUMNA_ESTADO not in df.columns:
    print(f"   ℹ️ La columna '{COLUMNA_ESTADO}' no existía. Creándola...")
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
driver = webdriver.Chrome()
driver.maximize_window()
driver.get(URL_RAYUELA)

print("\n🛑 PASO 1: Inicia sesión en Rayuela.")
print("🛑 PASO 2: Navega hasta el listado de profesores (Página 1).")
input("🟢 Pulsa ENTER en esta consola cuando estés listo para empezar...")

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
    except: pass

def ir_a_botonera():
    """Navega por los frames hasta llegar al botón Aceptar."""
    driver.switch_to.default_content()
    wait = WebDriverWait(driver, 2)
    try:
        wait.until(EC.frame_to_be_available_and_switch_to_it("inferior"))
        wait.until(EC.frame_to_be_available_and_switch_to_it("principal"))
        wait.until(EC.frame_to_be_available_and_switch_to_it("centro"))
        wait.until(EC.frame_to_be_available_and_switch_to_it("botoneraTitulo"))
    except: pass

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
    
    # XPath mágico: ignora mayúsculas, minúsculas, tildes y eñes
    origen  = "abcdefghijklmnopqrstuvwxyzáéíóúñÁÉÍÓÚÑ"
    destino = "ABCDEFGHIJKLMNOPQRSTUVWXYZAEIOUNAEIOUN"
    
    while paginas < MAX_PAGINAS:
        ir_a_cuerpo()
        
        # Busca si el nombre está en la pantalla actual
        xpath_nombre = f"//a[contains(translate(text(), '{origen}', '{destino}'), '{nombre_buscado}')]"
        
        if click_seguro(By.XPATH, xpath_nombre, tiempo_espera=1):
            print(f"   ✅ Encontrado (Página {paginas + 1})")
            return True
        
        # Si no, busca botón "Siguiente >"
        # print(f"   ...buscando en pág {paginas + 1}...") # (Opcional: Descomentar para ver log detallado)
        try:
            xpath_siguiente = "//a[contains(text(), 'Siguiente >')]"
            btn_siguiente = driver.find_element(By.XPATH, xpath_siguiente)
            btn_siguiente.click()
            paginas += 1
            time.sleep(3) # Espera de carga
        except NoSuchElementException:
            return False # No hay más páginas
            
    return False

# ==========================================
# 4. BUCLE PRINCIPAL
# ==========================================
for index, fila in df.iterrows():
    nombre = fila['Nombre']
    
    # --- FILTRO DE ESTADO ---
    # Si la columna 'Registrado' ya tiene 'OK', saltamos.
    if str(fila[COLUMNA_ESTADO]).strip().upper() == 'OK':
        print(f"\n[{index+1}/{len(df)}] ⏭️ Saltando: {nombre} (Ya registrado).")
        omitidos_count += 1
        continue

    # Si no tiene OK, procesamos
    f_ini = fila['Fecha Inicio']
    f_fin = fila['Fecha Fin']
    horas = fila['Horas Lectivas']
    motivo = fila['Motivo']

    print(f"\n[{index+1}/{len(df)}] Procesando: {nombre}")

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
            driver.switch_to.default_content() # Busca fuera del frame
            exito_menu = click_seguro(By.ID, "menuItemText0", tiempo_espera=2)
        if not exito_menu:
            raise Exception("No se pudo hacer clic en 'Nueva Ausencia'.")

        # C. RELLENAR FORMULARIO
        time.sleep(2) 
        if len(driver.window_handles) > 1: driver.switch_to.window(driver.window_handles[-1])
        else: ir_a_cuerpo()

        print("   -> Rellenando formulario...")
        wait = WebDriverWait(driver, 5)
        
        # Fechas
        try:
            c_ini = wait.until(EC.presence_of_element_located((By.NAME, "F_INICIO")))
            c_ini.clear(); c_ini.send_keys(f_ini); c_ini.send_keys(Keys.TAB)
            
            c_fin = driver.find_element(By.NAME, "F_FINAL")
            c_fin.clear(); c_fin.send_keys(f_fin); c_fin.send_keys(Keys.TAB)
        except TimeoutException:
            raise Exception("El formulario no cargó (Fallo F_INICIO).")

        # Horas (si aplica) print (horas)
        time.sleep(1)
        if horas != "":
            try:
                c_horas = driver.find_element(By.NAME, "N_HORASL")
                c_horas.clear(); c_horas.send_keys(horas); c_horas.send_keys(Keys.ENTER)
                c_horas_comp = driver.find_element(By.NAME, "N_HORASC")
                c_horas_comp.clear(); c_horas_comp.send_keys("00:00"); c_horas_comp.send_keys(Keys.ENTER)
            except: pass

        # Motivo (Selección compleja en lista)
        try:
            xpath_trigger = "//img[@id='img_id_comboC_MOTIVO' and contains(@class, 'search')]"
            if not click_seguro(By.XPATH, xpath_trigger, tiempo_espera=2):
                driver.find_element(By.ID, "special_input_id_comboC_MOTIVO").click()
            time.sleep(1)
            xpath_opcion = f"//div[@id='autocomplete_id_comboC_MOTIVO']/div[contains(text(), '{motivo}')]"
            wait.until(EC.element_to_be_clickable((By.XPATH, xpath_opcion))).click()
            # print("      ✅ Motivo seleccionado.")
        except Exception as e:
            raise Exception(f"Error seleccionando Motivo: {e}")

        # D. GUARDAR Y ACTUALIZAR CSV
        print("   -> Guardando...")
        ir_a_botonera()
        time.sleep(0.5)
        
        # IMPORTANTE: ID del botón de aceptar
        if click_seguro(By.ID, "i_ACEPTAR", tiempo_espera=3):
            print("   💾 ¡Guardado en Rayuela!")
            ok_count += 1
            
            # --- ACTUALIZACIÓN DEL CSV ---
            # 1. Actualizamos la celda en memoria
            df.at[index, COLUMNA_ESTADO] = 'OK'
            
            # 2. Guardamos el archivo en disco (utf-8-sig para compatibilidad Excel/Tildes)
            df.to_csv(ARCHIVO_CSV, sep=';', index=False, encoding='utf-8-sig')
            print("      📝 CSV actualizado: Registrado = OK")
            
            time.sleep(4) # Esperamos a que la web vuelva al listado
        else:
            raise Exception("No se encontró el botón 'i_ACEPTAR' en la botonera.")

    except Exception as e:
        print(f"   ❌ ERROR: {e}")
        err_count += 1
        
        # Guardamos en log de errores
        writer.writerow(fila)
        f_err.flush()
        
        # Intentamos recuperar el control volviendo al inicio
        try:
            if len(driver.window_handles) > 1: driver.close()
            driver.switch_to.window(ventana_principal)
        except: pass

# CIERRE
f_err.close()
print("\n" + "="*40)
print(f"RESUMEN FINAL:")
print(f"⏭️ Omitidos (Ya OK): {omitidos_count}")
print(f"✅ Procesados OK:    {ok_count}")
print(f"❌ Errores:          {err_count}")
print(f"📁 Errores en:       {ARCHIVO_ERRORES}")
print("="*40)