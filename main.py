import time
import random
import re
import logging
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager
import os

def check_run_limit(limit=3):
    appdata = os.getenv('APPDATA')
    folder_path = os.path.join(appdata, "scraper_demo")
    file_path = os.path.join(folder_path, "run_count.txt")

    # Crear carpeta si no existe
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        # Marcar como oculta
        os.system(f'attrib +h "{folder_path}"')

    # Crear archivo si no existe (primera ejecución)
    if not os.path.exists(file_path):
        with open(file_path, "w") as f:
            f.write("1")
        return True

    # Leer número de ejecuciones
    with open(file_path, "r") as f:
        try:
            count = int(f.read().strip())
        except ValueError:
            count = 0  # Corrompido = empezar desde 0

    if count >= limit:
        print("❌ Límite de ejecuciones alcanzado. Contacta al desarrollador.")
        return False

    # Incrementar contador
    with open(file_path, "w") as f:
        f.write(str(count + 1))
    return True

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("scraper.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def create_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

def scroll_results(driver, max_results, scroll_pause=2, max_time=120):
    logger.info("📜 Iniciando scroll para cargar más resultados...")
    start_time = time.time()
    attempts_without_new = 0
    max_attempts = 10
    last_count = 0

    try:
        panel = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[@role='feed']"))
        )
    except TimeoutException:
        logger.error("❌ No se pudo localizar el panel de resultados.")
        return False

    while time.time() - start_time < max_time:
        cards = driver.find_elements(By.XPATH, "//div[contains(@class,'Nv2PK')]")
        links = driver.find_elements(By.XPATH, "//a[contains(@href, '/maps/place')]")
        unique_links = list({link.get_attribute("href") for link in links if link.get_attribute("href")})
        current_count = len(unique_links)
        logger.info(f"🔎 Resultados únicos encontrados: {current_count}/{max_results}")

        if current_count >= max_results:
            break

        if current_count == last_count:
            attempts_without_new += 1
            if attempts_without_new >= max_attempts:
                logger.warning("⚠️ No se detectaron nuevos resultados tras múltiples intentos.")
                break
        else:
            attempts_without_new = 0
            last_count = current_count

        try:
            panel.click()
            for _ in range(5):  # envía flechas abajo para simular navegación humana
                panel.send_keys(Keys.DOWN)
                time.sleep(random.uniform(0.2, 0.4))

            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", panel)
            time.sleep(random.uniform(scroll_pause, scroll_pause + 1))
        except Exception as e:
            logger.warning(f"Error durante scroll: {e}")
            break

    return True

def extract_business_data(driver, city_filter):
    def safe_get(xpath, attr=None):
        try:
            el = driver.find_element(By.XPATH, xpath)
            return el.get_attribute(attr) if attr else el.text
        except Exception:
            return None

    name = safe_get("//h1[contains(@class,'DUwDvf')]")
    phone = None
    website = None
    address = None

    try:
        info_blocks = driver.find_elements(By.XPATH, "//div[contains(@class, 'Io6YTe')]")
        for el in info_blocks:
            text = el.text.strip()
            if not phone and re.match(r"^\+?\d[\d\s.\-()]{7,}\d$", text):
                phone = text
            elif not address and "," in text and city_filter.lower() in text.lower():
                address = text
    except Exception:
        pass

    if not website:
        try:
            website_button = driver.find_element(By.XPATH, "//a[contains(., 'Sitio web') or contains(., 'Website')]")
            website = website_button.get_attribute('href')
        except Exception:
            pass

    if not website:
        try:
            for el in info_blocks:
                text = el.text.strip()
                if "." in text and " " not in text and not text.startswith("+"):
                    website = "https://" + text if not text.startswith("http") else text
                    break
        except Exception:
            pass

    if not phone:
        try:
            tel_links = driver.find_elements(By.XPATH, "//a[starts-with(@href, 'tel:')]")
            if tel_links:
                phone = tel_links[0].get_attribute('href').replace('tel:', '')
        except Exception:
            pass

    city = None
    if address:
        try:
            parts = address.split(",")
            city = parts[-2].strip() + " " + parts[-1].strip().split()[0] if len(parts) > 1 else parts[-1].strip()
        except Exception:
            city = None

    if city and city_filter.lower() not in city.lower():
        return None

    if not name or not address:
        return None

    return {
        "Nombre": name,
        "Teléfono": phone,
        "Web": website,
        "Dirección": address,
        "Ciudad": city,
        "URL": driver.current_url
    }

def get_user_input():
    print("\n" + "="*50)
    print("🔍 SCRAPER DE GOOGLE MAPS".center(50))
    print("="*50 + "\n")
    
    while True:
        search_term = input("Ingrese el término de búsqueda (ej: 'clínicas estéticas en Málaga'): ").strip()
        if search_term:
            break
        print("⚠️ El término de búsqueda no puede estar vacío")
    
    while True:
        try:
            max_results = int(input("Ingrese el número máximo de resultados a extraer (1-100): "))
            if 1 <= max_results <= 100:
                break
            print("⚠️ Por favor ingrese un número entre 1 y 100")
        except ValueError:
            print("⚠️ Debe ingresar un número válido")
    
    return search_term, max_results

def main():
    if not check_run_limit():
        return
    search_term, max_results = get_user_input()
    driver = create_driver()
    
    try:
        driver.set_page_load_timeout(30)
        driver.get("https://www.google.com/maps")
        wait = WebDriverWait(driver, 20)

        try:
            accept = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(., 'Aceptar todo') or contains(., 'Accept all')]")))
            accept.click()
        except:
            logger.info("No se mostró el botón de cookies")

        search_input = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//input[@id='searchboxinput']")))
        search_input.clear()
        search_input.send_keys(search_term)
        search_input.send_keys(Keys.RETURN)

        wait.until(EC.presence_of_element_located(
            (By.XPATH, "//div[contains(@class, 'Nv2PK')]")))

        scroll_results(driver, max_results)

        links = driver.find_elements(By.XPATH, "//a[contains(@href, '/maps/place')]")
        valid_links = list({link.get_attribute("href") for link in links if link.get_attribute("href")})
        business_links = valid_links[:max_results]
        logger.info(f"🔗 Se encontraron {len(business_links)} enlaces únicos de {max_results} solicitados")

        extracted_data = []
        city_name = search_term.split(" en ")[-1]

        for i, link in enumerate(business_links, 1):
            try:
                logger.info(f"\n📄 Procesando negocio {i}/{len(business_links)}")
                logger.info(f"🌐 URL: {link}")
                
                for attempt in range(2):
                    try:
                        driver.get(link)
                        time.sleep(random.uniform(3.5, 6))
                        driver.execute_script("window.scrollBy(0, 300);")
                        time.sleep(random.uniform(0.5, 1.5))
                        break
                    except WebDriverException as e:
                        logger.warning(f"Intento {attempt + 1} fallido para {link}: {e}")
                        time.sleep(random.uniform(2, 4))
                        if attempt == 1:
                            logger.error(f"❌ Error al procesar {link} después de reintentos")
                            break

                data = extract_business_data(driver, city_name)
                if data:
                    extracted_data.append(data)
                    logger.info(f"✅ Datos extraídos: {data['Nombre']}")
                else:
                    logger.warning("⚠️ Negocio sin datos válidos o filtrado por ciudad")
            except Exception as e:
                logger.error(f"❌ Error inesperado en {link}: {str(e)}")
                continue

        if extracted_data:
            df = pd.DataFrame(extracted_data)
            filename = re.sub(r'[^\w\-_]', '_', search_term)[:50]
            output_file = f"{filename}.xlsx"

            writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='Negocios')

            workbook = writer.book
            worksheet = writer.sheets['Negocios']

            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#4472C4',
                'font_color': 'white',
                'border': 1
            })

            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
                column_width = max(df[value].astype(str).map(len).max(), len(value)) + 2
                worksheet.set_column(col_num, col_num, column_width)

            writer.close()
            logger.info(f"📁 Archivo guardado correctamente: {output_file}")
        else:
            logger.warning("⚠️ No se extrajeron datos de negocios")

    finally:
        driver.quit()
        logger.info("🧹 Driver cerrado correctamente")

if __name__ == "__main__":
    main()
