# Automação para ler ads.txt com Selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from openpyxl import load_workbook
from urllib.parse import urlparse
import time

arquivo_excel = "checklist.xlsx"

df = pd.read_excel(arquivo_excel, skiprows=1, header=None)

wb = load_workbook(arquivo_excel)
ws = wb.active

options = webdriver.ChromeOptions()
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 20)

for index, row in df.iterrows():
    linha_excel = index + 2  

    url = str(row[0]).strip()
    if not url or url.lower() in ("nan", "none"):
        ws[f"B{linha_excel}"] = "(vazio)"
        ws[f"D{linha_excel}"] = "N/A"
        continue

    if not urlparse(url).scheme:
        url = "http://" + url

    if not url.endswith("/ads.txt"):
        if url.endswith("/"):
            url = url + "ads.txt"
        else:
            url = url + "/ads.txt"

    try:
        driver.get(url)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(1)

        body_text = driver.find_element(By.TAG_NAME, "body").text.strip()
        lines = [l for l in (ln.strip() for ln in body_text.splitlines()) if l]
        primeiro_texto = lines[0] if lines else ""

        ws[f"B{linha_excel}"] = primeiro_texto
        ws[f"C{linha_excel}"] = (body_text[:32000] if body_text else "")
        ws[f"D{linha_excel}"] = "OK"

    except Exception as e:
        ws[f"B{linha_excel}"] = f"ERRO: {str(e)[:200]}"
        ws[f"D{linha_excel}"] = "ERROR"

wb.save(arquivo_excel)
driver.quit()
print("Finalizado!")
