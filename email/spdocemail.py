from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
import pdb

f = open("C:\\Users\\Cliente\\OneDrive\\Documentos\\Automa-o-spdoc\\modelo_consulta.txt", "r", encoding="utf-8")
texto = f.read()

driver = webdriver.Chrome(executable_path=r'chromedriver.exe')
driver.get("https://webmail.fazenda.sp.gov.br/owa/")


driver.implicitly_wait(5)
driver.find_element(By.CSS_SELECTOR, '[autoid="_n_j"]').click()

driver.implicitly_wait(5)
#email
WebDriverWait(driver, timeout=20).until(lambda d: d.find_element(By.CSS_SELECTOR, '[autoid="_fp_7"]'))
driver.find_element(By.CSS_SELECTOR, '[autoid="_fp_7"]').send_keys("flavioafj@yahoo.com.br")

#assunto
driver.implicitly_wait(5)
driver.find_element(By.CSS_SELECTOR, '[autoid="_mcp_k"]').send_keys("teste")

#pdb.set_trace()

#corpo do texto
driver.implicitly_wait(5)
WebDriverWait(driver, timeout=20).until(lambda e: e.find_element(By.CSS_SELECTOR, '#EditorBody'))
driver.execute_script('var iframe = document.getElementById("EditorBody");var quadro = iframe.contentWindow.document.getElementById("MicrosoftOWAEditorRegion");var paragrafos = quadro.getElementsByTagName("p"); paragrafos[0].innerHTML = "' + texto + '";')



driver.implicitly_wait(5)
#driver.find_element(By.CSS_SELECTOR, '[autoid="_mcp_6"]').click()

