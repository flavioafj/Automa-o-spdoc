from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
import pdb


driver = webdriver.Chrome(executable_path=r'chromedriver.exe')
driver.get("https://webmail.fazenda.sp.gov.br/owa/")


driver.implicitly_wait(5)
driver.find_element(By.CSS_SELECTOR, '[autoid="_n_j"]').click()

#email
WebDriverWait(driver, timeout=3).until(lambda d: d.find_element(By.CSS_SELECTOR, '[autoid="_fp_7"]'))
driver.find_element(By.CSS_SELECTOR, '[autoid="_fp_7"]').send_keys("flavioafj@yahoo.com.br")

#assunto
driver.implicitly_wait(3)
driver.find_element(By.CSS_SELECTOR, '[autoid="_mcp_k"]').send_keys("teste")

#corpo do texto
driver.implicitly_wait(2)
driver.execute_script('var iframe = document.getElementById("EditorBody");var quadro = iframe.contentWindow.document.getElementById("MicrosoftOWAEditorRegion");var paragrafos = quadro.getElementsByTagName("p"); paragrafos[0].innerHTML = "<u> 2oi </u>";')



driver.implicitly_wait(5)
driver.find_element(By.CSS_SELECTOR, '[autoid="_mcp_6"]').click()

