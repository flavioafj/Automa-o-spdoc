from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pdb

driver = webdriver.Chrome('chromedriver.exe')
driver.get("https://www.documentos.spsempapel.sp.gov.br/siga/public/app/login?cont=https://www.documentos.spsempapel.sp.gov.br/siga/app/principal")

#===Autenticação
driver.find_element(By.ID, "username").send_keys("SFP25402")

senha = input("Enter senha:")
driver.find_element(By.ID, "password").send_keys(senha + Keys.ENTER)

driver.implicitly_wait(2)
driver.find_element(By.ID, "collapse-header-7").click()

#===escolha do processo na mesa
driver.implicitly_wait(2)
driver.find_element_by_link_text("SFP-EXP-2020/210311-A").click()


#===Uploads de documentos
driver.implicitly_wait(1)
driver.find_element(By.CSS_SELECTOR, '[accesskey="d"]').click()

driver.implicitly_wait(1)
driver.find_element(By.ID, "dropdownMenuButton").click()

driver.implicitly_wait(1)
driver.find_element(By.CSS_SELECTOR, '[placeholder="Pesquisar modelo..."]').send_keys("cap" + Keys.DOWN + Keys.ENTER)

#pdb.set_trace()


#pdb.set_trace()
driver.implicitly_wait(1)
driver.find_element(By.ID, "Assunto").send_keys("Nota Fiscal Paulista")
driver.implicitly_wait(2)
driver.find_element(By.ID, "especie").send_keys("Extrato")
driver.implicitly_wait(2)
driver.find_element(By.ID, "conferencia").send_keys("Cópia autenticada administrativamente")
driver.implicitly_wait(1)
driver.find_element(By.ID, "arquivo").send_keys("C:\\Users\\flavi\\Downloads\\Convite Assessment VLI (2).pdf")

driver.implicitly_wait(1)
driver.find_element(By.ID, "btnGravar").click()