import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.common.exceptions import TimeoutException
import time
# Função para aguardar a presença do elemento
def wait_for_element_visible(by, value, timeout=3):
    try:
        time.sleep(2)
        contato_element = driver.find_element("xpath", "//span[@class='p357zi0d r15c9g6i']".format(search_term))
        if(contato_element.click()):
            return True
    except TimeoutException:
        print(f"Elemento {value} não encontrado após {timeout} segundos.")
        return False

# Carregando o arquivo Excel
tabela = pd.read_excel('contatos.xlsx', engine='openpyxl')

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")

chrome_driver_path = './drivers/chromedriver.exe'

service = ChromeService(executable_path=chrome_driver_path)
driver = webdriver.Chrome(service=service, options=options)

driver.get('https://web.whatsapp.com/')
print("Aguarde enquanto o WhatsApp Web é carregado...")

input("Pressione Enter após escanear o código QR no WhatsApp Web...")

for index, contato in tabela.iterrows():
    search_box = driver.find_element("xpath", "//div[@class='g0rxnol2 ln8gz9je lexical-rich-text-input']")

    # Use o número de telefone para pesquisar o contato
    search_term = str(contato['numero'])

    # Simule a digitação como se fosse feita pelo usuário
    ActionChains(driver).send_keys_to_element(search_box, search_term).perform()

    try:
        # Aguarde a visibilidade do elemento
        if wait_for_element_visible(By.XPATH, "//div[@class='lhggkp7q ln8gz9je rx9719la']".format(search_term)):
            print('Entrou aqui')
        else:
            try:
                WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@class='_13mgZ']"))
            )
            except TimeoutException as e:
                print(f"Tempo de espera expirado: {e}")

            media_path = contato["caminho_arquivo"]

            if pd.isna(media_path):
                send_box = driver.find_element("xpath", "//div[@class='to2l77zo gfz4du6o ag5g9lrv bze30y65 kao4egtt']")

                # Mensagem para o contato
                send_term = contato["mensagem"]

                ActionChains(driver).send_keys_to_element(send_box, send_term).perform()

                time.sleep(2)

                # Pressione Enter para enviar a mensagem
                ActionChains(driver).send_keys(Keys.ENTER).perform()

            else:
                # Anexar Mídia
                time.sleep(1)
                attach_button = driver.find_element("xpath", "//div[@title='Anexar']")
                attach_button.click()
                # Localize o botão de upload de arquivo
                upload_input = driver.find_element("xpath", "//input[@accept='image/*,video/mp4,video/3gpp,video/quicktime']")

                # Envie o caminho do arquivo
                upload_input.send_keys(media_path)

                # Aguarde o upload da mídia
                time.sleep(2)

                send_box = driver.find_element("xpath", "//div[@class='to2l77zo gfz4du6o ag5g9lrv fe5nidar kao4egtt']")

                # Mensagem para o contato
                send_term = contato["mensagem"]

                if not (pd.isna(send_term)):
                    ActionChains(driver).send_keys_to_element(send_box, send_term).perform()

                time.sleep(2)

                # Pressione Enter para enviar a mensagem
                ActionChains(driver).send_keys(Keys.ENTER).perform()

                # Aguarde um tempo antes de prosseguir para o próximo contato
                time.sleep(2)
            
    except NoSuchElementException:
        print("Contato {} não encontrado. Pulando para o próximo contato.".format(search_term))
    finally:
        # Limpe o campo de busca antes de prosseguir para o próximo contato
        ActionChains(driver).send_keys_to_element(search_box, Keys.BACKSPACE * len(search_term)).perform()

time.sleep(5)
driver.quit()