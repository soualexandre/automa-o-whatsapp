import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time

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

# Loop através dos contatos
for index, contato in tabela.iterrows():
    search_box = driver.find_element("xpath", "//div[@class='g0rxnol2 ln8gz9je lexical-rich-text-input']")

    # Use o número de telefone para pesquisar o contato
    search_term = str(contato['numero'])

    # Simule a digitação como se fosse feita pelo usuário
    ActionChains(driver).send_keys_to_element(search_box, search_term).perform()

    # Aguarde alguns segundos para que a mensagem seja enviada
    time.sleep(2)
    ActionChains(driver).send_keys(Keys.ENTER).perform()
    time.sleep(1)

    # Adicione o caminho do arquivo de mídia desejado
    media_path = contato["caminho_arquivo"]

    if pd.isna(media_path):
        send_box = driver.find_element("xpath", "//div[@class='to2l77zo gfz4du6o ag5g9lrv bze30y65 kao4egtt']")

        # Mensagem para o contato
        send_term = "Mensagem para {}".format(contato['nome'])

        ActionChains(driver).send_keys_to_element(send_box, send_term).perform()

        time.sleep(2)

        # Pressione Enter para enviar a mensagem
        ActionChains(driver).send_keys(Keys.ENTER).perform()

        # Aguarde um tempo antes de prosseguir para o próximo contato
        time.sleep(2)
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
        send_term = "Mensagem para {}".format(contato['nome'])

        ActionChains(driver).send_keys_to_element(send_box, send_term).perform()

        time.sleep(2)

        # Pressione Enter para enviar a mensagem
        ActionChains(driver).send_keys(Keys.ENTER).perform()

        # Aguarde um tempo antes de prosseguir para o próximo contato
        time.sleep(2)

# Feche o navegador após enviar para todos os contatos
time.sleep(2000000)
    
driver.quit()