import urllib.parse
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
import random
import datetime as dt

# Carregar os contatos do arquivo Excel
contacts_df = pd.read_excel('test1.xlsx')

# Configurar o WebDriver do Chrome
browser = webdriver.Chrome()

# Abrir o WhatsApp Web
browser.get('https://web.whatsapp.com/')

# Esperar o WhatsApp Web carregar completamente
while len(browser.find_elements(By.ID, 'side')) < 1:
    time.sleep(4)

# Enviar mensagens
for i, message in enumerate(contacts_df['Message']):
    name = contacts_df.loc[i, 'Name']
    phone = contacts_df.loc[i, 'phone']
    text = urllib.parse.quote (message)
    link = f'https://web.whatsapp.com/send?phone={phone}&text={text}'
    hora_atual = dt.datetime.now().strftime("%H:%M:%S")
    linha = ("linha nº")
    print(f"{hora_atual} {name} {linha} {i+2}")

    # Abrir o link com a mensagem
    browser.get(link)
    
    # Esperar a página carregar completamente
    while len(browser.find_elements(By.ID, 'side')) < 1:
        time.sleep(5)
    
    # Aguardar a interface de envio de mensagem carregar
    ##Função para criar um atraso aleatório entre 8 e 12 segundos
    def pausa_aleatoria():
        tempo_aleatorio = random.uniform(8, 12)
        time.sleep(tempo_aleatorio)
        print("    ",tempo_aleatorio)
    pausa_aleatoria()
    
    # Enviar a mensagem
    try:
        # Localizar o campo de mensagem e enviar a mensagem
        message_box = browser.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/div[2]/div[1]/p')
        message_box.send_keys(Keys.RETURN)
    except Exception as e:
        print(f"Erro ao enviar mensagem para {name}: {e}")

    ## Esperar alguns segundos antes de enviar a próxima mensagem
    def pausa_aleatoria1():
        tempo_aleatorio = random.uniform(15, 30)
        time.sleep(tempo_aleatorio)
        print("    ",tempo_aleatorio)
    pausa_aleatoria1()