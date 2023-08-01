import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException

# Carregar os dados da planilha
df = pd.read_excel('C:\\Users\\Usuario\\Documents\\otimizacao_web\\planilha\\challenge.xlsx')
df.columns = df.columns.str.strip()  # Remove espaços extras nos nomes das colunas

# Configurar o driver do Selenium
chrome_options = Options()
# chrome_options.add_argument("--headless")  # Executar em modo headless (sem janela)
driver = webdriver.Chrome(options=chrome_options)

# Acessar o formulário
driver.get('https://rpachallenge.com/?lang=EN')

# Função para localizar o campo pelo texto do rótulo
def find_field_by_label(label_text):
    label_elements = driver.find_elements(By.XPATH, "//label[contains(text(), '{}')]".format(label_text))
    for label_element in label_elements:
        field_element = label_element.find_element(By.XPATH, "./following-sibling::input")
        if field_element.is_displayed():
            return field_element
    return None

# Função para preencher o campo
def fill_field(field_element, value):
    field_element.clear()
    field_element.send_keys(value)

# Função para localizar o botão de envio
def find_submit_button():
    submit_button = driver.find_element(By.XPATH, "//input[@type='submit' and @value='Submit']")
    if submit_button.is_displayed():
        return submit_button
    return None

# Função para localizar o botão de start
def find_start_button():
    start_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Start')]")
    if start_button.is_displayed():
        return start_button
    return None

# Automatizar o preenchimento
start_button = find_start_button()
if start_button:
    start_button.click()  # Iniciar o desafio

for round_number in range(1, 11):
    for index, row in df.iterrows():
        first_name = row['First Name']
        last_name = row['Last Name']
        company_name = row['Company Name']
        role = row['Role in Company']
        address = row['Address']
        email = row['Email']
        phone = row['Phone Number']

        # Preencher os campos de acordo com o rótulo
        field_element = find_field_by_label('First Name')
        if field_element:
            fill_field(field_element, first_name)

        field_element = find_field_by_label('Last Name')
        if field_element:
            fill_field(field_element, last_name)

        field_element = find_field_by_label('Company Name')
        if field_element:
            fill_field(field_element, company_name)

        field_element = find_field_by_label('Role in Company')
        if field_element:
            fill_field(field_element, role)

        field_element = find_field_by_label('Address')
        if field_element:
            fill_field(field_element, address)

        field_element = find_field_by_label('Email')
        if field_element:
            fill_field(field_element, email)

        field_element = find_field_by_label('Phone Number')
        if field_element:
            fill_field(field_element, phone)

        # Submeter o formulário
        submit_button = find_submit_button()
        if submit_button:
            submit_button.click()

        # Esperar a próxima página carregar
        time.sleep(5)

        print("Envio concluído corretamente.")

    if round_number < 10:
        try:
            # Clicar no botão "Next Round" para avançar para o próximo round
            next_round_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'Next Round')]")))
            next_round_button.click()
        except (NoSuchElementException, TimeoutException):
            break

# Fechar o driver do Selenium
driver.quit()

print("Código executado corretamente.")
