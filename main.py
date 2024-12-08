from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import WebDriverException, NoSuchElementException, ElementNotInteractableException
import pandas as pd
import time
import os
import smtplib
from email.message import EmailMessage
import mimetypes

options = Options()
options.add_argument('--ignore-certificate-errors')
options.add_argument('--disable-web-security')
options.add_argument('--allow-running-insecure-content')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--no-sandbox')
options.add_experimental_option('excludeSwitches', ['enable-logging'])

service = Service(r"C:/Users/gusta/chromedriver.exe")
output_dir = r"C:/Users/gusta/Desktop/projeto_robo_magalu/output"

def load_site(url, retries=3):
    while retries > 0:
        try:
            driver = webdriver.Chrome(service=service, options=options)
            driver.get(url)
            time.sleep(10)
            return driver
        except WebDriverException as e:
            print(f"Erro ao carregar o site: {e}")
            retries -= 1
            if retries == 0:
                with open("error_log.txt", "w") as log:
                    log.write("Site fora do ar")
                return None

def search_product(driver, search_query):
    search_box = driver.find_element(By.CSS_SELECTOR, "input[type='search']")
    search_box.send_keys(search_query)
    search_box.send_keys(Keys.RETURN)
    time.sleep(6)

    all_data = []
    page_count = 0

    while page_count < 20:
        time.sleep(4)
        data = extract_data(driver)
        all_data.extend(data)

        try:
            next_button = driver.find_element(By.CSS_SELECTOR, "button[type='next']")
            if "disabled" in next_button.get_attribute("class"):
                break

            ActionChains(driver).move_to_element(next_button).click().perform()
            time.sleep(4)
            page_count += 1
        except (NoSuchElementException, ElementNotInteractableException):
            break

    return all_data

def extract_data(driver):
    try:
        data = []
        products = driver.find_elements(By.XPATH, "//*[@id='__next']/div/main/section[4]/div[5]/div/ul/li/a")

        for product in products:
            try:
                name = product.find_element(By.XPATH, "div[3]/h2").text
                url = product.get_attribute("href")
                
                try:
                    review_text = product.find_element(By.XPATH, "div[3]/div[2]/div/span").text
                    reviews = review_text.split('(')[1].split(')')[0].replace('.', '').replace(',', '')
                    reviews = int(reviews)
                except (NoSuchElementException, IndexError, ValueError):
                    reviews = "ND"
                
                image_url = product.find_element(By.XPATH, "div[2]/img").get_attribute("src")

                data.append([name, reviews, url, image_url])
                print(f"Produto: {name}, Avaliações: {reviews}, URL: {url}, Imagem: {image_url}")
            except NoSuchElementException as e:
                print(f"Erro ao extrair informações de um produto: {e}")
                continue
        return data
    except NoSuchElementException as e:
        print(f"Erro ao extrair dados: {e}")
        return []

def save_to_excel(data):
    df = pd.DataFrame(data, columns=["PRODUTO", "QTD_AVAL", "URL", "IMAGEM"])
    df_with_reviews = df[df["QTD_AVAL"] != "ND"].copy()  # Exclui "Sem Avaliação"
    df_with_reviews["QTD_AVAL"] = pd.to_numeric(df_with_reviews["QTD_AVAL"])

    # Filtra os produtos para cada aba
    melhores = df_with_reviews[df_with_reviews["QTD_AVAL"] >= 100]
    piores = df_with_reviews[df_with_reviews["QTD_AVAL"] < 100]

    # Salva no Excel
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    output_path = os.path.join(output_dir, "Notebooks.xlsx")

    with pd.ExcelWriter(output_path) as writer:
        if not piores.empty:
            piores.to_excel(writer, sheet_name="Piores", index=False)
        if not melhores.empty:
            melhores.to_excel(writer, sheet_name="Melhores", index=False)

    print(f"Planilha salva em: {output_path}")
    return output_path

def send_email(filepath):
    msg = EmailMessage()
    msg['Subject'] = 'Relatório Notebooks'
    msg['From'] = 'gustavoheenri02@gmail.com'  
    msg['To'] = 'gustavoheenri02@gmail.com'
    msg.set_content('''Olá, aqui está o seu relatório dos notebooks extraídos da Magazine Luiza.

Atenciosamente,
Robô''')

    mime_type, _ = mimetypes.guess_type(filepath)
    mime_type, mime_subtype = mime_type.split('/')

    with open(filepath, 'rb') as file:
        msg.add_attachment(file.read(), maintype=mime_type, subtype=mime_subtype, filename=os.path.basename(filepath))

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login('gustavoheenri02@gmail.com', 'Gugu3518#')
        smtp.send_message(msg)
    print("E-mail enviado com sucesso!")

def main():
    url = "https://www.magazineluiza.com.br/"
    driver = load_site(url)
    if not driver:
        return
    
    all_data = search_product(driver, "notebooks")
    driver.quit()
    
    if all_data:
        filepath = save_to_excel(all_data)
        send_email(filepath)
    else:
        print("Nenhum dado extraído.")

if __name__ == "__main__":
    main()