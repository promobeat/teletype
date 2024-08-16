import os
import time
import random
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import pyperclip
import logging
from transliterate import translit

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Конфигурация
TELETYPE_LOGIN_URL = "https://teletype.in/login"
TELETYPE_NEW_POST_URL = "https://teletype.in/@testpython/editor"
USERNAME = "test@test.tu"
PASSWORD = ".8ZFu)sJW4DD6SY"
POSTING_FOLDER = "posting"
RESULT_FILE = "posting_teletype.xlsx"
CHROME_DRIVER_PATH = r"C:\WebDrivers_Selenium\chromedriver.exe"

def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-extensions')
    options.add_argument('--disable-software-rasterizer')
    options.add_argument('--use-gl=desktop')
    options.add_argument('--use-angle=gl-angle')
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    service = Service(CHROME_DRIVER_PATH)
    return webdriver.Chrome(service=service, options=options)

def random_sleep(min_seconds=1, max_seconds=5):
    time.sleep(random.uniform(min_seconds, max_seconds))

def login(driver):
    driver.get(TELETYPE_LOGIN_URL)
    try:
        logging.info("Попытка входа в систему...")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "email")))
        
        login_element = driver.find_element(By.NAME, "email")
        login_element.clear()
        login_element.send_keys(USERNAME)
        logging.info("Введен логин")
        random_sleep()
        
        password_element = driver.find_element(By.NAME, "password")
        password_element.clear()
        password_element.send_keys(PASSWORD)
        logging.info("Введен пароль")
        random_sleep()
        
        submit_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".button.m_type_filled.m_border_rounded"))
        )
        submit_button.click()
        logging.info("Кнопка отправки формы нажата")
        
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="app"]/div[1]/div[1]/div[2]/div[3]/div[1]/div/a'))
        )
        logging.info("Успешный вход в систему")
        return True
    except Exception as e:
        logging.error(f"Ошибка при входе: {str(e)}")
        return False

def open_editor(driver):
    try:
        driver.get(TELETYPE_NEW_POST_URL)
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".editor.m_line.m_empty"))
        )
        logging.info("Открыт редактор для новой статьи")
        return True
    except Exception as e:
        logging.error(f"Ошибка при открытии редактора: {str(e)}")
        return False

def read_html_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        return file.read()

def extract_content(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    h1_tag = soup.find('h1')
    h1_text = h1_tag.text.strip() if h1_tag else "Без заголовка"
    
    body_content = soup.body
    if body_content:
        if h1_tag:
            h1_tag.decompose()
        formatted_content = ''.join(str(tag) for tag in body_content.contents if tag.name != 'script')
    else:
        formatted_content = ""
    
    links = soup.find_all('a')
    anchor = links[0].text if links else ""
    external_link = links[0]['href'] if links else ""
    
    logging.info(f"Извлеченный заголовок: {h1_text}")
    logging.info(f"Извлеченный анкор: {anchor}")
    logging.info(f"Извлеченная внешняя ссылка: {external_link}")
    
    return h1_text, formatted_content, anchor, external_link

def transliterate_filename(filename):
    name_without_extension = os.path.splitext(filename)[0]
    transliterated = translit(name_without_extension, 'ru', reversed=True)
    cleaned = ''.join(c if c.isalnum() or c == '-' else '' for c in transliterated.replace(' ', '-'))
    return cleaned.lower()

def post_to_teletype(driver, h1_text, formatted_content, filename):
    try:
        h1_field = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".editor.m_line.m_empty"))
        )
        h1_field.send_keys(h1_text)
        h1_field.send_keys(Keys.RETURN)
        
        time.sleep(3)
        
        content_field = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".editorPage__text.text.editor.m_block"))
        )
        
        content_field.send_keys(Keys.HOME)
        
        pyperclip.copy(formatted_content)
        
        ActionChains(driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
        time.sleep(2)
        
        random_sleep(1, 3)

        publish_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div[1]/div[1]/div[2]/div[3]/button[2]'))
        )
        driver.execute_script("arguments[0].click();", publish_button)
        logging.info("Нажата кнопка 'Опубликовать'")

        filename_field = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="app"]/div[1]/div[3]/div/div/div[1]/div[4]/div[3]/div/input'))
        )

        filename_field.clear()
        filename_field.send_keys(Keys.CONTROL + "a")
        filename_field.send_keys(Keys.DELETE)
        
        transliterated_filename = transliterate_filename(filename)
        filename_field.send_keys(transliterated_filename)
        logging.info(f"Вставлено транслитерированное имя файла: {transliterated_filename}")

        time.sleep(3)

        try:
            final_publish_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, ".editorPublisher__submit"))
            )
            final_publish_button.click()
        except:
            try:
                final_publish_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Опубликовать')]"))
                )
                final_publish_button.click()
            except:
                driver.execute_script("""
                    var buttons = document.getElementsByTagName('button');
                    for(var i = 0; i < buttons.length; i++) {
                        if(buttons[i].textContent.includes('Опубликовать')) {
                            buttons[i].click();
                            break;
                        }
                    }
                """)

        logging.info("Нажата кнопка окончательной публикации")
        
        try:
            WebDriverWait(driver, 60).until(EC.url_contains(transliterated_filename))
            current_url = driver.current_url
            logging.info(f"Статья успешно опубликована. URL: {current_url}")
            return current_url
        except TimeoutException:
            logging.error(f"Превышено время ожидания публикации статьи. Текущий URL: {driver.current_url}")
        except Exception as e:
            logging.error(f"Ошибка при ожидании публикации статьи: {str(e)}")
        
        return None
    except Exception as e:
        logging.error(f"Ошибка при публикации статьи: {str(e)}")
        return None

def check_for_captcha_or_popup(driver):
    try:
        captcha = driver.find_elements(By.XPATH, "//div[contains(@class, 'captcha')]")
        if captcha:
            logging.warning("Обнаружена капча. Требуется ручное вмешательство.")
            input("Нажмите Enter после решения капчи...")
            return True

        popup = driver.find_elements(By.XPATH, "//div[contains(@class, 'popup') or contains(@class, 'modal')]")
        if popup:
            logging.warning("Обнаружено всплывающее окно. Попытка закрыть...")
            close_button = driver.find_element(By.XPATH, "//button[contains(@class, 'close') or contains(@class, 'dismiss')]")
            close_button.click()
            return True

    except Exception as e:
        logging.error(f"Ошибка при проверке капчи или всплывающего окна: {str(e)}")
    
    return False

def logout(driver):
    try:
        user_menu = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".userMenu"))
        )
        user_menu.click()

        logout_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Выйти')]"))
        )
        logout_button.click()

        logging.info("Успешный выход из системы")
        return True
    except Exception as e:
        logging.error(f"Ошибка при выходе из системы: {str(e)}")
        return False

def process_single_file():
    driver = None
    try:
        driver = setup_driver()
        if login(driver):
            html_files = [f for f in os.listdir(POSTING_FOLDER) if f.endswith('.html')]
            
            if not html_files:
                logging.info("В папке posting нет HTML файлов для публикации.")
                return False

            filename = html_files[0]
            file_path = os.path.join(POSTING_FOLDER, filename)
            html_content = read_html_file(file_path)
            h1_text, formatted_content, anchor, external_link = extract_content(html_content)
            
            random_sleep(5, 10)
            
            if open_editor(driver):
                if check_for_captcha_or_popup(driver):
                    return False
                url = post_to_teletype(driver, h1_text, formatted_content, filename)
                if url:
                    logging.info(f"Опубликовано: {filename} -> {url}")
                    
                    try:
                        df = pd.DataFrame([[h1_text, url, anchor, external_link]], 
                                          columns=['Название статьи', 'Ссылка на статью', 'Анкор ссылки', 'Внешняя ссылка'])
                        
                        if os.path.exists(RESULT_FILE):
                            existing_df = pd.read_excel(RESULT_FILE)
                            df = pd.concat([existing_df, df], ignore_index=True)
                        
                        df.to_excel(RESULT_FILE, index=False)
                        logging.info(f"Результат сохранен в {RESULT_FILE}")
                    except Exception as e:
                        logging.error(f"Ошибка при сохранении результатов в Excel: {str(e)}")
                    
                    try:
                        os.remove(file_path)
                        logging.info(f"Файл {filename} удален из папки posting")
                    except Exception as e:
                        logging.error(f"Ошибка при удалении файла {filename}: {str(e)}")
                    
                    return True
                else:
                    logging.warning(f"Не удалось опубликовать: {filename}")
            else:
                logging.warning(f"Не удалось открыть редактор для {filename}")
        else:
            logging.error("Не удалось войти в систему. Проверьте логин и пароль.")
    except Exception as e:
        logging.error(f"Произошла неожиданная ошибка: {str(e)}")
    finally:
        if driver:
            try:
                logout(driver)
            except:
                logging.warning("Не удалось выполнить выход из системы")
            driver.quit()
    
    return False

def main():
    if not os.path.exists(POSTING_FOLDER):
        logging.error(f"Папка {POSTING_FOLDER} не существует. Создайте папку и поместите в нее HTML файлы для публикации.")
        return

    while True:
        html_files = [f for f in os.listdir(POSTING_FOLDER) if f.endswith('.html')]
        if not html_files:
            logging.info("Все файлы обработаны. Завершение работы.")
            break
        
        success = process_single_file()
        if not success:
            logging.warning("Не удалось обработать файл. Попытка повторного запуска через 60 секунд.")
            time.sleep(60)
        else:
            logging.info("Файл успешно обработан. Переход к следующему файлу.")
            time.sleep(10)  # Небольшая пауза перед следующим запуском

if __name__ == "__main__":
    main()