from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
import sys
import os
from pathlib import Path
from dotenv import load_dotenv
import time

def load_environment():
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    else:
        application_path = Path(__file__).parent.parent.parent

    env_path = os.path.join(application_path, ".env")

    if not os.path.exists(env_path):
        raise FileNotFoundError(".env file not found")

    load_dotenv(env_path)
    print(f"Loaded .env from {env_path}")
    return env_path

def crawl_answer(queries: list[str]):
    load_environment()

    result = []

    options = webdriver.EdgeOptions()
    options.add_experimental_option("detach", True)
    driver = webdriver.Edge(options=options)

    driver.get(os.environ.get("host"))

    wait = WebDriverWait(driver, 40)

    # ===== Log in =====
    email_input_el = wait.until(
        EC.presence_of_element_located((By.XPATH, "//input[@type='email' and @placeholder='あなたのメールアドレス']"))
    )
    password_input_el = driver.find_element(By.XPATH, "//input[@type='password' and @placeholder='あなたのパスワード']")

    email_input_el.clear()
    email_input_el.send_keys(os.environ.get("log_mail", ""))
    password_input_el.clear()
    password_input_el.send_keys(os.environ.get("log_pass", ""))

    driver.find_element(By.XPATH, "//button[.//span[normalize-space(text())='ログイン']]").click()

    # ===== Query =====
    for query in queries:
        query_input_el = wait.until(
            EC.presence_of_element_located((By.XPATH, "//textarea[@placeholder='何でも質問してください...']"))
        )
        query_input_el.click()
        query_input_el.send_keys(query)

        driver.find_element(By.XPATH, "//button[.//img[@alt='icon-send']]").click()
        time.sleep(3)

        answer_el = driver.find_element(By.XPATH, "//div[contains(@class, 'justify-start')]//div[contains(@class, 'prose-img:my-0')]")
        last_text = ""
        stable_count = 0

        while stable_count < 3:
            current_text = answer_el.text
            if current_text == last_text:
                stable_count += 1
            else:
                stable_count = 0
                last_text = current_text
            time.sleep(1)

        document_refs = driver.find_elements(By.XPATH, "//div[contains(text(), '情報源')]/following-sibling::div//span[contains(@class, 'text-[#27283C]')]")

        file_names = [doc_ref.text for doc_ref in document_refs if doc_ref.text.strip() != ""]
        result.append({
            "answer": answer_el.text,
            "references": file_names
        })

        driver.refresh()

    time.sleep(2)        
    driver.quit()

    return result