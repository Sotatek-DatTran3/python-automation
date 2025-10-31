from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
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
        EC.presence_of_element_located((By.XPATH, "//input[@type='email' and @placeholder='Your email address']"))
    )
    password_input_el = driver.find_element(By.XPATH, "//input[@type='password' and @placeholder='Your password']")

    email_input_el.clear()
    email_input_el.send_keys(os.environ.get("log_mail", ""))
    password_input_el.clear()
    password_input_el.send_keys(os.environ.get("log_pass", ""))

    driver.find_element(By.XPATH, "//button[.//span[normalize-space(text())='Login']]").click()

    # ===== Query =====
    for query in queries:
        query_input_el = wait.until(
            EC.presence_of_element_located((By.XPATH, "//textarea[@placeholder='Ask anything...']"))
        )
        query_input_el.click()
        query_input_el.send_keys(query)

        driver.find_element(By.XPATH, "//button[.//img[@alt='icon-send']]").click()
        time.sleep(3)

        stable_count = 0
        last_texts = []

        while stable_count < 3:
            try: 
                answer_el = driver.find_elements(
                    By.XPATH,
                    "//div[contains(@class, 'prose-img')]"
                )

                responses = [el.text for el in answer_el]

                # Check if content is stable
                if responses == last_texts:
                    stable_count += 1
                else:
                    stable_count = 0
                    last_texts = responses

            except StaleElementReferenceException:
                time.sleep(1)
                continue
            time.sleep(1)

        time.sleep(0.5)

        document_refs = driver.find_elements(By.XPATH, "//div[contains(@class, 'text-[#6E6F7C')]/following-sibling::div[contains(@class, 'border')]//span")

        responses = [r for r in responses if r.strip()]
        full_response = "\n".join(responses)

        file_names = [doc_ref.text for doc_ref in document_refs if doc_ref.text.strip() != ""]
        result.append({
            "answer": full_response,
            "references": file_names
        })

        print("\n---")
        print("Completed query:", query)
        print(f"Response {full_response[:100]}...")
        print("References:", file_names, "\n")

        driver.find_element(By.XPATH, "//button[.//img[@alt='icon-add']]").click()

    return result, driver