import time
import csv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains

# === Подключение к уже запущенному Chrome ===
chrome_options = Options()
chrome_options.debugger_address = "127.0.0.1:9222"
service = Service()
driver = webdriver.Chrome(service=service, options=chrome_options)

# === Читаем CSV, берём только числовые строки ===
counters = {}
with open("data.csv", newline="", encoding="utf-8") as f:
    reader = csv.reader(f)
    for row in reader:
        if not row:
            continue
        try:
            counter = str(int(row[0]))
            value = row[1]
            counters[counter] = value
        except ValueError:
            continue

print(f"Загружено {len(counters)} счётчиков из CSV")

# === Перебираем строки таблицы ===
rows = driver.find_elements(By.CSS_SELECTOR, "tr.controls-DataGridView__tr")

# Нейтральное место для сброса focus
neutral_click = driver.find_element(By.CSS_SELECTOR, "thead.controls-DataGridView__thead")

for tr in rows:
    try:
        tds = tr.find_elements(By.CSS_SELECTOR, "td.controls-DataGridView__td")
        if len(tds) < 2:
            continue

        num_el = tds[1].find_element(By.CSS_SELECTOR, "div.fed-CellContent")
        number = num_el.text.strip()
        if not number:
            continue

        normalized = str(int(number))

        if normalized in counters:
            value = counters[normalized]
            print(f"[MATCH] {normalized} → {value}")

            # ищем ячейку для ввода
            input_cell = None
            for td in tds:
                try:
                    div = td.find_element(
                        By.CSS_SELECTOR,
                        'div.fed-CellContent[name="ТекущиеПоказания.Показание"]'
                    )
                    input_cell = div
                    break
                except:
                    continue

            if not input_cell:
                print(f"[WARN] не нашёл поле ввода для {normalized}")
                continue

            # Прокрутка к элементу и клик
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", input_cell)
            time.sleep(0.1)
            input_cell.click()
            time.sleep(0.1)

            # ждём появления активного input
            active_input = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "input.controls-TextBox__field"))
            )

            # имитация ввода
            actions = ActionChains(driver)
            for ch in value:
                actions.send_keys(ch)
                actions.pause(0.1)
            actions.send_keys(Keys.ENTER)
            actions.perform()

            # Сбрасываем focus, чтобы следующий клик прошёл
            neutral_click.click()
            time.sleep(0.1)

            print(f"[OK] Заполнил {normalized} значением {value}")
            time.sleep(0.2)

    except Exception as e:
        print(f"[ERR] строка пропущена: {e}")
        continue

print("Готово ✅")
