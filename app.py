import secrets
import time
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options

secret = secrets.token_urlsafe(23)

path = "C:\Program Files (x86)\cdriver\chromedriver"
service = Service(executable_path=path)
options = Options()
driver = webdriver.Chrome(service=service, options=options)

driver.get("https://ladder.8b.africa/accounts/signup/?ref=lll")
driver.maximize_window()

first_name = driver.find_element(By.ID, "id_first_name")
last_name = driver.find_element(By.ID, "id_last_name")
email = driver.find_element(By.ID, "id_email")
phone = driver.find_element(By.ID, "id_phone_number_1")
pass1 = driver.find_element(By.ID, "id_password1")
pass2 = driver.find_element(By.ID, "id_password2")
sign_up = driver.find_element(By.XPATH, '//div[@class="mb-3 d-flex justify-content-end"]/button[@class="btn btn-main-rounded"]')
selectname = Select(driver.find_element(By.NAME, "phone_number_0"))


actions = ActionChains(driver)


wb = load_workbook('testing.xlsx')
ws= wb.active
rows = ws.rows
headers_value = [cell.value for cell in next(rows)] # is a generator yields so needs next
all_rows = []
for row in rows:
    data = {}
    for title,cell in zip(headers_value, row):
        data[title] = cell.value

    all_rows.append(data)

for row in all_rows:
    first_name.send_keys(row["First_name"])
    last_name.send_keys(row["Last_name"])
    email.send_keys(row["email"])
    selectname.select_by_visible_text("Kenya +254")
    phone.send_keys(row["number"])
    pass1.send_keys(secret)
    pass2.send_keys(secret)
    driver.implicitly_wait(10)

    elem = driver.find_element(By.XPATH, '//div[@class="mb-3 d-flex justify-content-end"]/button[@class="btn btn-main-rounded"]' )
    driver.execute_script('arguments[0].click()', elem)

    driver.implicitly_wait(10)
    selectdate = Select(driver.find_element(By.NAME, "date_of_birth_day"))
    selectmonth = Select(driver.find_element(By.NAME, "date_of_birth_month"))
    selectyear = Select(driver.find_element(By.NAME, "date_of_birth_year"))
    selectcountry = Select(driver.find_element(By.NAME, "country_of_birth"))

    selectdate.select_by_visible_text("12")
    driver.implicitly_wait(5)
    selectmonth.select_by_visible_text("12")
    driver.implicitly_wait(5)
    selectyear.select_by_visible_text("1995")
    driver.implicitly_wait(5)
    selectcountry.select_by_visible_text("Kenya")
    driver.implicitly_wait(5)
    driver.back()







