import re

import pandas as pd
import wait

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC

df = pd.read_excel("./DeliverySlips/DS-Fep-Dev.xlsx")
name = df.iloc[3, 2]
print(name)

phone = df.iloc[6, 2]
print(phone)

address = df.iloc[7, 2]
print(address)

city = df.iloc[8, 2]

list_city = city.split(", ")

city = list_city[0]
state = list_city[1]
zipcode = list_city[2]

print(city)
print(state)
print(zipcode)


signature_req = df.iloc[18,2]
if signature_req.casefold() == "yes":
    signature_req = "Signature Required"
elif signature_req.casefold() == "no":
    signature_req = "No Signature Required"

notes = df.iloc[22, 2]
print(notes)

drug_number = round(df.iloc[24, 2])
print(drug_number)



rx_numbers = df.iloc[26, 2]
print(rx_numbers)
rxNumRegex = re.compile(r'\d+')
rx_numbers = rxNumRegex.findall(rx_numbers)
print(rx_numbers)




# rx_number variable index starts at 0
rx_number = {}
for i in range(len(rx_numbers)):
    rx_number[f'rx_number{i+1}'] = rx_numbers[i]
print(rx_number)

rx_number_list = list(rx_number.values())
print(rx_number_list)

sort_code = df.iloc[39,2]
print(sort_code)

if sort_code =="CH":
    global clinic_email
    clinic_email = "wcmcsschelseaID@nyp.org"
    print(clinic_email)
else:
    print("unknown_clinic_email")
    
    

chrome_driver_path = "/Users/felixpatawah/PycharmProjects/pythonProject/chromedriver"
driver = webdriver.Chrome(executable_path=chrome_driver_path)
driver.get("https://healthex.e-courier.com/healthex/home/wizard_oe2.asp?UserGUID={#######}")

full_name = driver.find_element(By.XPATH, '//*[@id="DName"]')
full_name.send_keys(name)

addy = driver.find_element(By.XPATH, '//*[@id="DAddress"]')
addy.send_keys(address)

healthexcity = driver.find_element(By.XPATH, '//*[@id="DCity"]')
healthexcity.send_keys(city)

healthex_zipcode = driver.find_element(By.XPATH, '//*[@id="DZip"]')
healthex_zipcode.send_keys(zipcode)

healthex_phone = driver.find_element(By.XPATH, '//*[@id="DPhone"]')
healthex_phone.send_keys(phone)

healthex_notes = driver.find_element(By.XPATH, '//*[@id="Textarea1"]')
healthex_notes.send_keys(Keys.RETURN,
                         f'{signature_req}',
                         Keys.RETURN*2,
                         f'{notes}',
                         Keys.RETURN*2,
                        )

healthex_pieces = driver.find_element(By.XPATH, '//*[@id="Pieces"]')
healthex_pieces.send_keys(drug_number, Keys.TAB)


# Click the link which opens in a new window
driver.find_element(By.XPATH, '//*[@id="Pieces"]').click()


driver.switch_to.window(driver.window_handles[1])
driver.implicitly_wait(5)
driver.maximize_window()

for i in range(0, drug_number):
    rxnumber_field = driver.find_element(By.ID, f"Reference{i+1}")
    rxnumber_field.click()
    rxnumber_field.send_keys(rx_number_list[i])
driver.find_element(By.NAME, "Submit").click()

driver.implicitly_wait(5)
driver.switch_to.window(driver.window_handles[0])

driver.find_element(By.XPATH, '//*[@id="frmInput"]/table[2]/tbody/tr[23]/td[2]/table/tbody/tr[4]/td[1]/input').click()
driver.find_element(By.XPATH, '//*[@id="frmInput"]/table[2]/tbody/tr[23]/td[2]/table/tbody/tr[4]/td[1]/input').send_keys(clinic_email)
driver.find_element(By.NAME, "Notify3rdPartyEmail").click()
driver.find_element(By.NAME, "Notify3rdPartyEmail").send_keys(clinic_email)
driver.switch_to.window(driver.window_handles[0])
driver.find_element(By.ID, "btnQuote").click()

