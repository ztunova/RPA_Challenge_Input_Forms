import openpyxl
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By


# check if actual row of xlsx file contains data
def check_row(row_to_check):
    #print(row_to_check[0].value)
    if row_to_check[0].value is None:
        return False
    return True


# load xlsx worksheet and create map with indexes of labels - which atribute occurs in which column
workbook = openpyxl.load_workbook("challenge.xlsx")
worksheet = workbook.active
label_index_map = {}
for i in range(len(worksheet[1])):
    label_index_map[str(worksheet[1][i].value).lower().replace(" ", "")] = i

s = Service('C:\Program Files (x86)\chromedriver.exe')
driver = webdriver.Chrome(service = s)
driver.get("https://rpachallenge.com/")

start_button = driver.find_element(By.CLASS_NAME, "waves-effect.uiColorButton")
start_button.click()

row = 2
# while there are data to be submitted, identify all fields for data and submit button
while check_row(worksheet[row]):
    form_fields = driver.find_elements(By.TAG_NAME, "rpa1-field")
    form = driver.find_element(By.TAG_NAME, 'form')
    submit_button = form.find_element(By.CLASS_NAME, 'btn')

    # identify label for individual input fields
    # get column (index) in which these data are stored in worksheet and fill in
    for field in form_fields:
        form_text = field.text.lower().replace(" ", "")
        input_field = field.find_element(By.TAG_NAME, 'input')
        column_idx = label_index_map.get(form_text)
        input_field.send_keys(worksheet[row][column_idx].value)

    submit_button.click()
    row = row + 1

time.sleep(5)
driver.quit()
