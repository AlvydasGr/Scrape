import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import openpyxl
import pandas as pd
import os
from selenium.webdriver.common.action_chains import ActionChains
import warnings


Skaitliukas = 0

warnings.filterwarnings('ignore')

current_dir = os.getcwd()
new_input = f'{current_dir}\input.xlsx'
screenshots_dir = f'{current_dir}\\screenshots'
workbook = openpyxl.load_workbook(new_input)
sh=workbook.active
excel = pd.read_excel(new_input, engine='openpyxl', sheet_name="Sheet1")


chrome_options = Options()

#prefs = {"profile.managed_default_content_settings.images": 2} #Seleniumas nebekrauna image
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"]) #Nerodo cmd ka veikia browseris
#chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_argument("--headless") #Vizualiai neijungia chromo
chrome_options.add_argument('--window-size=1920,2200')
chrome_options.add_argument('disable-blink-features=AutomationControlled')
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.106 Safari/537.36")
chrome_options.add_argument("--force-device-scale-factor=0.6")
driver = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)
driver.execute_cdp_cmd('Network.setUserAgentOverride', {"userAgent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.97 Safari/537.36'})
driver.get('https://post.lt/siuntu-sekimas')

atsakymas = {"ID": [],
             "status": [],
             "Numeris": []}

actions = ActionChains(driver)

for index, row in enumerate(sh):
    if index == 0:
        continue
    ID = excel['ID'][Skaitliukas]
    Numeris = excel['Numeriai'][Skaitliukas]
    print(f'Tikrinamas numeris: {Numeris}')
    driver.find_element(By.ID, "parcelsInput").send_keys(Numeris)
    driver.find_element(By.XPATH, "//button[@class='btn btn-black btn-lg']").click()
    time.sleep(4)
    status = driver.find_elements(By.XPATH, "//div[@class='table-wrap']/table[@class='table table-bordered table-shipment border-collapse']")
    for i in status:
        if "Siunta įteikta gavėjui " in i.text:
            atsakymas['ID'].append(ID)
            atsakymas['status'].append("Siunta iteikta")
            atsakymas['Numeris'].append(Numeris)
            move = driver.find_element(By.XPATH, "//div[@class='table-notice']/p[@class='f-xs text-uppercase']")
            actions.move_to_element(move)
            time.sleep(1)
            os.chdir(screenshots_dir)
            driver.save_screenshot(f'{Numeris}.png')
        elif "Siunta grąžinta siuntėjui " in i.text:
            atsakymas['ID'].append(ID)
            atsakymas['status'].append("Siunta grazinta")
            atsakymas['Numeris'].append(Numeris)
        else:
            atsakymas['ID'].append(ID)
            atsakymas['status'].append("Siunta tranzite")
            atsakymas['Numeris'].append(Numeris)
    Skaitliukas = Skaitliukas + 1
    driver.find_element(By.ID, "parcelsInput").clear()
    time.sleep(5)


df = pd.DataFrame(atsakymas)

path = f'{current_dir}\\output.xlsx'

with warnings.catch_warnings(record=True):
    warnings.simplefilter('always')
    with pd.ExcelWriter(path, engine='openpyxl', if_sheet_exists="replace", mode='a') as writer:
        df.to_excel(writer, sheet_name="Sheet1", index=False)

print("-------")
print("PABAIGA")
print("-------")

os.system('taskkill /f /im chromedriver.exe')
