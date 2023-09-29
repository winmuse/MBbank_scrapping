import io, os, time
import json
import pyautogui
# from google.cloud import vision
from datetime import datetime,timedelta
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from winreg import *
from PIL import Image
from io import BytesIO
import requests, json, base64

class AutoDownloadMBbank:
    def __init__(self, user_name, pass_word):
        self.user_name = user_name
        self.pass_word = pass_word

        # self.excel_file_name = "MBbank_Account_Statement.xlsx"
        # with OpenKey(HKEY_CURRENT_USER, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders') as key:
        #     self.dir_download = QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0] + "\\"
        # self.dir_sample_input = os.getcwd() + "\\sample input\\"

        self.driver = webdriver.Chrome()
        self.driver.maximize_window()

        self.runDownload()

    def timeoutToken(self,):
        while True:
            element = '//*[@id="btn-step1"]/button[2]'
    def loadCompleted(self, locator, timeout):
        """ check website load complete """
        try:
            WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((By.XPATH, locator))
            )
            return True
        except TimeoutException:
            return False
    def loadCompletedID(self, locator, timeout):
        """ check website load complete """
        try:
            WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((By.ID, locator))
            )
            return True
        except TimeoutException:
            return False

    def clickElement(self, xpath_element):
        """ find element on website then click """
        try:
            if self.loadCompleted(xpath_element, 50):
                element = self.driver.find_element(By.XPATH, xpath_element)
                element.click()

        except NoSuchElementException:
            print("can not find element:", xpath_element)
        except Exception:
            print("can not click try perform ")
            time.sleep(10)
            # ex_element = WebDriverWait(self.driver, 30).until(
            #     EC.visibility_of_element_located((By.XPATH, xpath_element)))
            ex_element = self.driver.find_element(By.XPATH, xpath_element)
            ActionChains(self.driver).click(ex_element).perform()

    def clickElementID(self, ID_element):
        """ find element on website then click """
        try:
            if self.loadCompletedID(ID_element, 50):
                element = self.driver.find_element(By.ID, ID_element)
                element.click()

        except NoSuchElementException:
            print("can not find element:", ID_element)
        except Exception:
            print("can not click try perform ")
            time.sleep(10)
            ex_element = self.driver.find_element(By.ID, ID_element)
            ActionChains(self.driver).click(ex_element).perform()

    def click_select_date(self, id_btn):
        try:
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.ID, id_btn)))
            down_arrow_btn = self.driver.find_element(By.ID, id_btn)
            down_arrow_btn.click()
            print("click:" + id_btn)
        except Exception:
            print("can't find %s, try run javaScript" % id_btn)

    def Recognition(self):
        try:
            img_auth = self.driver.find_element(By.XPATH, '/html/body/app-root/div/mbb-welcome/div[2]/div[1]/div/div/mbb-login/form/div/div[5]/mbb-word-captcha/div/div[2]/div[1]/div[1]/img')
            location = img_auth.location
            size = img_auth.size
            
            img = self.driver.get_screenshot_as_png()

            left = location['x']
            top = location['y']
            right = location['x'] + size['width']
            bottom = location['y'] + size['height']

            im = Image.open(BytesIO(img))
            im = im.crop((left, top, right, bottom)) # defines crop points
            im.save('screenshot.png')

            df = open('screenshot.png', 'rb').read()

            imgStr = base64.b64encode(df).decode()

            url = 'https://api.1stcaptcha.com/Recognition'
            data = {
                "Apikey": "df96cfbaada3488a977aa88df5c11b8d",
                "Type": "imagetotext",
                "Image": imgStr,
            }
            header = {
                "Content-Type": "application/json"
            }
            req = requests.post(url, data=json.dumps(data), headers=header)
            req_json = req.json()
            print(req_json)

            # text = json.loads(req_json)
            print(req_json['TaskId'])

            self.getresult(str(req_json['TaskId']))
        except:
            print("error")

    def getresult(self, orderId):
        url = 'https://api.1stcaptcha.com/getresult?apikey=df96cfbaada3488a977aa88df5c11b8d&taskid='+orderId
        get_json = requests.get(url).json()
        print(get_json['Data'])

        image_auth = self.driver.find_element(By.XPATH,'/html/body/app-root/div/mbb-welcome/div[2]/div[1]/div/div/mbb-login/form/div/div[5]/mbb-word-captcha/div/div[2]/div[1]/div[2]/input')
        image_auth.clear()
        image_auth.send_keys(get_json['Data'])
    # Login to MBbank
    def loginMBbank(self):
        try:
            self.driver.get("https://online.mbbank.com.vn/pl/login?returnUrl=%2F")
            print("get success")
            user = self.driver.find_element(By.ID, 'user-id')
            password = self.driver.find_element(By.ID, 'new-password')
            user.send_keys(self.user_name)
            password.send_keys(self.pass_word)
            # '/html/body/app-root/div/mbb-welcome/div[2]/div[1]/div/div/mbb-login/form/div/div[5]/mbb-word-captcha/div/div[2]/div[1]/div[1]/img'
            # '/html/body/app-root/div/mbb-welcome/div[2]/div[1]/div/div/mbb-login/form/div/div[5]/mbb-word-captcha/div/div[2]/div[1]/div[2]/input'
            # '/html/body/app-root/div/mbb-welcome/div[2]/div[1]/div/div/mbb-login/form/div/div[6]/div/button'
            self.Recognition()
            
            # if self.loadCompleted('/html/body/app-root/div/mbb-welcome/div[2]/div[1]/div/div/mbb-login/form/div/div[6]/div/button',20):
            #     self.clickElement('/html/body/app-root/div/mbb-welcome/div[2]/div[1]/div/div/mbb-login/form/div/div[6]/div/button')

            # if self.loadCompleted('/html/body/ngb-modal-window/div/div/ng-component/div/div[3]/div/div',20):
            #     self.loginMBbank()
            
            first_login_element = 'MNU_GCME_040000'
            if self.loadCompletedID(first_login_element,2000):
                self.clickElementID(first_login_element)

            print("login success")
        except TimeoutException:
            print("Login MBbank timeout")
            time.sleep(5)
            return
        except:
            time.sleep(10)
            print("has been login MBbank - can't find element")

    def runDownload(self):
        """ start download MBbank Transaction """
        self.loginMBbank()

        time.sleep(10)
        login_time = datetime.now()
        login_time_history = login_time.strftime("%#d/%#m/%Y/%Hh/%Mm/%Ss")
        with open('login_history.txt', mode='w', encoding='utf-8') as log_file:
            log_file.write(login_time_history + '\n')
        if self.loadCompletedID('MNU_GCME_040001',100):
            self.clickElementID('MNU_GCME_040001')

        # input end and start_date
        WebDriverWait(self.driver, 500).until(
            EC.visibility_of_element_located((By.CLASS_NAME, 'chartjs-render-monitor'))
        )
        print("aaaaaaaaa")

        if self.loadCompleted('/html/body/app-root/div/ng-component/div[1]/nav/mbb-navbar/div[2]/perfect-scrollbar/div/div[1]/div[1]/mat-tree/mat-nested-tree-node[1]/li/ul/mat-tree-node[1]/a/li/span',2000):
            self.clickElement('/html/body/app-root/div/ng-component/div[1]/nav/mbb-navbar/div[2]/perfect-scrollbar/div/div[1]/div[1]/mat-tree/mat-nested-tree-node[1]/li/ul/mat-tree-node[1]/a/li/span')

        print("bbbbbbbbb")

        if self.loadCompleted('/html/body/app-root/div/ng-component/div[1]/div/div/div[1]/div/div/div/mbb-information-account/mbb-source-account/div/div[4]/div/div[1]/form/div[3]/div[2]/button',2000):
            self.clickElement('/html/body/app-root/div/ng-component/div[1]/div/div/div[1]/div/div/div/mbb-information-account/mbb-source-account/div/div[4]/div/div[1]/form/div[3]/div[2]/button')

        print('query clicked...')
        
        # extract table data
        
        table_element = WebDriverWait(self.driver, 30).until(
            EC.visibility_of_element_located((By.CLASS_NAME, 'table-striped'))
        )
        
        rows = table_element.find_elements(By.TAG_NAME, 'tr')
        table_data = []
        with open('extract_data.txt', mode='w', encoding='utf-8') as log_file:
            log_file.write(' ' + '\n')
        
        print("length",len(rows))
        # time.sleep(30000)
        for row in rows:
            header_rows = row.find_elements(By.TAG_NAME, "th")
            h_row_data = [h_cell.text for h_cell in header_rows]
            table_data.append(h_row_data)
            with open('extract_data.txt', mode='a', encoding='utf-8') as log_file:
                log_file.write(' '.join(h_row_data) + '\n')

            cells = row.find_elements(By.TAG_NAME, 'td')
            row_data = [cell.text for cell in cells]
            table_data.append(row_data)
            with open('extract_data.txt', mode='a', encoding='utf-8') as log_file:
                log_file.write(' '.join(row_data) + '\n')

        time.sleep(30)
        # self.driver.quit()
        self.runDownload()

    def isLoginError(self):
        xpath_element = '//*[@id="maincontent"]/ng-component/div[1]/div/div[3]/div/div/div/app-login-form/div/div/div[4]/p'
        login_error = WebDriverWait(self.driver, 20).until(
            EC.visibility_of_element_located((By.XPATH, xpath_element))).text
        print("login_error", login_error)

        if login_error == 'Mã kiểm tra không chính xác. Quý khách vui lòng kiểm tra lại.':
            return True
        return False

if __name__ == "__main__":
    file_name = f'.\\setting.json'
    with open(file_name) as file:
        info = json.load(file)
    AutoDownloadMBbank(info['USER_NAME'],info['PASSWORD'])