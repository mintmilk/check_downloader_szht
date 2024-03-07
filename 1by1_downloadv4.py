import os
import csv
import time
import base64
import ddddocr
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, UnexpectedAlertPresentException

def web_login(driver, username, password):
    max_attempts = 5
    for attempt in range(1, max_attempts + 1):
        try:
            driver.get("http://cwc.cau.edu.cn:8080/dlpt/Login_cau.aspx")
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="login"]/div/div/div[2]/p/img'))).click()
            driver.find_element(By.ID, "Txt_UserName").send_keys(username)
            driver.find_element(By.ID, "Txt_PassWord").send_keys(password)
            code = driver.find_element(By.CSS_SELECTOR, "#Txt_Yzm")
            imgCode = driver.find_element(By.CSS_SELECTOR, "#yzm")
            imgCode.screenshot("code.png")
            time.sleep(1)
            ocr = ddddocr.DdddOcr()
            with open("code.png", "rb") as fp:
                image = fp.read()
            catch = ocr.classification(image)
            code.send_keys(catch)
            time.sleep(1)
            driver.find_element(By.ID, "Btn_login").click()
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'LinkButton_wscx'))).click()
            time.sleep(10)
            print('Successfully logged in.')
            # 切换到新打开的标签页
            driver.switch_to.window(driver.window_handles[-1])
            
            return True
        except Exception as e:
            print(f'Login failed due to error: {str(e)}. Retrying... ({attempt}/{max_attempts})')
            if attempt < max_attempts:
                time.sleep(3)
            else:
                print(f'Login failed after {max_attempts} attempts.')
                raise

def combine_string(row):
    column1 = row[0]
    column2 = row[1]
    column1 = column1.replace("-", "")
    mid = len(column1) // 2
    if column2.startswith('X'):
        column2 = column2[1:]
        string = f"directory=/images/VoucherImages/&filename={column1}-{column2}.zip&type=pz&pzbh={column2}&pzrq={column1}&pznm={column1[mid-2:mid+2]}02{column2}&yhlx=teacher"
    elif column2.startswith('D'):
        column2 = column2[1:]
        string = f"directory=/images/VoucherImages/&filename={column1}-{column2}.zip&type=pz&pzbh={column2}&pzrq={column1}&pznm={column1[mid-2:mid+2]}01{column2}&yhlx=teacher"
    else:
        raise ValueError(f"Invalid column2 value: {column2}")
    print(f"Processing: {column1} {column2}")
    return string

def build_base64_url(string):
    encoded_string = base64.b64encode(string.encode()).decode()
    url = f"http://wszz.cau.edu.cn/cxzx/Views/Common/ImagePreview.aspx?urls={encoded_string}"
    return url

def download_file(url, driver, download_dir, username, password):
    '''
    依次打开每个页面下载，如果登录信息过期，则重新登录
    '''
    driver.execute_script("window.open('{}');".format(url))
    driver.switch_to.window(driver.window_handles[-1])

    print(f"Downloading: {url}")

    wait = WebDriverWait(driver, 10)
    
    try:
        wait.until(lambda driver: driver.execute_script("return document.readyState") == "complete")
    except TimeoutException:
        print(f"Page load timed out for: {url}. Retrying with new login...")
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        web_login(driver, username, password)
        return download_file(url, driver, download_dir, username, password)

    try:
        alert = wait.until(EC.alert_is_present())
        alert.send_keys("Wait")
        alert.accept()
    except (TimeoutException, UnexpectedAlertPresentException):
        pass

    try:
        download_button = wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_Lnk_download")))
        download_button.click()
    except TimeoutException:
        print(f"Download button not found for: {url}. Retrying with new login...")
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        web_login(driver, username, password)
        return download_file(url, driver, download_dir, username, password)

    time.sleep(5)  # 等待文件下载开始

    while any(filename.endswith('.crdownload') for filename in os.listdir(download_dir)):
        time.sleep(1)  # 等待文件下载完成

    print(f"Download completed: {url}")

def main(username, password, csv_file, download_dir):
    chrome_options = Options()
    chrome_options.add_experimental_option('prefs', {
        'download.default_directory': download_dir,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': False
    })
    chrome_options.add_argument("--unsafely-treat-insecure-origin-as-secure=http://wszz.cau.edu.cn")
    driver = webdriver.Chrome(options=chrome_options)

    try:
        web_login(driver, username, password)

        with open(csv_file, 'r') as file:
            csvreader = csv.reader(file)
            print('Reading data...')
            for row in csvreader:
                string = combine_string(row)
                url = build_base64_url(string)
                download_file(url, driver, download_dir, username, password)
                
                # 关闭除了前两个标签页以外的所有标签页
                while len(driver.window_handles) > 2:
                    driver.switch_to.window(driver.window_handles[-1])
                    driver.close()
                
                # 切换回第二个标签页
                driver.switch_to.window(driver.window_handles[1])
    finally:
        driver.quit()

if __name__ == "__main__":
    username = "08003"
    password = "203656"
    csv_file = "data.csv"
    download_dir = r"C:\Users\xbai6\Downloads\checks"

    main(username, password, csv_file, download_dir)