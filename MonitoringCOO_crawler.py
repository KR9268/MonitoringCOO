from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, UnexpectedAlertPresentException, WebDriverException
from selenium.webdriver.chrome.options import Options
from win32 import win32gui
import json
import random
import sqlite3
import ctypes # open_kochamr2
import pygetwindow as gw # open_kochamr2
import pyautogui

from pkb_selenium import *

file_path_db = r'\\23.20.135.83\파일 공유\MoniteringCOO' + '\\'
#file_path_db = 'C:\\python_source\\MonitoringCOO\\'
#file_path_db = 'D:\\파일 공유\\MoniteringCOO\\'


# 함수 정의: 인증서 창을 찾는 함수
def find_window_by_name(window_name='인증서'):
    hwnd = None
    while True:
        hwnd = ctypes.windll.user32.FindWindowW(None, window_name)
        if hwnd: return hwnd
        #time.sleep(1)  # 인증서 창이 나타날 때까지 대기  

def input_keybd_ctypes(keytext:hex, delay=0, cap=None, memo=None):
    if cap is not None:
        ctypes.windll.user32.keybd_event(0x10, 0, 0, 0) # Shift 누르기
    ctypes.windll.user32.keybd_event(keytext, 0, 0, 0)
    ctypes.windll.user32.keybd_event(keytext, 0, 2, 0)
    if cap is not None:
        ctypes.windll.user32.keybd_event(0x10, 0, 2, 0) # Shift 떼기
    time.sleep(delay)

def close_all_except_original_window(driver):
    original_window = driver.current_window_handle
    for handle in driver.window_handles:
        if handle != original_window:
            driver.switch_to.window(handle)
            driver.close()
    driver.switch_to.window(original_window)
    return driver

def login_gongdong_bykbd(gongdong_location, number_of_expire_soon):
    for _ in pyautogui.getAllWindows():
        if '인증서' in _.title:
            window = _
    while True:
        window.activate()
        if get_foreground_window_text()=='인증서':
            break        
    pyautogui.hotkey("tab",interval=0.2)
    if gongdong_location == 2:
        pyautogui.hotkey("down", "down",interval=0.2)
        for i in range(number_of_expire_soon):
            pyautogui.hotkey("tab",interval=0.2)
    elif gongdong_location == 1:
        for i in range(number_of_expire_soon):
            pyautogui.hotkey("tab",interval=0.2)
        pass
    else: #향후 별도로직을 고려
        pyautogui.hotkey("down", "down",interval=0.2)
    pyautogui.hotkey("tab",interval=0.2) #공통작업

def login_gongdong_byimg(img_company:str, pass_gondong):
    """
        img_company = SEC_NEGO1, SEC_NEGO2, SMC_NEGO1
    """
    while True:
        location_comp = pyautogui.locateCenterOnScreen(f'1_{img_company}_gray.png')
        if location_comp:
            pyautogui.leftClick(location_comp)
            break

        location_comp2 = pyautogui.locateCenterOnScreen(f'1_{img_company}_white.png')
        if location_comp2:
            pyautogui.leftClick(location_comp)
            break

        pyautogui.hotkey('winleft','m') # 이미지 인식용

    while True:
        location_pw = pyautogui.locateCenterOnScreen(r'2_pwbox.png')
        if location_pw:
            pyautogui.leftClick(location_pw)
            break
        pyautogui.hotkey('winleft','m') # 이미지 인식용
        #caps

    pyperclip.copy(pass_gondong)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(1)
    #time.sleep(2)  ################테스트용 임시추가
    pyautogui.hotkey("enter")
    time.sleep(2)


def open_korcham(korcham_opt, idtype, headless=None, size_mini=None):
    id, passwd, pass_gondong = korcham_opt[idtype]
    # 공동인증서 키보드 로그인시에만 사용
    # gongdong_location = korcham_opt['loc_and_expiry'][idtype[:3]][0]
    # number_of_expire_soon = int(0) # 날짜로직 반영 필요

     # 상공회의소 접속
    url = "https://cert.korcham.net/base/index.htm"
    driver = set_browser(url, headless=headless, size_mini=size_mini)
    # 팝업끄기
    driver = close_all_except_original_window(driver)
    # 로그인(웹)
    driver.implicitly_wait(10)
    button_click(driver, '공동인증서 로그인 선택', By.CSS_SELECTOR, "#loginForm > div > fieldset > ul > li:nth-child(2)")
    #button_click(driver, '공동인증서 로그인 선택', By.XPATH, "#/html/body/div[2]/div[5]/section/form/div/fieldset/ul/li[2]")
    box_input(driver, 'id_box', By.CSS_SELECTOR, "#inputCoId", id)
    box_input(driver, 'pass_box', By.CSS_SELECTOR, "#userId", passwd)
    button_click(driver, '로그인 버튼', By.CSS_SELECTOR,"#btnLogin")

    # 설치오류시 대응 (URL로 판단  https://rd.korcham.net/NX_TPKIENT/Install/security.html  )


    # 공동인증서 로그인
    time.sleep(2)
    img_company = idtype[:3].lower()
    login_gongdong_byimg(img_company, pass_gondong)
    # 팝업끄기(로그인 후)
    driver = close_all_except_original_window(driver)
    # 발급현황 메뉴가기
    driver.implicitly_wait(10)
    try:
        button_click(driver, '원산지증명서 메뉴선택', By.CSS_SELECTOR,"#gnb > ul > li:nth-child(2) > a")
    except UnexpectedAlertPresentException:
            alert = driver.switch_to.alert
            print(alert.text)
            driver.switch_to.default_content



    driver.implicitly_wait(10)
    button_click(driver, '발급현황 메뉴선택', By.CSS_SELECTOR,"body > div.wrapper > article > div.contWrap.clearfix > nav > ul > li:nth-child(5) > a")
    driver.implicitly_wait(10)

    return driver

def open_korcham2(korcham_opt, idtype, headless=None, size_mini=None):
    id, passwd, pass_gondong = korcham_opt[idtype]
    gongdong_location = korcham_opt['loc_and_expiry'][idtype[:3]][0]
    # 날짜로직 반영 필요
    number_of_expire_soon = int(0)

    # 상공회의소 접속
    url = "https://cert.korcham.net/base/index.htm"
    driver = set_browser(url, headless=headless, size_mini=size_mini)
    # 팝업끄기
    driver = close_all_except_original_window(driver)
    # 로그인(웹)
    driver.implicitly_wait(10)
    button_click(driver, '공동인증서 로그인 선택', By.CSS_SELECTOR, "#loginForm > div > fieldset > ul > li:nth-child(2)")
    #button_click(driver, '공동인증서 로그인 선택', By.XPATH, "#/html/body/div[2]/div[5]/section/form/div/fieldset/ul/li[2]")
    box_input(driver, 'id_box', By.CSS_SELECTOR, "#inputCoId", id)
    box_input(driver, 'pass_box', By.CSS_SELECTOR, "#userId", passwd)
    button_click(driver, '로그인 버튼', By.CSS_SELECTOR,"#btnLogin")
    
    # 공동인증서 로그인
    time.sleep(2)
    hwnd = find_window_by_name('인증서')
    while True:
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        time.sleep(1)
        active_window = gw.getActiveWindow()
        if active_window.title == '인증서':
            break

    input_keybd_ctypes(0x09, 0.5, 'Tab')
    if gongdong_location == 2:
        input_keybd_ctypes(0x28, 0.5, '↓')
        input_keybd_ctypes(0x28, 0.5, '↓')
    for i in range(number_of_expire_soon):
        input_keybd_ctypes(0x09, 0.5, 'Tab')
    input_keybd_ctypes(0x09, 0.5, 'Tab')
    time.sleep(1)

    for char in pass_gondong:
        if char.isupper():
            input_keybd_ctypes(ord(char), cap='Y', delay=0.2, memo='each_chr_of_password')
        elif char.islower():
            input_keybd_ctypes(ord(char), cap=None, delay=0.2, memo='each_chr_of_password')
        elif char.isdigit():
            input_keybd_ctypes(ord(char), cap=None, delay=0.2, memo='each_chr_of_password')
        else:
            input_keybd_ctypes(0x40, cap=None, delay=0.2, memo='each_chr_of_password')
            print(char)

    time.sleep(1)
    input_keybd_ctypes(0x0D, 0.2, 'Enter')


    # 팝업끄기(로그인 후)
    driver = close_all_except_original_window(driver)
    # 발급현황 메뉴가기
    driver.implicitly_wait(10)
    button_click(driver, '원산지증명서 메뉴선택', By.CSS_SELECTOR,"#gnb > ul > li:nth-child(2) > a")
    driver.implicitly_wait(10)
    button_click(driver, '발급현황 메뉴선택', By.CSS_SELECTOR,"body > div.wrapper > article > div.contWrap.clearfix > nav > ul > li:nth-child(5) > a")
    driver.implicitly_wait(10)

    return driver

def coo_crawler(driver, last_updated_date)->list:
    data = []
    page_qty = len(get_multi_elements(driver, '페이지 수', By.XPATH, '/html/body/div[2]/article/div[2]/div/section/form/div[4]/span/a'))
    #page_qty = 1
    for j in range(page_qty):
        driver.implicitly_wait(10)
        time.sleep(random.randrange(1,2))
        button_click(driver, '페이지번호 클릭', By.XPATH, f'/html/body/div[2]/article/div[2]/div/section/form/div[4]/span/a[{j+1}]')

        row_qty_table = len(get_multi_elements(driver, '표 행수', By.XPATH, '/html/body/div[2]/article/div[2]/div/section/form/div[3]/table/tbody/tr'))
        for i in range(row_qty_table):
            korcham_idx = get_text(driver, '접수번호', By.XPATH, f'/html/body/div[2]/article/div[2]/div/section/form/div[3]/table/tbody/tr[{i+1}]/td[1]')
            if korcham_idx == '검색된 결과가 없습니다.':
                return None
            doc_type = get_text(driver, '증명서종류', By.XPATH, f'/html/body/div[2]/article/div[2]/div/section/form/div[3]/table/tbody/tr[{i+1}]/td[2]')
            invoice = get_text(driver, '대표Invoice', By.XPATH, f'/html/body/div[2]/article/div[2]/div/section/form/div[3]/table/tbody/tr[{i+1}]/td[3]')
            receive_time = get_text(driver, '접수일시', By.XPATH, f'/html/body/div[2]/article/div[2]/div/section/form/div[3]/table/tbody/tr[{i+1}]/td[4]')
            # if last_updated_date is not None and (receive_time < last_updated_date):
            #     return data # 기존에 작업했던 건이면 현재까지 작업한 데이터만 넘긴다
            status = get_text(driver, '처리상태', By.XPATH, f'/html/body/div[2]/article/div[2]/div/section/form/div[3]/table/tbody/tr[{i+1}]/td[5]')

            if status == '오류통보':
                button_click(driver, '접수번호 클릭', By.XPATH, f'/html/body/div[2]/article/div[2]/div/section/form/div[3]/table/tbody/tr[{i+1}]/td[1]/a')
                driver.implicitly_wait(10)
                time.sleep(random.randrange(1,2))
                reason = get_text(driver, '오류텍스트 확인', By.CSS_SELECTOR, '#sendForm > div.boardWrap.list.center > div > table > tbody > tr > td > textarea')
                driver.implicitly_wait(10)
                time.sleep(random.randrange(1,2))
                driver.get('https://cert.korcham.net/cert/cert/status/list.htm')
                button_click(driver, '페이지번호 클릭', By.XPATH, f'/html/body/div[2]/article/div[2]/div/section/form/div[4]/span/a[{j+1}]')
            else:
                reason = None

            data.append((korcham_idx, doc_type, invoice, receive_time, status, reason))
    return data

def manage_db(each_company, sqltype, data=None):
    conn_db = sqlite3.connect(file_path_db + 'Korcham_status.db')
    db_cursor = conn_db.cursor()
    if sqltype == 'insert':
        db_cursor.executemany(f'INSERT or REPLACE INTO 상공_{each_company} VALUES (?,?,?,?,?,?)', data)
        # db_cursor.executemany(f'''INSERT INTO 상공_{each_company} (접수번호, 증명서종류, 대표Invoice, 접수일시, 처리상태, Remark)
        #                       SELECT 접수번호, 증명서종류, 대표Invoice, 접수일시, 처리상태, Remark
        #                       WHERE NOT EXISTS (
        #                       SELECT 1 FROM 상공_{each_company} WHERE 접수번호 = 접수번호 AND 처리상태 = 처리상태
        #                       )''', data)
        conn_db.commit()
        conn_db.close()
    elif sqltype == 'selectmax':
        max_data = db_cursor.execute(f'SELECT MAX(접수일시) FROM 상공_{each_company}').fetchall()[0][0]
        conn_db.close()
        return max_data


def main_crawler():
    # 기본정보 로딩
    with open(file_path_db + 'MonitoringCOO_crawler.json', 'r', encoding='utf-8')as f:
        korcham_opt = json.load(f, strict=False)

    # 크롤링 실행
    for each_company in [i for i in list(korcham_opt.keys()) if i not in['loc_and_expiry','last_update']]:
        selenium_driver =  open_korcham(korcham_opt, each_company, size_mini=[1000,200])
        last_updated_date = manage_db(each_company, 'selectmax')
        data = coo_crawler(selenium_driver, last_updated_date)
        if data is not None:
            manage_db(each_company, 'insert', data)

            korcham_opt['last_update'][each_company] = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
            with open(file_path_db + 'MonitoringCOO_crawler.json', 'w', encoding='utf-8')as f:
                json.dump(korcham_opt, f, indent=2, ensure_ascii=False)
        
        selenium_driver.close()

if __name__ == '__main__':
    main_crawler()

