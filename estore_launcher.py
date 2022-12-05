import sys
import time
import datetime

import tkinter as tk
from tkinter import messagebox, scrolledtext
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_auto_update import check_driver
import glob
import os
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import os.path
import win32com.client
import pathlib
from pathlib import Path
from tkinter import *

current_time = datetime.datetime.now()
dir_path = pathlib.Path.cwd()
check_driver('C:\WINDOWS\system32')

with open('cello.txt', 'r', encoding='utf-8') as f:
    for line in f:
        line = line.strip()
        login, password, domain = line.split('::')

print(login, domain)


def cello():
    options = Options()
    # options.add_argument("--headless")  # Runs Chrome in headless mode.
    options.add_argument('--no-sandbox')  # Bypass OS security model
    options.add_argument('--disable-gpu')  # applicable to Windows os only
    options.add_argument('start-maximized')
    options.add_argument('disable-infobars')
    options.add_argument('--disable-extensions')
    options.add_experimental_option("prefs", {
        "safebrowsing_for_trusted_sources_enabled": False,
        "safebrowsing.enabled": False,
        "download.default_directory": str(dir_path),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
    })
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 20)
    try:
        driver.get(domain)
        time.sleep(4)
    except:
        messagebox.showerror('Ошибка открытия браузера', 'Браузер не открылся!')
        print('ошибка - не открылся браузер')

        try:
            driver.stop_client()
            driver.close()
            driver.quit()
        except:
            messagebox.showerror('Ошибка закрытия браузера', 'Браузер не закрылся после выполнения!')
            print('браузер не закрыт')
    driver.find_element(By.ID, "loginId").click()
    driver.find_element(By.ID, "loginId").send_keys(login)
    driver.find_element(By.ID, "loginPw").send_keys(password)
    driver.find_element(By.ID, "btn-login").click()
    wait = WebDriverWait(driver, 20)
    try:
        element = wait.until(EC.element_to_be_clickable((By.ID, 'btn-ok')))
        driver.find_element(By.ID, 'btn-ok').click()
    except:
        print('Login stop')
        driver.stop_client()
        driver.close()
        driver.quit()
        messagebox.showerror('Cello', 'Login problem')

    try:
        driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div/div/div[4]/button').click()
        print('login after')
    except:
        print('login normal')

    # TMS
    element = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="menuTBox"]/section/div[1]/ul/li[5]/a')))
    driver.find_element(By.XPATH, '//*[@id="menuTBox"]/section/div[1]/ul/li[5]/a').click()
    # Prime
    element = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="sideMenu"]/section/ul/li[10]/label/div/p')))
    driver.find_element(By.XPATH, '//*[@id="sideMenu"]/section/ul/li[10]/label/div/p').click()
    # Transport Order
    driver.find_element(By.XPATH, '/html/body/div[1]/main/div[1]/div/section/ul/li[10]/ul/li[2]/label/div/p').click()
    # T/O List
    driver.find_element(By.XPATH, '/html/body/div[1]/main/div[1]/div/section/ul/li[10]/ul/li[2]/ul/li/p').click()

    time.sleep(6)
    iframe = driver.find_element(By.TAG_NAME, 'iframe')
    driver.switch_to.frame(iframe)
    # Ввод списка DO
    do_list = get_form_text()
    if len(do_list) > 0:
        driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[1]/div/input[1]').send_keys(
            str(do_list).strip('[]'))
    else:
        driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/span[2]').click()
        driver.find_element(By.XPATH, '/html/body/div[52]/div/div[2]/div[2]/div/span[1]').click()

    driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[9]/button[2]').click()

    time.sleep(3)
    element = wait.until(EC.element_to_be_clickable(
        (By.XPATH, '/html/body/div[4]/div/div[2]/div/div[3]/div[2]/div/div[1]/div[1]/div/div[1]/div')))
    driver.find_element(By.ID, 'btn_multiExlDown').click()
    time.sleep(1)
    # Ожидание скачивания файла
    while True:
        list_of_files = glob.glob(os.path.join(os.path.join(os.path.join(os.getcwd()), 'TOListWithItemInfo_*.xlsx')))
        time.sleep(5)
        if len(list_of_files) > 0:
            break

    driver.switch_to.default_content()
    # закрываем вкладку T/O list
    driver.find_element(By.XPATH, '/html/body/div[1]/main/div[2]/div[1]/div[2]/div[1]/ul/li[3]/div').click()

    driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[1]').click()

    driver.find_element(By.XPATH,
                        '/html/body/div[1]/main/div/div[2]/div[1]/div[1]/div[2]/section/div[1]/ul/li[1]').click()
    time.sleep(6)
    iframe = driver.find_element(By.TAG_NAME, 'iframe')
    driver.switch_to.frame(iframe)
    driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[1]/div/div/div[3]/div/div[1]/div').click()
    driver.find_element(By.XPATH, '/html/div[1]/div[3]/button[3]').click()
    driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[2]/div/div/div[3]/div/div[1]/div').click()
    driver.find_element(By.XPATH, '/html/div[2]/div[3]/button[3]').click()
    driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[3]/button[2]').click()
    driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[7]/div/div/div/div[2]/div').click()

    drop_down = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[7]/div/div/div/div[2]/div')
    ActionChains(driver).move_to_element(drop_down).perform()
    ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
    time.sleep(0.2)
    ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
    ActionChains(driver).send_keys(Keys.ENTER).perform()

    driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[14]/button[2]').click()
    time.sleep(2)
    element = wait.until(EC.element_to_be_clickable(
        (By.XPATH, '/html/body/div[4]/div/div[2]/div/div[3]/div[2]/div/div[1]/div[1]/div/div[1]/div')))
    driver.find_element(By.XPATH, '/html/body/div[4]/div/div[2]/div/div[3]/div[1]/div/div[1]/div/div/div/div').click()

    driver.find_element(By.ID, 'btn_exlDownAllMiw').click()
    time.sleep(2)
    while True:
        list_of_files = glob.glob(
            os.path.join(os.path.join(os.path.join(os.getcwd()), 'wmsOrderSerialScan_AllInfo_*.xls')))
        time.sleep(5)
        if len(list_of_files) > 0:
            break
    try:
        driver.stop_client()
        driver.close()
        driver.quit()
    except:
        print('браузер не закрыт')
        messagebox.showerror('Ошибка закрытия браузера', 'Браузер не закрылся после выполнения!')


# cello()


def parse_tolist():
    cello()
    if os.path.exists(Path(pathlib.Path.cwd(), 'TOListWithItemInfo.xlsx')):
        os.remove(Path(pathlib.Path.cwd(), 'TOListWithItemInfo.xlsx'))
    try:
        list_of_files = glob.glob(os.path.join(os.path.join(os.path.join(os.getcwd()), 'TOListWithItemInfo_*.xlsx')))
        # list_of_files = glob.glob('C:\\Users\\K.Zakharov\\PycharmProjects\\estore\\wmsOrderSerialScan_AllInfo_*.xls')
        wmsbook = max(list_of_files, key=os.path.getctime)
    except ValueError:
        messagebox.showerror('Не найден файл', 'Файл(ы) TOListWithItemInfo_* не найдены в папке!')
        print('файл tolist не найден')
        sys.exit()

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel_path = os.path.expanduser(wmsbook)
    print(excel_path)
    try:
        workbook = excel.Workbooks.Open(Filename=excel_path)
        workbook.SaveAs(Filename := str(Path(pathlib.Path.cwd(), 'TOListWithItemInfo.xlsx')), FileFormat := 51)
        workbook.Close()
        excel.Application.Quit()
        del excel
        os.remove(wmsbook)
    except IOError:
        print('файл tolist  недоступен')
        messagebox.showerror('Файл TOListWithItemInfo недоступен',
                             'Файл(ы) TOListWithItemInfo_* не найден или недоступен из папки!')


def parse_wms():
    parse_tolist()
    if os.path.exists(Path(pathlib.Path.cwd(), 'wms.xlsx')):
        os.remove(Path(pathlib.Path.cwd(), 'wms.xlsx'))
    try:
        list_of_files = glob.glob(
            os.path.join(os.path.join(os.path.join(os.getcwd()), 'wmsOrderSerialScan_AllInfo_*.xls')))
        # list_of_files = glob.glob('C:\\Users\\K.Zakharov\\PycharmProjects\\estore\\wmsOrderSerialScan_AllInfo_*.xls')
        wmsbook = max(list_of_files, key=os.path.getctime)
    except ValueError:
        print('файл wms не найден')
        messagebox.showerror('Файл wms.xlsx недоступен',
                             'Файл(ы) wms.xlsx не найден или недоступен из папки!')
        sys.exit()

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel_path = os.path.expanduser(wmsbook)
    print(excel_path)
    try:
        workbook = excel.Workbooks.Open(Filename=excel_path)
        workbook.SaveAs(Filename := str(Path(pathlib.Path.cwd(), 'wms.xlsx')), FileFormat := 51)
        workbook.Close()
        excel.Application.Quit()
        del excel
        os.remove(wmsbook)
    except IOError:
        print('файл wms недоступен')
        messagebox.showerror('Файл wms.xlsx недоступен',
                             'Файл(ы) wms.xlsx не найден или недоступен из папки!')


def runup_macro():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.Workbooks.Open(os.path.abspath("estore_macro.xlsb"), ReadOnly=1)
    excel.Application.Run("estore_macro.xlsb!Module1.create_doc_act")
    excel.Application.Quit()  # Comment this out if your excel script closes


def get_form_text() -> list[str]:
    """
     функция получает данные из формы
    :return: массив ордеров
    """
    res = txt.get('1.0', 'end-1c')
    order_arr = []
    for do in res.split("\n"):
        line_splitted: list[str] = do.strip().split("\n")
        order_arr += line_splitted
    print(order_arr)
    return order_arr


def cleanup():
    txt.delete('1.0', END)


window = tk.Tk()
window.title("Estore")
window.geometry('620x400')
window.columnconfigure([0, 1], minsize=60, pad=60, weight=1)
window.rowconfigure([0, 1, 2], minsize=50, weight=2)
lbl = Label(window, text="Input DO No.:", font="Courier 16")
lbl.grid(column=0, row=0)
txt = scrolledtext.ScrolledText(window, width=20, height=50, font="Courier 16")
txt.grid(column=0, row=1)
cello_btn = Button(window, text="get from Cello", command=parse_wms, font="Courier 12", fg="white", bg="Green")

clean_btn = Button(window, text="Cleanup form", command=cleanup, bg="#aeb6bf", fg="white")
clean_btn.grid(column=0, row=2, sticky='NSWE')
cello_btn.grid(column=1, row=1, sticky='nsew')
macro_btn = Button(window, text="START MACROS", command=runup_macro, bg="#748efa", fg="white")
macro_btn.grid(column=1, row=2, sticky='nsew')
window.mainloop()
