import sys
import time
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
import os.path
import win32com.client
import pathlib
from pathlib import Path
from tkinter import *

dir_path = pathlib.Path.cwd()
check_driver('C:\WINDOWS\system32')

with open('cello.txt', 'r', encoding='utf-8') as f:
    for line in f:
        line = line.strip()
        login, password, domain = line.split('::')


def cello():
    options = Options()
    options.add_argument("--headless")  # Runs Chrome in headless mode.
    options.add_argument('--no-sandbox')  # Bypass OS security model
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
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 20)
    driver.implicitly_wait(20)
    try:
        driver.get(domain)
        print("browser opened", file=TextWrapper(txt_log))
        window.update()
    except:
        messagebox.showerror('Ошибка открытия браузера', 'Браузер не открылся!')
        print('ошибка - не открылся браузер')
        print("ошибка - не открылся браузер", file=TextWrapper(txt_log))

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

    try:
        wait.until(EC.element_to_be_clickable((By.ID, 'btn-ok')))
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
        print("login 2", file=TextWrapper(txt_log))
        window.update()
    except:
        print('login normal')
        print("login 1", file=TextWrapper(txt_log))
        window.update()

    def to_download():

        # TMS
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="menuTBox"]/section/div[1]/ul/li[5]/a')))
        driver.find_element(By.XPATH, '//*[@id="menuTBox"]/section/div[1]/ul/li[5]/a').click()
        # Prime
        wait.until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="sideMenu"]/section/ul/li[10]/label/div/p')))
        driver.find_element(By.XPATH, '//*[@id="sideMenu"]/section/ul/li[10]/label/div/p').click()
        # Transport Order
        driver.find_element(By.XPATH,
                            '/html/body/div[1]/main/div[1]/div/section/ul/li[10]/ul/li[2]/label/div/p').click()
        # T/O List
        driver.find_element(By.XPATH, '/html/body/div[1]/main/div[1]/div/section/ul/li[10]/ul/li[2]/ul/li/p').click()

        time.sleep(6)
        iframe = driver.find_element(By.TAG_NAME, 'iframe')
        driver.switch_to.frame(iframe)
        # Ввод списка DO
        do_list = get_form_text()
        if do_list != ['']:
            driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[1]/div/input[1]').send_keys(
                str(do_list).strip('[]'))
        # для скачивания по диапазону дат без DO
        # else:
        #    driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/span[2]').click()
        #    driver.find_element(By.XPATH, '/html/body/div[52]/div/div[2]/div[2]/div/span[1]').click()

        driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[9]/button[2]').click()

        time.sleep(3)
        element = wait.until(EC.element_to_be_clickable(
            (By.XPATH, '/html/body/div[4]/div/div[2]/div/div[3]/div[2]/div/div[1]/div[1]/div/div[1]/div')))
        driver.find_element(By.ID, 'btn_multiExlDown').click()
        time.sleep(1)
        # Ожидание скачивания файла
        while True:
            list_of_files = glob.glob(
                os.path.join(os.path.join(os.path.join(os.getcwd()), 'TOListWithItemInfo_*.xlsx')))
            time.sleep(5)
            if len(list_of_files) > 0:
                break

        driver.switch_to.default_content()
        # закрываем вкладку T/O list
        driver.find_element(By.XPATH, '/html/body/div[1]/main/div[2]/div[1]/div[2]/div[1]/ul/li[3]/div').click()
        print("TOList file ok", file=TextWrapper(txt_log))
        window.update()

    to_download()

    def wms_download():
        # WMS
        driver.switch_to.default_content()
        driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[1]').click()
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[2]/div[2]/section/div[1]/ul/li[4]')))
        driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[2]/section/div[1]/ul/li[4]').click()
        # Outbound
        wait.until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/main/div[1]/div/section/ul/li[6]/label/div/p')))
        driver.find_element(By.XPATH, '/html/body/div[1]/main/div[1]/div/section/ul/li[6]/label/div/p').click()
        # Serial Scan
        driver.find_element(By.XPATH, '/html/body/div[1]/main/div[1]/div/section/ul/li[6]/ul/li[6]/label/div/p').click()
        # Order
        driver.find_element(By.XPATH, '/html/body/div[1]/main/div[1]/div/section/ul/li[6]/ul/li[6]/ul/li[1]/p').click()

        time.sleep(2)
        driver.find_element(By.TAG_NAME, 'iframe')
        iframe = driver.find_element(By.TAG_NAME, 'iframe')
        driver.switch_to.frame(iframe)
        # calendar
        driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[1]/div/div/div[3]/div/div[1]/div').click()
        driver.find_element(By.XPATH, '/html/div[1]/div[3]/button[3]').click()
        driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[2]/div/div/div[3]/div/div[1]/div').click()
        driver.find_element(By.XPATH, '/html/div[2]/div[3]/button[3]').click()
        driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[3]/button[2]').click()
        driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[7]/div/div/div/div[2]').click()
        wait.until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[23]/div/div/div/div[2]/div/div[4]/span')))
        driver.find_element(By.XPATH, '/html/body/div[23]/div/div/div/div[2]/div/div[4]/span').click()

        # Ввод списка DO
        do_list = get_form_text()
        if do_list != ['']:
            driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[3]/div/input[1]').send_keys(str(do_list).strip('[]'))
        time.sleep(1)
        driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[15]/button[2]').click()
        time.sleep(2)
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '/html/body/div[4]/div/div[2]/div/div[3]/div[2]/div/div[1]/div[1]/div/div[1]/div')))
        driver.find_element(By.XPATH,
                            '/html/body/div[4]/div/div[2]/div/div[3]/div[1]/div/div[1]/div/div/div/div').click()

        driver.find_element(By.ID, 'btn_exlDownAllMiw').click()
        time.sleep(2)
        while True:
            list_of_files = glob.glob(
                os.path.join(os.path.join(os.path.join(os.getcwd()), 'wmsOrderSerialScan_AllInfo_*.xls')))
            time.sleep(5)
            if len(list_of_files) > 0:
                break
        print("wms file ok", file=TextWrapper(txt_log))
        window.update()

    wms_download()

    try:
        driver.stop_client()
        driver.close()
        driver.quit()
    except:

        messagebox.showerror('Ошибка закрытия браузера', 'Браузер не закрылся после выполнения!')


def parse_tolist():
    if os.path.exists(Path(pathlib.Path.cwd(), 'TOListWithItemInfo.xlsx')):
        os.remove(Path(pathlib.Path.cwd(), 'TOListWithItemInfo.xlsx'))
    try:
        list_of_files = glob.glob(os.path.join(os.path.join(os.path.join(os.getcwd()), 'TOListWithItemInfo_*.xlsx')))
        wmsbook = max(list_of_files, key=os.path.getctime)
    except ValueError:
        messagebox.showerror('Не найден файл', 'Файл(ы) TOListWithItemInfo_* не найдены в папке!')
        print('файл tolist не найден')
        sys.exit()

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel_path = os.path.expanduser(wmsbook)

    try:
        workbook = excel.Workbooks.Open(Filename=excel_path)
        workbook.SaveAs(Filename := str(Path(pathlib.Path.cwd(), 'TOListWithItemInfo.xlsx')), FileFormat := 51)
        workbook.Close()
        excel.Application.Quit()
        del excel
        os.remove(wmsbook)
    except IOError:

        messagebox.showerror('Файл TOListWithItemInfo недоступен',
                             'Файл(ы) TOListWithItemInfo_* не найден или недоступен из папки!')


def parse_wms():
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
    try:
        wb = excel.Workbooks.Open(os.path.abspath("estore_macro.xlsb"), ReadOnly=1)
        print("estore_macro.xlsb открыт ", file=TextWrapper(txt_log))
        window.update()
        macros = excel.Application.Run("estore_macro.xlsb!Module1.create_doc_act")
        print("макрос запущен ", file=TextWrapper(txt_log))
        window.update()
        wb.Close(False)
        print("estore_macro.xlsb закрыт ", file=TextWrapper(txt_log))
        window.update()
        print("акты созданы", file=TextWrapper(txt_log))
        print("работа завершена", file=TextWrapper(txt_log))
        window.update()
    except:
        messagebox.showerror('file not found', 'Файл с макросом не найден')


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


def get_from_cello():
    do_list = get_form_text()
    print(do_list)
    if do_list != ['']:
        cello()
        parse_tolist()
        parse_wms()
    else:
        messagebox.showerror('input DO', 'Введите DO')


window = tk.Tk()
window.title("Estore")
window.geometry('620x400')
window.columnconfigure([0, 1, 2], minsize=60, pad=60, weight=1)
window.rowconfigure([0, 1, 2], minsize=60, weight=2)
lbl = Label(window, text="Input DO No.:", font="Courier 16")
lbl.grid(column=0, row=0)
lbl_log = Label(window, text="Log:", font="Courier 14")
lbl_log.grid(column=2, row=0)
txt = scrolledtext.ScrolledText(window, width=30, height=50, font="Courier 16")
txt.grid(column=0, row=1)
txt_log = scrolledtext.ScrolledText(window, width=50, height=50, font="Courier 12", bg="black", fg="green")
txt_log.grid(column=2, row=1)
cello_btn = Button(window, text="get from Cello", command=get_from_cello, font="Courier 12", fg="white", bg="Green")
cello_btn.grid(column=0, row=2, sticky='nsew')
clean_btn = Button(window, text="Cleanup", command=cleanup, bg="#aeb6bf", fg="white")
clean_btn.grid(column=1, row=2, sticky='NSWE')
macro_btn = Button(window, text="START MACROS", command=runup_macro, bg="#748efa", fg="white")
macro_btn.grid(column=2, row=2, sticky='nsew')
window.update()


class TextWrapper:
    text_field: tk.Text

    def __init__(self, text_field: tk.Text):
        self.text_field = text_field

    def write(self, text: str):
        self.text_field.insert(tk.END, text)

    def flush(self):
        self.text_field.update()


window.mainloop()
