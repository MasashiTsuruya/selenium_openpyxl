from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import sys
import os
import datetime
import time
import glob
import configparser
import openpyxl

INIFILE = "config.ini"
INISECTION = "web"
INIADDRESS = "address"
INIID = "id"
INIPASSWORD = "password"

class GrapeCity:
    """seleniumによるブラウザ操作"""
    def __init__(self, driver):
        self.driver = driver

    def login(self, config):
        try:
            #webページを開く
            self.driver.get(config[INIADDRESS])
            #要素のIdで検索　　send_keys()でフォームに入力
            self.driver.find_element_by_id('userId').send_keys(config[INIID])
            #要素の名前で検索　
            self.driver.find_element_by_name('password').send_keys(config[INIPASSWORD])
            time.sleep(1)
            self.driver.find_element_by_id('submit').click()
            print('ログイン成功')
            return True
        except:
            print('ログインエラー')
            return False

    def get_File(self):
        try:
            self.driver.find_element_by_id('globalRecentProjectLink').click()
            time.sleep(1)
            #リンクテキスト名で検索
            self.driver.find_element_by_link_text('課題').click()
            #クラス名(複数)で検索    完了以外を選択する
            self.driver.find_elements_by_class_name('filter-nav__link')[5].click()
            self.driver.find_elements_by_class_name('dropdown')[4].click()
            time.sleep(1)
            self.driver.find_element_by_link_text('Excel').click()
            print('ファイルダウンロード完了')
            return True
        except:
            print('ファイルを取得できませんでした')
            return False

class Excel():
    """openpyxlによるExcel操作"""

    def __init__(self, sheet, new_sheet):
        self.sheet = sheet
        self.new_sheet = new_sheet

    def copy_cell(self):
        try:
            for row in self.sheet:
                for cell in row:
                    #一行目はコピーしない
                    if (cell.row == 1):
                        pass
                    #a~z列までのコピー
                    elif (cell.column > 28):
                        pass
                    else:
                        self.new_sheet[cell.coordinate].value = cell.value
            return True
        except:
            print('セルのコピーに失敗しました')
            return False

if os.path.exists(INIFILE):
    config = configparser.ConfigParser()
    config.read(INIFILE, encoding='utf-8')
else:
    print(INIFILE + "がありません。")
    sys.exit()

def return_check(result):
    """Trueが返って来ているかチェック"""
    if result is False:
        sys.exit()

options = Options()
#画面サイズ最大を指定
options.add_argument('--start-maximized')
driver = webdriver.Chrome(options=options)
grape_city = GrapeCity(driver)
result = grape_city.login(config[INISECTION])

return_check(result)

xl_file = grape_city.get_File()
return_check(xl_file)
time.sleep(1)

today = str(datetime.date.today()).replace('-', '')
username = os.getlogin()
files = glob.glob('C:\\Users\\{0}\\Downloads\\Backlog-Issues-{1}*.xlsx'.format(username, today))

if not files:
    print('ファイルが見つかりません')
    sys.exit()
#一番最新のファイルを選択
file = files[-1]
wb = openpyxl.load_workbook(file, keep_vba=True)
new_wb = openpyxl.load_workbook('テンプレートYYYYMMDD.xlsx')

sheet = wb.worksheets[0]
new_sheet = new_wb.worksheets[0]

excel = Excel(sheet, new_sheet)
ed_xl = excel.copy_cell()

return_check(ed_xl)
try:
    new_wb.save('C:\\Users\\{0}\\Desktop\\課題管理一覧{1}.xlsx'.format(username, today))
    print('success')
except:
    print('Excelの保存に失敗しました')
