from curses.ascii import NUL
from pty import slave_open
from xml.etree.ElementPath import xpath_tokenizer
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from datetime import datetime, timedelta
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome import service

import time
from datetime import datetime, date, timedelta
# import jpholiday

from datetime import datetime as dt
import calendar

# from ast import Or
# from urllib import request  # urllib.requestモジュールをインポート

# 重複防止用
import fcntl

# 自作モジュール
from modules.sendLine import send_line_notify


# # ################################
# # スプレッドシート読み込み
# # ################################
import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope =['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
# creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
creds = ServiceAccountCredentials.from_json_keyfile_name('/home/admin/TennisCourtChecker-Eto/client_secret.json', scope)
client = gspread.authorize(creds)

spreadsheet = client.open('TennisCourtChecker') # 操作したいスプレッドシートの名前を指定する
worksheet = spreadsheet.worksheet('江東区スポーツネット') # シートを指定する

configSheet = spreadsheet.worksheet('設定_江東')
SUN = configSheet.acell('C3').value
MON = configSheet.acell('C4').value
TUE = configSheet.acell('C5').value
WED = configSheet.acell('C6').value
THU = configSheet.acell('C7').value
FRI = configSheet.acell('C8').value
SAT = configSheet.acell('C9').value
HOL = configSheet.acell('C10').value
TOKEN = configSheet.acell('C14').value
LNT_FOR_ERROR = configSheet.acell('C15').value

result = []

def writeSheet(data):
    worksheet.delete_rows(2) # 1行目(空き状況)を削除
    worksheet.delete_rows(1) # 2行目(計測時間)を削除
    worksheet.append_row(data) # dataを最終行に挿入

# ###########################
# 要素を取得する（単数）
# ###########################
def getElement(xpath, wait_second, retries_count):
    error = ''
    for _ in range(retries_count):
        try:
            # 失敗しそうな処理
            selector = xpath
            element = WebDriverWait(driver, wait_second).until(
              EC.presence_of_element_located((By.XPATH, selector))
            )
        except Exception as e:
            # エラーメッセージを格納する
            error = e
        else:
            # 失敗しなかった場合は、ループを抜ける
            break
    else:
        # リトライが全部失敗したときの処理。エラー内容(error)や実行時間、操作中のURL、セレクタ、スクショなどを通知する。
        send_line_notify(error, LNT_FOR_ERROR)
        exit() # プログラムを強制終了する
    return element

# ###########################
# 要素を取得する（複数）
# ###########################
def getElements(xpath, wait_second, retries_count):
    error = ''
    for _ in range(retries_count):
        try:
            # 失敗しそうな処理
            selector = xpath
            elements = WebDriverWait(driver, wait_second).until(
              EC.presence_of_all_elements_located((By.XPATH, selector))
            )
        except Exception as e:
            # エラーメッセージを格納する
            error = e
        else:
            # 失敗しなかった場合は、ループを抜ける
            break
    else:
        # リトライが全部失敗したときの処理。エラー内容(error)や実行時間、操作中のURL、セレクタ、スクショなどを通知する。
        send_line_notify(error, LNT_FOR_ERROR)
        exit() # プログラムを強制終了する
    return elements

# ###########################
# 要素を取得する（タグで複数）
# ###########################
def getElementsByTag(tagname, wait_second, retries_count):
    error = ''
    for _ in range(retries_count):
        try:
            # 失敗しそうな処理
            selector = tagname
            elements = WebDriverWait(driver, wait_second).until(
              EC.presence_of_all_elements_located((By.TAG_NAME, selector))
            )
        except Exception as e:
            # エラーメッセージを格納する
            error = e
        else:
            # 失敗しなかった場合は、ループを抜ける
            break
    else:
        # リトライが全部失敗したときの処理。エラー内容(error)や実行時間、操作中のURL、セレクタ、スクショなどを通知する。
        send_line_notify(error, LNT_FOR_ERROR)
        exit() # プログラムを強制終了する
    return elements


def generateText(n, startIndex, i, month, day, dow, num):
      # テニス場を設定
      if n == 0:
          court = '深川庭球場'
      if n == 1:
          court = '潮見庭球場'
      if n == 2:
          court = '豊住庭球場'
      if n == 3:
          court = '亀戸庭球場'
      if n == 4:
          court = '東砂庭球場'
      if n == 5:
          court = '荒川・砂町庭球場'
      if n == 6:
          court = '新砂運動場'
      # スタート時間を設定
      if startIndex + i == 1:
          startTime = '08'
          endTime = '09'
      if startIndex + i == 2:
          startTime = '09'
          endTime = '10'
      if startIndex + i == 3:
          startTime = '10'
          endTime = '11'
      if startIndex + i == 4:
          startTime = '11'
          endTime = '12'
      if startIndex + i == 5:
          startTime = '12'
          endTime = '13'
      if startIndex + i == 6:
          startTime = '13'
          endTime = '14'
      if startIndex + i == 7:
          startTime = '14'
          endTime = '15'
      if startIndex + i == 8:
          startTime = '15'
          endTime = '16'
      if startIndex + i == 9:
          startTime = '16'
          endTime = '17'

      # ソート用に月と日を2桁にする
      if int(month) < 10:
          month = '0' + str(month)
      if int(day) < 10:
          day = '0' + str(day)

      return f'{month}/{day}（{dow}）_{startTime}:00〜{endTime}:00 @{court}{num}面'




def main():

    # 処理時間計測①：開始
    start_time = time.perf_counter()

    hour = int(datetime.now().strftime('%H'))
    print(hour)
    if hour >= 0 and hour < 7:
        print('メンテナンス中')
        # 処理時間計測②：修了
        end_time = time.perf_counter()
        # 経過時間を出力(秒)
        elapsed_time = end_time - start_time
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        worksheet.update_acell('A2', now)
        worksheet.update_acell('B2', elapsed_time)
        driver.delete_all_cookies()
        driver.close()
        driver.quit()
        exit()
    else:
        print('続行')

    for x in range(3):  # 最大3回実行
        try:
            # 「江東区スポーツネット」のURL
            url1 = 'https://yoyaku.koto-sports.net'

            driver.get(url1)

            a1 = getElement('/html/body/div/center/table/tbody/tr[3]/td[1]/ul/p[1]/a', 20, 3)
            a1.click()
            a2 = getElement('//*[@id="contents"]/ul[1]/li[2]/dl/dt/form/input[2]', 20, 3)
            a2.click()
            a3 = getElement('//*[@id="local-navigation"]/dd/ul/li[1]/a', 20, 3)
            a3.click()

            # 屋外を選択
            btn1 = getElement('//*[@id="contents"]/form[1]/div/div/div/div[1]/div/div[1]/select[1]/option[2]', 20, 3)
            btn1.click()
            btn1_confirm = getElement('//*[@id="contents"]/form[1]/div/div/div/div[1]/div/div[1]/input[2]', 20, 3)
            btn1_confirm.click()

            # 硬式テニスを選択
            btn2 = getElement('//*[@id="contents"]/form[4]/div/div/div/div[1]/div/div[1]/select/option[2]', 20, 3)
            btn2.click()
            btn2_confirm = getElement('//*[@id="contents"]/form[4]/div/div/div/div[1]/div/div[1]/input[2]', 20, 3)
            btn2_confirm.click()

            # コートを全選択
            btn3 = getElement('//*[@id="contents"]/form[5]/div/div/div/input[3]', 20, 3)
            btn3.click()
            btn3_confirm = getElement('//*[@id="contents"]/form[5]/div/div/div/input[2]', 20, 3)
            btn3_confirm.click()

            if len(SUN) > 1:
                sun = getElement('//*[@id="contents"]/form[6]/div/div/dl/dd[2]/input[1]', 20, 3)
                sun.click()
            if len(MON) > 1:
                mon = getElement('//*[@id="contents"]/form[6]/div/div/dl/dd[2]/input[3]', 20, 3)
                mon.click()
            if len(TUE) > 1:
                tue = getElement('//*[@id="contents"]/form[6]/div/div/dl/dd[2]/input[5]', 20, 3)
                tue.click()
            if len(WED) > 1:
                wed = getElement('//*[@id="contents"]/form[6]/div/div/dl/dd[2]/input[7]', 20, 3)
                wed.click()
            if len(THU) > 1:
                thu = getElement('//*[@id="contents"]/form[6]/div/div/dl/dd[2]/input[9]', 20, 3)
                thu.click()
            if len(FRI) > 1:
                fri = getElement('//*[@id="contents"]/form[6]/div/div/dl/dd[2]/input[11]', 20, 3)
                fri.click()
            if len(SAT) > 1:
                sat = getElement('//*[@id="contents"]/form[6]/div/div/dl/dd[2]/input[13]', 20, 3)
                sat.click()
            if len(HOL) > 1:
                hol = getElement('//*[@id="contents"]/form[6]/div/div/dl/dd[2]/input[15]', 20, 3)
                hol.click()             

            btn4 = getElement('//*[@id="btnOK"]', 20, 3)
            btn4.click()

            # ###################################
            # 取得終了日(月と日を計算)
            # ###################################
            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            today = datetime.today()
            today_year =  int(datetime.strftime(today, '%Y'))
            print(today_year)

            today_month = datetime.strftime(today, '%m')
            if today_month[0] == '0':
              today_month = today_month[1:2]
            today_month = int(today_month)
            print(today_month)

            today_day = datetime.strftime(today, '%d')
            if today_day[0] == '0':
              today_day = today_day[1:2]
            today_day = int(today_day)
            print(today_day)

            if today_day < 10:
                print('今月まで')
                until_month = today_month
                until_day = calendar.monthrange(today_year, until_month)[1]
            else:
                print('来月まで')
                if today_month == 12:
                    until_month = 1
                else:
                    until_month = today_month + 1
                until_day = calendar.monthrange(today_year, until_month)[1]

            print('ループ開始')

            while True:
                #日付：令和04年03月10日(木)
                date = getElement('//*[@id="contents"]/div[2]/div/h3/span', 20, 3).text
                year = int(date[2:4]) + 2018
                month = int(date[5:7])
                day = int(date[8:10])
                dow = date[12:13]

                print(date)

                # 範囲ないか計算
                if  until_month < month:
                  if until_month == 1 and month != 2:
                      print('年こえる')
                  else:
                      break

                if dow == '日':
                    start = int(SUN.split('-')[0])
                    end = int(SUN.split('-')[1])
                if dow == '月':
                    start = int(MON.split('-')[0])
                    end = int(MON.split('-')[1])
                if dow == '火':
                    start = int(TUE.split('-')[0])
                    end = int(TUE.split('-')[1])
                if dow == '水':
                    start = int(WED.split('-')[0])
                    end = int(WED.split('-')[1])
                if dow == '木':
                    start = int(THU.split('-')[0])
                    end = int(THU.split('-')[1])
                if dow == '金':
                    start = int(FRI.split('-')[0])
                    end = int(FRI.split('-')[1])
                if dow == '土':
                    start = int(SAT.split('-')[0])
                    end = int(SAT.split('-')[1])

                # 検索する時間帯のテーブル・インデックスを算出
                startIndex = start - 7
                if startIndex < 1:
                    startIndex = 1
                endIndex = startIndex + (end - start) - 1
                if endIndex > 9:
                    endIndex = 9

                # テーブルのセルの情報を取得
                for n in range(7):
                    for i2 in range(endIndex - startIndex + 1):
                        checkRows = len(getElementsByTag('thead', 20, 3))

                        if checkRows == 2:
                            # 最後じゃなかったら
                            if i2 + 1 != endIndex - startIndex + 1:
                                cell_text = getElement(f'//*[@id="td1{n + 1}_{startIndex + i2}"]', 20, 3).text
                                if cell_text != 'Ｘ':
                                    img = getElement(f'//*[@id="td1{n + 1}_{startIndex + i2}"]/a/img', 20, 3)
                                    num = img.get_attribute('src')[-5:-4] # 空きコート数
                                    r = generateText(n, startIndex, i2, month, day, dow, num)
                                    result.append(r)
                            # 最後だったら（スクロールが必要だから取得方法が違う）
                            else:
                                cell_text = getElement(f'//*[@id="td1{n + 1}_{startIndex + i2}"]', 20, 3).get_attribute("textContent")
                                if cell_text != 'Ｘ':
                                    img = getElement(f'//*[@id="td1{n + 1}_{startIndex + i2}"]/a/img', 20, 3)
                                    num = img.get_attribute('src')[-5:-4] # 空きコート数
                                    r = generateText(n, startIndex, i2, month, day, dow, num)
                                    result.append(r)
                        else:
                            # 要修正
                            # 最後じゃなかったら
                            tyousei = 0
                            if n == 1:
                                tyousei = 1
                            elif n > 1:
                                tyousei = 2
                            if i2 + 1 != endIndex - startIndex + 1:
                                cell_text = getElement(f'//*[@id="td1{n + tyousei + 1}_{startIndex + i2}"]', 20, 3).text
                                if cell_text != 'Ｘ':
                                    img = getElement(f'//*[@id="td1{n + tyousei + 1}_{startIndex + i2}"]/a/img', 20, 3)
                                    num = img.get_attribute('src')[-5:-4] # 空きコート数
                                    r = generateText(n, startIndex, i2, month, day, dow, num)
                                    result.append(r)
                            # 最後だったら（スクロールが必要だから取得方法が違う）
                            else:
                                cell_text = getElement(f'//*[@id="td1{n + tyousei + 1}_{startIndex + i2}"]', 20, 3).get_attribute("textContent")
                                if cell_text != 'Ｘ':
                                    img = getElement(f'//*[@id="td1{n + tyousei + 1}_{startIndex + i2}"]/a/img', 20, 3)
                                    num = img.get_attribute('src')[-5:-4] # 空きコート数
                                    r = generateText(n, startIndex, i2, month, day, dow, num)
                                    result.append(r)

                next_btn = getElements(f'//*[@id="contents"]/div[2]/div/ul/li[2]/a[1]', 20, 3)
                if len(next_btn) > 0:
                    next_btn_style = next_btn[0].get_attribute("style")
                    if len(next_btn_style) > 0:
                        break
                    else:
                        next_btn[0].click()
                else:
                    break
                # ループここまで

            # driver.quit()

            print(result)
            # 結果をフォーマット
            result.sort()
            date1 = ''
            final_result = []

            for item in result:
                date2 = item.split('_')[0]
                others = item.split('_')[1]
                if others[0] == '0':
                    others = others[1:]
                if date1 == date2:
                    final_result.append(others)
                else:
                    final_result.append(date2)
                    final_result.append(others)
                    date1 = date2

            # spreadsheetの1業目を取得（これとfinal resultを比較！）
            history = worksheet.row_values(1)
            print('取得結果')
            print(final_result)
            print('前回の取得結果')
            print(history)

            # 変更チェック
            if history == final_result:
                print('前回と変更なし')
            # 空きが出た場合
            elif len(history) <= len(final_result):
                print('コート増えた：通知あり')
                writeSheet(final_result)

                # #######################################
                # LINE通知 ###############################
                # #######################################
                result_length = len(final_result) # 結果数
                pages = result_length // 40 + 1    # LINEの通数

                for page in range(pages):
                    if page == 0:
                        message = '\n【テニスコート空き状況：江東区】\n'
                    else:
                        message = '\n\n'

                    m = final_result[40 * page: 40 * (page + 1)]
                    for item2 in m:
                        #日付なら改行入れる
                        if '〜' not in item2:
                            message += '\n'
                        message += f'{item2}\n'
                    
                    send_line_notify(message, TOKEN)
                    send_line_notify(message, LNT_FOR_ERROR)


            # コートに空きがない場合
            elif len(final_result) == 0:
                print('空きなし：通知あり')
                # worksheet.delete_rows(0) # 1行目(空き状況)を削除
                writeSheet(final_result)
                send_line_notify('空きコートはありません。', TOKEN)
                send_line_notify('空きコートはありません。', LNT_FOR_ERROR)
            # コートが埋まった場合（通知なし）
            else:
                print('コートへった：通知なし')
                writeSheet(final_result)

            # 処理時間計測②：修了
            end_time = time.perf_counter()
            # 経過時間を出力(秒)
            elapsed_time = end_time - start_time
            worksheet.update_acell('A2', now)
            worksheet.update_acell('B2', elapsed_time)
            print(elapsed_time)
            # worksheet.append_row() # 時間を最終行(2行目)に挿入

        except Exception as e:
            # import traceback
            # traceback.print_exc()
            print(e)
            # driver.quit()
            if i == 2:
                err_title = e.__class__.__name__ # エラータイトル
                message = f'\n【江東】\n\n例外発生！\n\n{err_title}\n{e.args}'
                send_line_notify(message, LNT_FOR_ERROR)
            pass

        else: # 例外が発生しなかった時だけ実行される
            break  # 失敗しなかった時はループを抜ける



if __name__ == "__main__":
    lockfilePath = 'lockfile2.lock'
    with open(lockfilePath , "w") as lockFile:
        try:
            fcntl.flock(lockFile, fcntl.LOCK_EX | fcntl.LOCK_NB)

            # Headless Chromeをあらゆる環境で起動させるオプション
            # 省メモリ化しないとメモリ不足でクラッシュする
            options = Options()
            options.add_argument('--headless')
            options.add_argument('--disable-gpu')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            # options.add_argument('--remote-debugging-port=9222')
            # さくらでchrome not reachableエラー出たから追加
            # options.add_argument("--single-process") 
            # options.add_argument("--disable-setuid-sandbox") 
            # options.add_argument('window-size=500,500')
            # UA指定しないとTikTok弾かれる
            UA = 'SetEnvIfNoCase User-Agent "NaverBot" getout'
            options.add_argument(f'user-agent={UA}')

            # driver = webdriver.Chrome(executable_path=ChromeDriverManager().install(), options=options)
            chrome_service = service.Service(executable_path=ChromeDriverManager().install())
            driver = webdriver.Chrome(service=chrome_service, options=options) 
            driver.implicitly_wait(60)
            main()
            driver.delete_all_cookies()
            driver.close()
            driver.quit()

        except IOError:
            print('process already exists')