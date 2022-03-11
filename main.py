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

import config
import time
from datetime import datetime, date, timedelta
# import jpholiday

from datetime import datetime as dt
import calendar

# from ast import Or
# from urllib import request  # urllib.requestモジュールをインポート

# 自作モジュール
from modules.sendLine import send_line_notify


# # ################################
# # スプレッドシート読み込み
# # ################################
import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope =['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

spreadsheet = client.open('TennisCourtChecker') # 操作したいスプレッドシートの名前を指定する
worksheet = spreadsheet.worksheet('江東区スポーツネット') # シートを指定する

# Extract and print all of the values
# list_of_hashes = sheet.get_all_records()
# print(list_of_hashes)

# Headless Chromeをあらゆる環境で起動させるオプション
# 省メモリ化しないとメモリ不足でクラッシュする
options = Options()
options.add_argument('--headless')
options.add_argument('--disable-gpu')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--remote-debugging-port=9222')
# options.add_argument('window-size=500,500')
# UA指定しないとTikTok弾かれる
UA = 'SetEnvIfNoCase User-Agent "NaverBot" getout'
options.add_argument(f'user-agent={UA}')

result = []

def writeSheet(data):
    worksheet.delete_rows(2) # 1行目(空き状況)を削除
    worksheet.delete_rows(1) # 2行目(計測時間)を削除
    worksheet.append_row(data) # dataを最終行に挿入


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
      return f'{month}/{day}（{dow}）_{startTime}:00〜{endTime}:00 @{court}{num}面'




def main():

    # 処理時間計測①：開始
    start_time = time.perf_counter()

    for i in range(3):  # 最大3回実行
        try:
            # 「東京都スポーツ施設サービス」のURL
            url1 = 'https://yoyaku.koto-sports.net'

            driver = webdriver.Chrome(executable_path=ChromeDriverManager().install(), options=options)
            driver.implicitly_wait(60)
            driver.get(url1)

            a1 = driver.find_element(by=By.XPATH, value='/html/body/div/center/table/tbody/tr[3]/td[1]/ul/p[1]/a')
            a1.click()
            a2 = driver.find_element(by=By.XPATH, value='//*[@id="contents"]/ul[1]/li[2]/dl/dt/form/input[2]')
            a2.click()
            a3 = driver.find_element(by=By.XPATH, value='//*[@id="local-navigation"]/dd/ul/li[1]/a')
            a3.click()

            # 屋外を選択
            btn1 = driver.find_element(by=By.XPATH, value='//*[@id="contents"]/form[1]/div/div/div/div[1]/div/div[1]/select[1]/option[2]')
            btn1.click()

            btn1_confirm = driver.find_element(by=By.XPATH, value='//*[@id="contents"]/form[1]/div/div/div/div[1]/div/div[1]/input[2]')
            btn1_confirm.click()

            # 硬式テニスを選択
            btn2 = driver.find_element(by=By.XPATH, value='//*[@id="contents"]/form[4]/div/div/div/div[1]/div/div[1]/select/option[2]')
            btn2.click()

            btn2_confirm = driver.find_element(by=By.XPATH, value='//*[@id="contents"]/form[4]/div/div/div/div[1]/div/div[1]/input[2]')
            btn2_confirm.click()

            # コートを全選択
            btn3 = driver.find_element(by=By.XPATH, value='//*[@id="contents"]/form[5]/div/div/div/input[3]')
            btn3.click()
            btn3_confirm = driver.find_element(by=By.XPATH, value='//*[@id="contents"]/form[5]/div/div/div/input[2]')
            btn3_confirm.click()

            if len(config.SUN) > 0:
                sun = driver.find_element(by=By.XPATH, value='//*[@id="contents"]/form[6]/div/div/dl/dd[2]/input[1]')
                sun.click()
            if len(config.MON) > 0:
                mon = driver.find_element(by=By.XPATH, value='//*[@id="contents"]/form[6]/div/div/dl/dd[2]/input[3]')
                mon.click()
            if len(config.TUE) > 0:
                tue = driver.find_element(by=By.XPATH, value='//*[@id="contents"]/form[6]/div/div/dl/dd[2]/input[5]')
                tue.click()
            if len(config.WED) > 0:
                wed = driver.find_element(by=By.XPATH, value='//*[@id="contents"]/form[6]/div/div/dl/dd[2]/input[7]')
                wed.click()
            if len(config.THU) > 0:
                thu = driver.find_element(by=By.XPATH, value='//*[@id="contents"]/form[6]/div/div/dl/dd[2]/input[9]')
                thu.click()
            if len(config.FRI) > 0:
                fri = driver.find_element(by=By.XPATH, value='//*[@id="contents"]/form[6]/div/div/dl/dd[2]/input[11]')
                fri.click()
            if len(config.SAT) > 0:
                sat = driver.find_element(by=By.XPATH, value='//*[@id="contents"]/form[6]/div/div/dl/dd[2]/input[13]')
                sat.click()
            if len(config.HOL) > 0:
                hol = driver.find_element(by=By.XPATH, value='//*[@id="contents"]/form[6]/div/div/dl/dd[2]/input[15]')
                hol.click()             

            btn4 = driver.find_element(by=By.XPATH, value='//*[@id="btnOK"]')
            btn4.click()

            # ###################################
            # 取得終了日(月と日を計算)
            # ###################################
            today = datetime.today()
            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            today_year =  int('20' + datetime.strftime(today, '%y'))
            today_month = datetime.strftime(today, '%m')
            if today_month[0] == '0':
              today_month = today_month[1:2]
            today_month = int(today_month)
            today_day = datetime.strftime(today, '%d')
            if today_day[0] == '0':
              today_day = today_day[1:2]
            today_day = int(today_day)

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
                date = driver.find_element(by=By.XPATH, value='//*[@id="contents"]/div[2]/div/h3/span').text
                year = int(date[2:4]) + 2018
                month = int(date[5:7])
                day = int(date[8:10])
                dow = date[12:13]

                # 範囲ないか計算
                if  until_month < month:
                  if until_month == 1 and month != 2:
                      print('年こえる')
                  else:
                      break

                if dow == '日':
                    start = int(config.SUN.split('-')[0])
                    end = int(config.SUN.split('-')[1])
                if dow == '月':
                    start = int(config.MON.split('-')[0])
                    end = int(config.MON.split('-')[1])
                if dow == '火':
                    start = int(config.TUE.split('-')[0])
                    end = int(config.TUE.split('-')[1])
                if dow == '水':
                    start = int(config.WED.split('-')[0])
                    end = int(config.WED.split('-')[1])
                if dow == '木':
                    start = int(config.THU.split('-')[0])
                    end = int(config.THU.split('-')[1])
                if dow == '金':
                    start = int(config.FRI.split('-')[0])
                    end = int(config.FRI.split('-')[1])
                if dow == '土':
                    start = int(config.SAT.split('-')[0])
                    end = int(config.SAT.split('-')[1])

                # 検索する時間帯のテーブル・インデックスを算出
                startIndex = start - 7
                if startIndex < 1:
                    startIndex = 1
                endIndex = startIndex + (end - start) - 1
                if endIndex > 9:
                    endIndex = 9

                # テーブルのセルの情報を取得
                for n in range(7):
                    for i in range(endIndex - startIndex + 1):
                        checkRows = len(driver.find_elements(by=By.TAG_NAME, value="thead"))
                        if checkRows == 2:
                            if i + 1 != endIndex - startIndex + 1:
                                cell_text = driver.find_element(by=By.XPATH, value=f'//*[@id="td1{n + 1}_{startIndex + i}"]').text
                                if cell_text != 'Ｘ':
                                    img = driver.find_element(by=By.XPATH, value=f'//*[@id="td1{n + 1}_{startIndex + i}"]/a/img')
                                    num = img.get_attribute('src')[-5:-4] # 空きコート数
                                    r = generateText(n, startIndex, i, month, day, dow, num)
                                    result.append(r)
                            else:
                                cell_text = driver.find_element(by=By.XPATH, value=f'//*[@id="td1{n + 1}_{startIndex + i}"]').get_attribute("textContent")
                                if cell_text != 'Ｘ':
                                    img = driver.find_element(by=By.XPATH, value=f'//*[@id="td1{n + 1}_{startIndex + i}"]/a/img')
                                    num = img.get_attribute('src')[-5:-4] # 空きコート数
                                    r = generateText(n, startIndex, i, month, day, dow, num)
                                    result.append(r)
                        else:
                            # 要修正
                            print('フォーマット違う')
                
                next_btn = driver.find_elements(by=By.XPATH, value=f'//*[@id="contents"]/div[2]/div/ul/li[2]/a[1]')
                if len(next_btn) > 0:
                    next_btn_style = next_btn[0].get_attribute("style")
                    if len(next_btn_style) > 0:
                        break
                    else:
                        next_btn[0].click()
                else:
                    break
                # ループここまで

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
                # LINE通知
                message = '\n【テニスコート空き状況：江東】\n'
                message2 = '\n\n'
                message3 = '\n\n'
                message4 = '\n\n'

                if len(final_result) > 120:
                    m1 = final_result[0:40]
                    for item2 in m1:
                        #日付なら改行入れる
                        if '〜' not in item2:
                            message += '\n'
                        message += f'{item2}\n'
                    send_line_notify(message)

                    m2 = final_result[40:80]
                    for item2 in m2:
                        #日付なら改行入れる
                        if '〜' not in item2:
                            message2 += '\n'
                        message2 += f'{item2}\n'
                    send_line_notify(message2)

                    m3 = final_result[80:120]
                    for item2 in m3:
                        #日付なら改行入れる
                        if '〜' not in item2:
                            message3 += '\n'
                        message3 += f'{item2}\n'
                    send_line_notify(message3)

                    m4 = final_result[120:]
                    for item2 in m4:
                        #日付なら改行入れる
                        if '〜' not in item2:
                            message4 += '\n'
                        message4 += f'{item2}\n'
                    send_line_notify(message4)

                elif len(final_result) > 80:
                    m1 = final_result[0:40]
                    for item2 in m1:
                        #日付なら改行入れる
                        if '〜' not in item2:
                            message += '\n'
                        message += f'{item2}\n'
                    send_line_notify(message)

                    m2 = final_result[40:80]
                    for item2 in m2:
                        #日付なら改行入れる
                        if '〜' not in item2:
                            message2 += '\n'
                        message2 += f'{item2}\n'
                    send_line_notify(message2)

                    m3 = final_result[80:]
                    for item2 in m3:
                        #日付なら改行入れる
                        if '〜' not in item2:
                            message3 += '\n'
                        message3 += f'{item2}\n'
                    send_line_notify(message3)

                elif len(final_result) > 40:
                    m1 = final_result[0:40]
                    for item2 in m1:
                        #日付なら改行入れる
                        if '〜' not in item2:
                            message += '\n'
                        message += f'{item2}\n'
                    send_line_notify(message)

                    m2 = final_result[40:]
                    for item2 in m2:
                        #日付なら改行入れる
                        if '〜' not in item2:
                            message2 += '\n'
                        message2 += f'{item2}\n'
                    send_line_notify(message2)

                else:
                    for item2 in final_result:
                        #日付なら改行入れる
                        if '〜' not in item2:
                            message += '\n'
                        message += f'{item2}\n'
                    send_line_notify(message)

            # コートに空きがない場合
            elif len(final_result) == 0:
                print('空きなし：通知あり')
                # worksheet.delete_rows(0) # 1行目(空き状況)を削除
                writeSheet(final_result)
                send_line_notify('空きコートはありません。')
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
            if i == 2:
                err_title = e.__class__.__name__ # エラータイトル
                message = f'例外発生！\n\n{err_title}\n{e.args}'
                send_line_notify(message, config.LNT_FOR_ERROR)
            pass

        else: # 例外が発生しなかった時だけ実行される
            break  # 失敗しなかった時はループを抜ける



if __name__ == "__main__":
    main()

