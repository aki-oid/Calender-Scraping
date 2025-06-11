import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager  # 自動更新用

import csv
from bs4 import BeautifulSoup

# Chromeのオプション設定
chrome_options = Options()
chrome_options.add_argument('--user-data-dir=C:/Users/hataa/chromedriver_win32/chromedriver-win32')
chrome_options.add_argument('--profile-directory=Default')

# ChromeDriverのサービスを自動更新
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

url = 'https://login.yahoo.co.jp/config/login?.src=yc&.done=https%3A%2F%2Fcalendar.yahoo.co.jp%2F'
# 対象ページを開く
driver.get(url)
# ページのHTMLを取得
page_source = driver.page_source
soup = BeautifulSoup(page_source, 'html.parser')

title = soup.select_one('head > title')
print("タイトル：" + title.text)
if title.text != 'Yahoo!カレンダー':
	time.sleep(1)
	# ユーザー名とパスワードの入力
	#elem_username = driver.find_element(By.XPATH, '//*[@id="login_handle"]')
	#elem_username.send_keys('111111111@yahoo.co.jp') #11111111←アカウント名

	time.sleep(3)
	# ログインボタンをクリック
	#elem_login_btn = driver.find_element(By.XPATH, '//*[@id="content"]/div[1]/div/form/div[1]/div[1]/div[2]/div/button')
	#elem_login_btn.click()

	time.sleep(3)
	#elem_login_btn2 = driver.find_element(By.XPATH, '//*[@id="content"]/div[1]/div/form/div[2]/div/div[2]/div[2]/div[3]/button')
	#elem_login_btn2.click()

	#PINの入力
	time.sleep(15)

# ページがロードされるのを待つ
WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, 'rbc-month-view')))

url2 = driver.current_url#現在のurlをキャッチ
print(f"現在のURLは: {url2}")

# ページのHTMLを取得
page_source = driver.page_source
soup = BeautifulSoup(page_source, 'html.parser')

####################################	関数

def back_btn():
	back_btn = driver.find_element(By.XPATH, '//*[@id="bcHeader"]/tbody/tr/td/div/div[1]')
	back_btn.click()

def future_btn():
	future_btn = driver.find_element(By.XPATH, '//*[@id="bcHeader"]/tbody/tr/td/div/div[3]')
	future_btn.click()

def close():
	try:
		pop_close = driver.find_element(By.XPATH, '//*[@id="read-event-popup"]/div/div/section/div[3]/button')
		pop_close.click()
		WebDriverWait(driver, 20).until(EC.invisibility_of_element_located((By.ID, "read-event-popup")))
	except Exception as e:
		print(f"ポップアップを閉じる際にエラーが発生しました: ")
		error()
		close()

def error():
	Count_error = Count_error + 1
	time.sleep(1)

####################################	ループ定義

time.sleep(3)
print('いつから繰り返す(年)：')
inp_pas_year = input()#入力
print('いつから繰り返す(月)：')
inp_pas_month = input()#入力
Range_past = inp_pas_year + '年' + inp_pas_month + '月'

print('いつまで繰り返す(年)：')
inp_fut_year = input()#入力
print('いつまで繰り返す(月)：')
inp_fut_month = input()#入力
Range_future = inp_fut_year + '年' + inp_fut_month + '月'

judg_past = 0
judg_future = 0

Count_error = 0
stanp_list = []
####################################書き込み

HEADER = ['Subject','Start Date','Start Time','End Date','End Time','ALL Day Event','Description','Location','Private']
with open('Calender(Y_to_G).csv', 'w', encoding='utf-8') as f: #Shift JIS
	writer = csv.writer(f)
	writer.writerow(HEADER)

	today = driver.find_element(By.XPATH, '//*[@id="bcHeader"]/tbody/tr/td/div').text
	
	y = today.find('年')
	today_year = today[:y]
	today_month = today[y+1:today.find('月')]

	while 1:

		Cal = driver.find_element(By.XPATH, '//*[@id="bcHeader"]/tbody/tr/td/div').text
		
		roop_judge = Cal[:Cal.find('(')]
		y = Cal.find('年')
		year = Cal[:y]
		month = Cal[y+1:Cal.find('月')]

		# Seleniumを使用して動的要素を取得
		try:
			
			# 各週の情報を取得
			weeks = driver.find_elements(By.CSS_SELECTOR, 'div.rbc-addons-dnd-row-body')
		
			for week in weeks:
				# 各週の日の情報を取得
				days = week.find_elements(By.CSS_SELECTOR, 'div.rbc-row')

				for day in days:
					# 各日の情報を取得して処理
					day_events = day.find_elements(By.CSS_SELECTOR, 'div.rbc-row-segment')

					if day_events:
						for event in day_events:
							# イベントの詳細を取得して処理
							event_title = event.text.strip()  # 例: イベント名を表示

							if event_title:
								#print(str(event_title))#####################################################################

################################ポップ

								try:
									# イベントをクリックする
									event.click()#ポップを表示するまで待機
									WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "read-event-popup")))
									
									# 再度ページソースを取得して、ポップアップの内容を取得
									page_source = driver.page_source
									soup = BeautifulSoup(page_source, 'html.parser')

									event_pop = soup.find('div', class_='yc-menu-box yc-pop-menu')

############日付時間
									date_all = event_pop.find('div', class_='pop-wrapper pop-time').text

									m1 = date_all.find('月')
									month1 = date_all[:m1]

									if month1 == month:
############件名
										Subject = event_pop.find('h2', class_='pop-title-wrapper is-red').text
										stanp = event_title.replace(Subject,'')
										
										if Subject == '予定':
											Subject = ''
										if len(stanp)>0:
											if ':' in stanp:
												st = stanp.find(':')
												stanp = stanp[st+3:]
											if '〜' in stanp:
												st = stanp.find('〜')
												stanp = stanp[st+1:]
											if stanp_list.count(stanp) == 0:
												stanp_list.append(stanp)
										elif stanp_list.count(Subject) == 1:
											stanp = Subject
											Subject = ''
												
										if len(stanp) > 0:
											if stanp in Subject:
												Subject = Subject.replace(stanp,'')
											if stanp == '誕生日':
												stanp = '誕'
											elif stanp == 'テレビ':
												stanp = 'TV'
											elif stanp == 'バイト':
												stanp = 'バ'
											elif stanp == '大学':
												stanp = '学'
											elif stanp == '部活':
												stanp = '部'
											

											Subject = '[' + stanp + ']' + Subject

										print('\n予定：'+Subject)

######日付
										day1 = date_all[m1+1:date_all.find('日')]
										Start_Date = str(year + '/' + month1 +'/'+day1)

										if date_all.count(') ') == 2:  #2日以上
											m2 = date_all.find('〜 ')
											date_all_back = date_all[m2+2:]

											m2 = date_all_back.find('月')
											month2 = date_all_back[:m2]
											day2 = date_all_back[m2+1:date_all_back.find('日')]
											End_Date = str(year + '/' + month2 +'/'+day2)
											print(Start_Date+'~'+End_Date)

										else:
											End_Date = Start_Date
											print(Start_Date)

######時間
										if "終日" in date_all:#終日
											ALL_Day_Event = "TRUE"
											
											Start_Time = ''
											End_Time = ''
										else :
											ALL_Day_Event = 'FALSE'

											if date_all.count(') ') == 2:
												t1 = date_all_back.find(' ')
												Time_all = date_all_back[t1+1:]
											else:
												t1 = date_all.find(' ')
												Time_all = date_all[t1+1:]

											h1 = Time_all.find(':')
											hour1 = Time_all[:h1]
											min1 = Time_all[h1+1:Time_all.find(' ')]
											Start_Time = str(hour1 +':'+min1)

											t2 = Time_all.find('〜 ')
											Time_all_back = Time_all[t2+2:]
											
											h2 = Time_all_back.find(':')
											hour2 = Time_all_back[:h2]
											min2 = Time_all_back[h2+1:Time_all.find(' ')]
											End_Time = str(hour2 +':'+min2)
											print(Start_Time+'~'+End_Time)

############詳細
										if event_pop.find('p', class_='pop-wrapper pop-detail'):#詳細
											Description = event_pop.find('p', class_='pop-wrapper pop-detail').text
											print(Description)
										else :
											Description = ''

############場所
										if event_pop.find('p', class_='pop-wrapper pop-place'):#場所
											Location = event_pop.find('p', class_='pop-wrapper pop-place').text
											print(Location)
										else :
											Location = ''

############公開
										Private = 'TRUE'#非公開

############書き込み
										try:
											row = [Subject, Start_Date, Start_Time, End_Date, End_Time, ALL_Day_Event, Description, Location, Private]
											writer.writerow(row)
										except:
											print(f"書き込み処理が失敗しました")
											error()
									close()

								except Exception as e:
									print(f"エラーが発生しましたaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa")
									close()
									error()
		except Exception as e:
			print(f"月間ビューの取得中にエラーが発生しました")
			error()

##################################	ループ条件分岐

		if Range_future == roop_judge:#未来完了
			print('〇〇〇未来範囲発見〇〇〇')
			judg_future = 1

		if Range_past == roop_judge:#過去完了
			print('〇〇〇過去範囲発見〇〇〇')
			judg_past = 1
			'''if month != today_month:
				try:#今日の日付に戻る
					today_btn = driver.find_element(By.XPATH, '//*[@id="bcHeader"]/tbody/tr/td/div/div[4]/button')
					today_btn.click()
					WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'btn-ignore')))
				except:#今月より前に戻らない場合ボタンなし
					print('現在に戻れませんでした')
					error()'''

		if(judg_past == 0):
			back_btn()
		else:
			if(judg_future == 0):#過去が終わりの場合、未来へ進む
				future_btn()
			else:
				print('\n全て終了')
				break

		time.sleep(0.75)
driver.quit()
print('\nerror:'+str(Count_error))


####################################	csvファイルの整理

import pandas as pd

# CSVファイルの読み込み
df = pd.read_csv('Calender(Y_to_G).csv', encoding='utf-8')

# 日付の形式を統一してからソート
df['Start Date'] = pd.to_datetime(df['Start Date'], format='%Y/%m/%d')
df['End Date'] = pd.to_datetime(df['End Date'], format='%Y/%m/%d')

# 日付を考慮し、1桁の数を2桁の数よりも前に来るようにソートする
df['Start Year'] = df['Start Date'].dt.year
df['Start Month'] = df['Start Date'].dt.month
df['Start Day'] = df['Start Date'].dt.day

# 日付順にソート（まず年、次に月、そして日）
df = df.sort_values(by=['Start Year', 'Start Month', 'Start Day', 'Start Time'])

# 重複する行の削除
df = df.drop_duplicates(subset=['Subject', 'Start Date', 'Start Time', 'End Date', 'End Time'], keep='first')

# ソートに使用した一時的な列を削除
df = df.drop(columns=['Start Year', 'Start Month', 'Start Day'])

# 結果を新しいCSVファイルに保存
df.to_csv('Sorted_Calendar(Y_to_G).csv', index=False, encoding='utf-8')#UTF=8

print("\n日付を考慮し、1桁の数が2桁の前に来るようにソートした結果を 'Sorted_Calendar(Y_to_G).csv' に保存しました。")
