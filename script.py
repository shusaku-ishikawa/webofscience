from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.common.exceptions import TimeoutException

from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from time import sleep 
import os, sys, datetime

''' Chromeドライバの初期設定 '''
def init_selenium(download_dir):
    ###Chromeへオプションを設定
    chop = webdriver.ChromeOptions() #
    chop.add_argument('--ignore-certificate-errors') #SSLエラー対策
    chop.add_argument('--headless')
    chop.add_argument('--no-sandbox')
    chop.add_argument('--disable-gpu')
    chop.add_argument('--window-size=1280,1024')
    chop.add_argument('log-level=3')
    driver = webdriver.Chrome(options = chop)
    driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
    driver.execute("send_command", {
        'cmd': 'Page.setDownloadBehavior',
        'params': {
            'behavior': 'allow',
            'downloadPath': download_dir # ダウンロード先
        }
    })
    driver.implicitly_wait(5) # seconds
    driver.delete_all_cookies()
    return driver

''' 検索画面のURL '''
BASE_URL = "https://apps.webofknowledge.com/WOS_GeneralSearch_input.do?product=WOS&search_mode=GeneralSearch&SID=F4r7VX5bnz9rUVkoygm&preferencesSaved="

''' １回のダウンロード件数 '''
MAX_REC = 500

''' 1回のダウンロードの待ち時間 '''
WAIT_SEC = 15

''' ファイル名 '''
DEFAULT_FILENAME = r'savedrecs.txt'

''' 以下メイン処理 '''
if __name__ == '__main__':
    
    argvs = sys.argv
    
    download_base_dir = os.path.join(os.getcwd(), 'download')
    
    ''' ダウンロードの基本フォルダがなければ作成 '''
    if not os.path.exists(download_base_dir):
        os.mkdir(download_base_dir)
    
    ''' ダウンロードフォルダがなければ作成 '''
    download_dir = os.path.join(download_base_dir, datetime.datetime.now().strftime('%Y%m%d-%H%M%S'))
    if not os.path.exists(download_dir):
        os.mkdir(download_dir)

    driver = init_selenium(download_dir)

    driver.get(BASE_URL)

    ''' 引数で検索文字列が指定された場合はそのワードで検索 '''
    if len(argvs) > 2:
        search_bar = driver.find_element_by_id('value(input1)')
        search_bar.clear()
        ''' 第二引数を検索窓に入力 '''
        search_bar.send_keys(argvs[1])

    ''' 検索ボタンを探し、クリック '''
    search_btn = driver.find_elements_by_xpath('//*[@id="searchCell1"]/span[1]/button')[0]
    search_btn.click()

    offset = 1
    limit = MAX_REC

    ''' ループ '''
    while True:
        ''' 出力ボタンを探し、クリック。その後ダイアログが出るまで待つ(タイムアウトした場合は検索結果なしと判断) '''
        try:
            export_btn = driver.find_elements_by_xpath('//*[@id="page"]/div[1]/div[26]/div[2]/div/div/div/div[2]/div[3]/div[3]/div[2]/button')[0]
        except Exception:
            print('検索結果がありませんでした。処理を終了します。')
            sys.exit()

        export_btn.click()
        export_dialog = driver.find_elements_by_xpath('//*[@id="page"]/div[11]')[0]

        ''' ダウンロード件数のチェックボックスを探し、クリック '''
        checkbox = driver.find_element_by_id('numberOfRecordsRange')
        checkbox.click()

        ''' 何件目から何件分の部分セット '''
        markfrom = driver.find_element_by_id('markFrom')
        markto = driver.find_element_by_id('markTo')
        
        markfrom.clear()
        markto.clear()
        markfrom.send_keys(offset)
        markto.send_keys(limit + offset - 1)

        ''' コンテンツ部分を選択 '''
        content = driver.find_element_by_id('bib_fields')
        content_select = Select(content)
        content_select.select_by_value('HIGHLY_CITED HOT_PAPER OPEN_ACCESS PMID USAGEIND AUTHORSIDENTIFIERS ACCESSION_NUM FUNDING SUBJECT_CATEGORY JCR_CATEGORY LANG IDS PAGEC SABBR CITREFC ISSN PUBINFO KEYWORDS CITTIMES ADDRS CONFERENCE_SPONSORS DOCTYPE CITREF ABSTRACT CONFERENCE_INFO SOURCE TITLE AUTHORS  ')

        ''' ファイル形式を選択 '''
        saveoption = driver.find_element_by_id('saveOptions')
        saveoption_select = Select(saveoption)
        saveoption_select.select_by_value('tabWinUTF8')

        ''' この内容でダウンロード開始 '''
        export_execute_btn = driver.find_element_by_id('exportButton')
        export_execute_btn.click()

        try:
            error_element = WebDriverWait(driver, 1).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="page"]/div[11]/div[2]/form/div[1]/div/div[2]/label'))
            )
            ''' 出力時にエラーが表示された場合は処理を終了する '''
            print(str(offset) + 'から' + str(limit + offset - 1) + '件目までの取得に失敗しました。残りの件数が' + str(limit) + '件未満の可能性があります。' )
            sys.exit()
        except TimeoutException:
            pass

        print(str(offset) + '―' + str(offset + limit - 1) + ' をダウンロードしています。。')
        
        ''' ファイルダウンロード数ち '''
        downloaded_file = os.path.join(download_dir, DEFAULT_FILENAME)
        while True:
            if not os.path.exists(downloaded_file):
                sleep(1)
            else:
                break

        ''' ダイアログを閉じるボタンが現れるまで待つ '''
        close_dialog_btn = driver.find_elements_by_xpath('//*[@id="page"]/div[11]/div[2]/form/div[2]/a')[0]
        
        ''' ダウンロードしたファイルが０バイトの場合は削除し、ループ抜ける '''
        if os.path.getsize(downloaded_file) == 0:
            os.remove(downloaded_file)
            break
        ''' ダウンロードしたファイルに中身がある場合はリネーム '''
        os.rename(downloaded_file, os.path.join(download_dir, str(offset) + '-' + str(offset + limit - 1) + '.txt' ))
        offset = offset + limit

        ''' ダイアログを閉じる '''
        close_dialog_btn.click()
    ''' ドライバを閉じる '''
    driver.close()

