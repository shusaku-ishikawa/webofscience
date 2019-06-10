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
import os 

def init_selenium(download_dir):
    ###Chromeへオプションを設定
    chop = webdriver.ChromeOptions() #
    chop.add_argument('--ignore-certificate-errors') #SSLエラー対策
    chop.add_argument('--headless')
    chop.add_argument('--no-sandbox')
    chop.add_argument('--disable-gpu')
    chop.add_argument('--window-size=1280,1024')
    chop.add_argument('log-level=3')
    driver = webdriver.Chrome(chrome_options = chop)
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
 
base_url = "https://apps.webofknowledge.com/WOS_GeneralSearch_input.do?product=WOS&search_mode=GeneralSearch&SID=F3Cev3f9k2gBb9OfOPm&preferencesSaved="
download_dir =  r'C:\Users\sishikawa17\OneDrive - Nihon Unisys, Ltd\python\webofscience.git\downloads'
filename = r'savedrecs.txt'
driver = init_selenium(download_dir)

driver.get(base_url)

search_btn = driver.find_elements_by_xpath('//*[@id="searchCell1"]/span[1]/button')[0]
search_btn.click()

offset = 1
limit = 500
while True:
    export_btn = driver.find_elements_by_xpath('//*[@id="page"]/div[1]/div[26]/div[2]/div/div/div/div[2]/div[3]/div[3]/div[2]/button')[0]
    export_btn.click()
    export_dialog = driver.find_elements_by_xpath('//*[@id="page"]/div[11]')[0]

    checkbox = driver.find_element_by_id('numberOfRecordsRange')
    checkbox.click()

    markfrom = driver.find_element_by_id('markFrom')
    markto = driver.find_element_by_id('markTo')
    markfrom.clear()
    markto.clear()
    markfrom.send_keys(offset)
    markto.send_keys(limit + offset - 1)

    content = driver.find_element_by_id('bib_fields')
    content_select = Select(content)
    content_select.select_by_value('HIGHLY_CITED HOT_PAPER OPEN_ACCESS PMID USAGEIND AUTHORSIDENTIFIERS ACCESSION_NUM FUNDING SUBJECT_CATEGORY JCR_CATEGORY LANG IDS PAGEC SABBR CITREFC ISSN PUBINFO KEYWORDS CITTIMES ADDRS CONFERENCE_SPONSORS DOCTYPE CITREF ABSTRACT CONFERENCE_INFO SOURCE TITLE AUTHORS  ')

    saveoption = driver.find_element_by_id('saveOptions')
    saveoption_select = Select(saveoption)
    saveoption_select.select_by_value('tabMacUTF8')

    export_execute_btn = driver.find_element_by_id('exportButton')
    export_execute_btn.click()
    print(str(offset) + '―' + str(offset + limit - 1) + ' をダウンロードしています。。')
    sleep(10)
    


    close_dialog_btn = driver.find_elements_by_xpath('//*[@id="page"]/div[11]/div[2]/form/div[2]/a')[0]
    thefilepath = os.path.join(download_dir, filename)
    if os.path.getsize(thefilepath) == 0:
        os.remove(thefilepath)
        break
    os.rename(thefilepath, os.path.join(download_dir, str(offset) + '-' + str(offset + limit - 1) + '.txt' ))
    offset = offset + limit

    close_dialog_btn.click()
    

driver.close()