import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from openpyxl import Workbook

# 파일 다운로드 완료 확인 함수
def wait_for_downloads(download_folder):
    max_wait_time = 60  # 최대 대기 시간 설정 (초)
    elapsed_time = 0
    while any(file.endswith('.crdownload') for file in os.listdir(download_folder)):
        time.sleep(1)
        elapsed_time += 1
        if elapsed_time > max_wait_time:
            raise TimeoutError("파일 다운로드를 기다리는 시간이 초과되었습니다.")

# 파일 이동 함수
def move_downloaded_files(download_folder, target_folder):
    for filename in os.listdir(download_folder):
        if not filename.endswith('.crdownload'):
            os.rename(
                os.path.join(download_folder, filename),
                os.path.join(target_folder, filename)
            )

# Chrome 드라이버 설정
download_folder = r"C:\Users\변상현\Desktop\downloads"
if not os.path.exists(download_folder):
    os.makedirs(download_folder)

excel_file_path = r"C:\Users\변상현\Desktop\게시글_내용.xlsx"  # Excel 파일 경로 정의
options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": download_folder,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
options.add_experimental_option('prefs', prefs)

driver = webdriver.Chrome(options=options)
driver.get('https://www.boho.or.kr/kr/bbs/list.do?menuNo=205022&bbsId=B0000132')

# Excel 파일 준비
wb = Workbook()
ws = wb.active
ws.title = "게시글 내용"

# 게시글 목록 찾기
posts = driver.find_elements(By.CSS_SELECTOR, 'td.sbj > a')

# 게시글을 순회하면서 첨부파일이 있는 게시글을 찾기
for post in posts:
    post.click()  # 게시물 클릭
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.content')))

    # 첨부파일 다운로드 시도
    try:
        attachment_link = driver.find_element(By.CSS_SELECTOR, 'a[onclick^="fileDown"]')
        driver.execute_script(attachment_link.get_attribute('onclick'))
        wait_for_downloads(download_folder)  # 다운로드 완료 대기
        
        # 내용 추출 및 Excel에 저장
        content = driver.find_element(By.CSS_SELECTOR, 'div.content').text
        ws.append([content])

        # 다운로드 폴더에서 첨부파일을 게시글 제목 폴더로 이동
        post_title = driver.find_element(By.CSS_SELECTOR, 'div.bbs_title').text.strip().replace('/', '_')
        post_folder_path = os.path.join(download_folder, post_title)
        if not os.path.exists(post_folder_path):
            os.makedirs(post_folder_path)
        for filename in os.listdir(download_folder):
            if not filename.endswith('.crdownload'):
                os.rename(
                    os.path.join(download_folder, filename),
                    os.path.join(post_folder_path, filename)
                )

        break  # 첫 번째 첨부파일이 있는 게시글을 찾았으므로 반복 종료

    except (NoSuchElementException, TimeoutException):
        driver.back()  # 첨부파일이 없으면 뒤로 가서 다음 게시글 확인

# Excel 파일 저장
wb.save(excel_file_path)

# 드라이버 종료
driver.quit()
