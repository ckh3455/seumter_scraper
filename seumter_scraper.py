import os
import time
import datetime
import json
import sys
import io
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Windows 환경에서 한글 출력 오류 해결 (cp1252 -> utf-8)
sys.stdout = io.TextIOWrapper(sys.stdout.detach(), encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.detach(), encoding='utf-8')

# Google Drive API
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# ====================================================================================
# 설정 (Configuration)
# ====================================================================================
BASE_DIR = os.getcwd()
EXCEL_FILENAME = "압구정동 주소.xlsx" 
EXCEL_PATH = os.path.join(BASE_DIR, EXCEL_FILENAME)
DOWNLOAD_DIR = os.path.join(BASE_DIR, "downloads")
PROCESSED_LOG_FILE = os.path.join(BASE_DIR, "processed.txt")
TARGET_URL = "https://www.eais.go.kr/"
CHUNK_SIZE = 50 

# Google Drive 설정
SCOPES = ['https://www.googleapis.com/auth/drive.file']

def log(msg):
    """타임스탬프와 함께 로그 출력"""
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{now}] {msg}")

def upload_to_drive(file_path, folder_id, credentials_json):
    """Google Drive로 파일 업로드"""
    try:
        creds_dict = json.loads(credentials_json)
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        service = build('drive', 'v3', credentials=creds)

        file_metadata = {
            'name': os.path.basename(file_path),
            'parents': [folder_id]
        }
        media = MediaFileUpload(file_path, resumable=True)
        
        file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        log(f"[Drive] 업로드 성공: {os.path.basename(file_path)} (ID: {file.get('id')})")
        return True
    except Exception as e:
        log(f"[Drive] 업로드 실패: {e}")
        return False

def main():
    log("=== 스크립트 시작 ===")
    log(f"작업 경로: {BASE_DIR}")
    log(f"다운로드 폴더: {DOWNLOAD_DIR}")

    # 0. 환경 변수 확인
    user_id = os.environ.get("SEUMTER_ID")
    user_pw = os.environ.get("SEUMTER_PW")
    
    # 구글 드라이브 관련 변수
    drive_creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    drive_folder_id = os.environ.get("GOOGLE_DRIVE_FOLDER_ID")
    
    is_github_actions = os.environ.get("GITHUB_ACTIONS") == "true"
    
    if is_github_actions:
        log("환경: GitHub Actions (Headless 모드)")
    else:
        log("환경: 로컬 PC (화면 표시 모드)")

    # 1. 파일 존재 여부 확인
    if not os.path.exists(EXCEL_PATH):
        log(f"[오류] 엑셀 파일을 찾을 수 없습니다: {EXCEL_PATH}")
        return

    # 2. 처리된 목록 로드
    processed_addrs = set()
    if os.path.exists(PROCESSED_LOG_FILE):
        with open(PROCESSED_LOG_FILE, "r", encoding="utf-8") as f:
            processed_addrs = set(line.strip() for line in f if line.strip())
    log(f"이미 처리된 주소: {len(processed_addrs)}개")

    # 3. 엑셀 로드 및 대상 선정
    try:
        df = pd.read_excel(EXCEL_PATH)
        if '주소' not in df.columns:
            log("[오류] 엑셀 파일에 '주소' 컬럼이 없습니다.")
            return
        
        all_addresses = df['주소'].dropna().unique().tolist()
        target_addresses = [addr for addr in all_addresses if addr not in processed_addrs]
        
        log(f"전체 주소: {len(all_addresses)}개, 남은 주소: {len(target_addresses)}개")
        
        if not target_addresses:
            log("[완료] 모든 주소가 처리되었습니다.")
            return

        current_chunk = target_addresses[:CHUNK_SIZE]
        log(f"이번 실행 처리 대상: {len(current_chunk)}개")

    except Exception as e:
        log(f"[오류] 엑셀 파일을 읽는 중 문제가 발생했습니다: {e}")
        return

    # 4. 브라우저 설정
    log("브라우저 설정 중...")
    options = webdriver.ChromeOptions()
    
    if is_github_actions:
        options.add_argument("--headless")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
    else:
        options.add_experimental_option("detach", True) 

    options.add_argument("--window-size=1920,1080")
    
    if not os.path.exists(DOWNLOAD_DIR):
        os.makedirs(DOWNLOAD_DIR)

    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True
    }
    options.add_experimental_option("prefs", prefs)

    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        wait = WebDriverWait(driver, 20)
    except Exception as e:
        log(f"[치명적 오류] 브라우저 실행 실패: {e}")
        return

    try:
        # 5. 세움터 접속 및 로그인
        log(f"사이트 접속 시도: {TARGET_URL}")
        driver.get(TARGET_URL)
        log("사이트 접속 완료")
        
        if user_id and user_pw:
            perform_login(driver, wait, user_id, user_pw)
        else:
            if is_github_actions:
                log("[주의] GitHub Actions 환경인데 로그인 정보가 없습니다.")
            else:
                log("[알림] 로컬 실행 중입니다. 수동으로 로그인해주세요.")
                log("로그인이 완료되고 주소 검색 준비가 되면 엔터키를 눌러주세요.")
                input(">>> 엔터키를 누르면 진행합니다...")

        # 6. 주소 반복 처리
        processed_count = 0
        
        for addr in current_chunk:
            log(f"--- [진행률 {processed_count+1}/{len(current_chunk)}] 주소 처리 시작: {addr} ---")
            
            # 다운로드 전 파일 목록 확인
            before_files = set(os.listdir(DOWNLOAD_DIR))
            
            success = process_address(driver, wait, addr)
            
            if success:
                # 다운로드 확인 및 업로드
                time.sleep(5) # 다운로드 대기
                after_files = set(os.listdir(DOWNLOAD_DIR))
                new_files = after_files - before_files
                
                if new_files:
                    log(f"다운로드된 파일: {new_files}")
                    # 구글 드라이브 업로드
                    if drive_creds_json and drive_folder_id:
                        for filename in new_files:
                            file_path = os.path.join(DOWNLOAD_DIR, filename)
                            upload_to_drive(file_path, drive_folder_id, drive_creds_json)
                    else:
                        log("[Drive] 구글 드라이브 설정이 없어 업로드를 건너뜁니다.")
                else:
                    log("[주의] 다운로드된 파일이 감지되지 않았습니다.")

                with open(PROCESSED_LOG_FILE, "a", encoding="utf-8") as f:
                    f.write(addr + "\n")
                processed_count += 1
                log(f"처리 완료 기록됨: {addr}")
            
            time.sleep(2)

    except Exception as e:
        log(f"[치명적 오류] 실행 중 예외 발생: {e}")
        if is_github_actions:
            driver.save_screenshot(os.path.join(BASE_DIR, "error_screenshot.png"))
    finally:
        log(f"작업 종료. 총 {processed_count}개 처리 완료.")
        if is_github_actions:
            driver.quit()
        else:
            log("로컬 실행이므로 브라우저를 닫지 않습니다.")

def perform_login(driver, wait, user_id, user_pw):
    log("자동 로그인 시도 중...")
    try:
        log("로그인 버튼 찾는 중...")
        login_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '로그인')] | //a[contains(text(), '로그인')]")))
        login_btn.click()
        
        log("아이디/비번 입력 중...")
        id_input = wait.until(EC.presence_of_element_located((By.ID, "id_input_id"))) 
        pw_input = driver.find_element(By.ID, "pw_input_id")
        
        id_input.clear()
        id_input.send_keys(user_id)
        pw_input.clear()
        pw_input.send_keys(user_pw)
        
        log("로그인 제출 버튼 클릭...")
        submit_btn = driver.find_element(By.ID, "login_submit_btn")
        submit_btn.click()
        
        wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(text(), '로그아웃')]")))
        log("로그인 성공!")
        
    except Exception as e:
        log(f"[오류] 로그인 실패: {e}")
        driver.save_screenshot("login_failed.png")
        raise e

def process_address(driver, wait, addr):
    try:
        log(f"검색어 입력: {addr}")
        try:
            search_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='건축물 소재지를 입력하세요.'] | //input[contains(@class, 'multiselect__input')]")))
        except TimeoutException:
            log("[오류] 검색창을 찾을 수 없습니다.")
            return False

        search_input.send_keys(Keys.CONTROL + "a")
        search_input.send_keys(Keys.DELETE)
        time.sleep(0.5)
        search_input.send_keys(addr)
        time.sleep(1)
        search_input.send_keys(Keys.ENTER)
        
        try:
            driver.find_element(By.XPATH, "//div[@id='eleasticSearch']//button").click()
        except:
            pass

        time.sleep(3)

        try:
            tab = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), '전유부')]")))
            tab.click()
            time.sleep(2)
        except:
            log("'전유부' 탭이 없거나 이미 선택됨.")

        log("그리드 스캔 및 다운로드 로직 실행 (생략됨 - 실제 구현 필요)")
        # 실제로는 여기서 다운로드 버튼을 눌러야 함
        # 테스트를 위해 가짜 파일 생성 (실제 구현 시 삭제)
        # fake_file = os.path.join(DOWNLOAD_DIR, f"{addr}.pdf")
        # with open(fake_file, "w") as f: f.write("test")
        
        return True

    except Exception as e:
        log(f"[오류] {addr} 처리 중 예외: {e}")
        return False

if __name__ == "__main__":
    main()
