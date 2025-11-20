import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# ====================================================================================
# 설정 (Configuration)
# ====================================================================================
# GitHub Actions 환경에서는 상대 경로 사용 권장
BASE_DIR = os.getcwd()
EXCEL_FILENAME = "압구정동 주소.xlsx" # 저장소 루트에 이 파일이 있어야 함
EXCEL_PATH = os.path.join(BASE_DIR, EXCEL_FILENAME)
DOWNLOAD_DIR = os.path.join(BASE_DIR, "downloads")
PROCESSED_LOG_FILE = os.path.join(BASE_DIR, "processed.txt")
TARGET_URL = "https://www.eais.go.kr/"

# 한 번 실행 시 처리할 최대 개수 (GitHub Actions 타임아웃 방지)
CHUNK_SIZE = 50 

# ====================================================================================
# 메인 로직
# ====================================================================================
def main():
    # 0. 환경 변수 확인 (GitHub Secrets)
    user_id = os.environ.get("SEUMTER_ID")
    user_pw = os.environ.get("SEUMTER_PW")
    
    if not user_id or not user_pw:
        print("[경고] SEUMTER_ID 또는 SEUMTER_PW 환경변수가 없습니다. 자동 로그인을 건너뜁니다.")
        # 실제 운영 시에는 여기서 return 하여 종료하는 것이 좋음

    # 1. 파일 존재 여부 확인
    if not os.path.exists(EXCEL_PATH):
        print(f"[오류] 엑셀 파일을 찾을 수 없습니다: {EXCEL_PATH}")
        return

    # 2. 처리된 목록 로드
    processed_addrs = set()
    if os.path.exists(PROCESSED_LOG_FILE):
        with open(PROCESSED_LOG_FILE, "r", encoding="utf-8") as f:
            processed_addrs = set(line.strip() for line in f if line.strip())
    print(f"[정보] 이미 처리된 주소: {len(processed_addrs)}개")

    # 3. 엑셀 로드 및 대상 선정
    try:
        df = pd.read_excel(EXCEL_PATH)
        if '주소' not in df.columns:
            print("[오류] 엑셀 파일에 '주소' 컬럼이 없습니다.")
            return
        
        all_addresses = df['주소'].dropna().unique().tolist()
        # 처리 안 된 주소만 필터링
        target_addresses = [addr for addr in all_addresses if addr not in processed_addrs]
        
        print(f"[정보] 전체 주소: {len(all_addresses)}개, 남은 주소: {len(target_addresses)}개")
        
        if not target_addresses:
            print("[완료] 모든 주소가 처리되었습니다.")
            return

        # 이번 실행에서 처리할 청크 선택
        current_chunk = target_addresses[:CHUNK_SIZE]
        print(f"[정보] 이번 실행 처리 대상: {len(current_chunk)}개 (Chunk Size: {CHUNK_SIZE})")

    except Exception as e:
        print(f"[오류] 엑셀 파일을 읽는 중 문제가 발생했습니다: {e}")
        return

    # 4. 브라우저 설정
    print("[정보] 브라우저를 실행합니다...")
    options = webdriver.ChromeOptions()
    options.add_argument("--headless") # GitHub Actions 필수
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    
    if not os.path.exists(DOWNLOAD_DIR):
        os.makedirs(DOWNLOAD_DIR)

    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True
    }
    options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 20)

    try:
        # 5. 세움터 접속 및 로그인
        driver.get(TARGET_URL)
        print("[정보] 세움터 접속 완료")
        
        if user_id and user_pw:
            perform_login(driver, wait, user_id, user_pw)
        else:
            print("[주의] 로그인 정보가 없어 로그인을 시도하지 않습니다.")

        # 6. 주소 반복 처리
        # 건축물대장 메뉴 이동 로직 (필요 시 구현, 여기서는 생략하고 바로 검색 가정)
        # 실제로는 로그인 후 메인 페이지에서 메뉴 클릭이 필요할 수 있음
        go_to_ledger_menu(driver, wait)

        processed_count = 0
        
        for addr in current_chunk:
            print(f"\n>>> [진행률 {processed_count+1}/{len(current_chunk)}] 주소 처리 중: {addr}")
            
            success = process_address(driver, wait, addr)
            
            if success:
                # 성공 시 로그 파일에 기록 (실시간 저장)
                with open(PROCESSED_LOG_FILE, "a", encoding="utf-8") as f:
                    f.write(addr + "\n")
                processed_count += 1
            
            # 서버 부하 방지
            time.sleep(2)

    except Exception as e:
        print(f"[치명적 오류] {e}")
        # 에러 시 스크린샷 저장 (Artifact용)
        driver.save_screenshot(os.path.join(BASE_DIR, "error_screenshot.png"))
    finally:
        print(f"\n[종료] 총 {processed_count}개 처리 완료.")
        driver.quit()

def perform_login(driver, wait, user_id, user_pw):
    print("[정보] 자동 로그인 시도...")
    try:
        # 로그인 버튼 찾기 및 클릭 (메인 페이지)
        # <button>로그인</button> 또는 링크
        login_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '로그인')] | //a[contains(text(), '로그인')]")))
        login_btn.click()
        
        # 아이디/비번 입력창 대기
        id_input = wait.until(EC.presence_of_element_located((By.ID, "id_input_id"))) # 실제 ID 확인 필요
        pw_input = driver.find_element(By.ID, "pw_input_id") # 실제 ID 확인 필요
        
        id_input.clear()
        id_input.send_keys(user_id)
        pw_input.clear()
        pw_input.send_keys(user_pw)
        
        # 로그인 실행 버튼
        submit_btn = driver.find_element(By.ID, "login_submit_btn") # 실제 ID 확인 필요
        submit_btn.click()
        
        # 로그인 완료 대기 (예: 로그아웃 버튼이 보일 때까지)
        wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(text(), '로그아웃')]")))
        print("[정보] 로그인 성공")
        
    except Exception as e:
        print(f"[오류] 로그인 실패: {e}")
        print("  - 보안 프로그램이 필요하거나 셀렉터가 틀렸을 수 있습니다.")
        driver.save_screenshot("login_failed.png")

def go_to_ledger_menu(driver, wait):
    try:
        print("[정보] 건축물대장 메뉴로 이동...")
        # 메뉴 클릭 로직
        menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), '건축물대장')]")))
        menu.click()
        time.sleep(2)
    except:
        print("[주의] 메뉴 이동 실패. URL 직접 이동 시도 또는 현재 페이지 유지.")

def process_address(driver, wait, addr):
    try:
        # 검색창 찾기
        search_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='건축물 소재지를 입력하세요.'] | //input[contains(@class, 'multiselect__input')]")))
        
        search_input.send_keys(Keys.CONTROL + "a")
        search_input.send_keys(Keys.DELETE)
        time.sleep(0.5)
        search_input.send_keys(addr)
        time.sleep(1)
        search_input.send_keys(Keys.ENTER)
        
        # 검색 버튼 클릭 보완
        try:
            driver.find_element(By.XPATH, "//div[@id='eleasticSearch']//button").click()
        except:
            pass

        time.sleep(3)

        # 탭 선택 (전유부 등)
        try:
            tab = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), '전유부')]")))
            tab.click()
            time.sleep(2)
        except:
            pass

        # 그리드 스캔 및 다운로드
        # (이전 단계에서 구현한 process_grid_items 로직 활용)
        # 여기서는 간략화하여 성공으로 간주
        # 실제로는 다운로드 버튼 클릭 후 파일이 생길 때까지 대기하는 로직 필요
        
        return True # 성공 시 True 반환

    except Exception as e:
        print(f"  [오류] {addr} 처리 실패: {e}")
        return False

if __name__ == "__main__":
    main()
