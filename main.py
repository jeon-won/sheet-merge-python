import json
import sys
from pathlib import Path
from module_xlsx import merge_xlsx

logo = r"""
 _____  _                  _                                       
/  ___|| |                | |                                      
\ `--. | |__    ___   ___ | |_  _ __ ___    ___  _ __   __ _   ___ 
 `--. \| '_ \  / _ \ / _ \| __|| '_ ` _ \  / _ \| '__| / _` | / _ \
/\__/ /| | | ||  __/|  __/| |_ | | | | | ||  __/| |   | (_| ||  __/
\____/ |_| |_| \___| \___| \__||_| |_| |_| \___||_|    \__, | \___|
                                                        __/ |      
                                                       |___/    
"""

# 설정 파일 불러오기
with open('config.json') as config_file:
    config = json.load(config_file)

# 뽀대용 로고 출력
print(logo)
print(f"v{config['version']}\n\n")

while True:
    select = int(input("# 어떤 파일을 합치겠습니까? [1. 엑셀(XLSX), 2. CSV, 0. 종료] => "))

    if select is 0:
        print("프로그램을 종료합니다.")
        sys.exit(0)
    elif select is 1:
        merge_xlsx()
    elif select is 2:
        print("아직 구현 안 함... 프로그램을 종료합니다.")
        sys.exit(0)
    else:
        print("다시 선택해주세요.")
