# sheet-merge-python

## 개요
여러 엑셀(XLSX) 파일 또는 CSV 파일의 내용을 하나의 파일로 수합하는 파이썬 프로그램


## 필요한 것들

### Python 3
파이썬으로 돌아가니 당연히 파이썬을 설치해야 합니다. https://www.python.org/downloads 에서 Python 3를 설치합니다.

### openpyxl
openpyxl을 사용하여 수합 결과를 엑셀(XLSX) 파일로 저장합니다.
`pip install openpyxl` 명령어로 설치합니다.

### numpy
numpy는 입력받은 데이터 행이 1개 뿐인지(1차원 list 여부), 2개 이상인지(2차원 list 여부)를 파악하기 위해 사용합니다.
`pip install numpy` 명령어로 설치합니다.


## 사용법
1. `data` 폴더에 수합할 엑셀(XLSX) 파일 또는 CSV 파일을 복사합니다.
2. `python main.py` 명령어를 실행하거나 `main.bat` 파일을 실행합니다.
3. 몇 번째 행이 컬럼 이름인지 입력하면 `Result.xlsx` 또는 `Result.csv` 파일로 수합됩니다.


## 프로그램 구조

### main.py
프로그램 실행부

### config.json
프로그램 설정 값. 프로그램 버전, 수합할 파일이 담긴 폴더 이름, 시트 이름, 저장할 파일 이름 설정.

### module_xlsx.py
엑셀(XLSX) 파일 수합을 위한 함수 모음
* get_xlsx_colname(): 엑셀(XLSX) 파일의 컬럼 이름을 얻어옴
* get_xlsx_data(): 엑셀(XLSX) 파일에서 컬럼 이름을 제외한 데이터를 얻어옴
* merge_xlsx(): 특정 폴더에 존재하는 모든 엑셀(XLSX) 파일들의 수합 시도

### module_csv.py
아직 안 만듦...
