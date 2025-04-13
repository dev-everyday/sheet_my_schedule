# ICS 캘린더 일정 엑셀 변환 프로그램

이 프로그램은 ICS 형식의 캘린더 일정 파일을 엑셀 파일로 변환해주는 도구입니다.

## 설치 방법

1. Python 가상환경 생성 및 활성화:
```bash
# 가상환경 생성
python -m venv venv

# 가상환경 활성화
# Windows CMD의 경우:
.\venv\Scripts\activate
# Windows PowerShell의 경우:
.\venv\Scripts\Activate.ps1
# Linux/Mac의 경우:
source venv/bin/activate
```

2. 필요한 패키지 설치:
```bash
pip install -r requirements.txt
```

## 사용 방법

1. 구글 캘린더에서 ICS 파일 다운로드:
   - 구글 캘린더 설정에서 내보내기 > ICS 파일 다운로드
   - 다운로드한 ICS 파일을 프로그램 폴더에 복사

2. 프로그램 실행:
```bash
python ics_to_excel.py
```

3. 화면에 나타나는 안내에 따라 입력:
   - Gmail 계정 이름 입력 (예: dev-everyday)
   - 변환할 연도 입력 (예: 2025)
   - 변환할 월 입력 (예: 4)

4. 변환된 엑셀 파일(`calendar_YYYY_MM.xlsx`)이 생성됩니다.

## 기능

- ICS 파일의 캘린더 일정을 월별 엑셀 달력으로 변환
- 일정 정보 포함:
  - 제목
  - 시작 시간
  - 종료 시간
  - 설명
  - 위치

## 출력 형식

- 엑셀 파일에 달력 형태로 출력 (7일 x 주차)
- 각 셀에 해당 날짜의 일정 시간과 내용 표시
- 한글 요일 표시 (일, 월, 화, 수, 목, 금, 토)

## 요구사항

- Python 3.7 이상
- 필요한 패키지:
  - icalendar==5.0.11
  - pandas==2.1.4
  - openpyxl==3.1.2

## 참고사항

- ICS 파일은 Gmail 계정 이름으로 자동 인식됩니다 (예: example@gmail.com.ics)
- 출력 엑셀 파일의 이름은 `calendar_YYYY_MM.xlsx` 형식으로 생성됩니다 