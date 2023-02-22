pyinstaller 터미널 명령어
- pyinstaller -w -F --icon=test.ico --onefile pdf_merge.py
    · 하나의 파일
    · 아이콘은 test.ico 이용

# 모듈
    ## pptx 모듈
   - pip install aspose.slides


## 추가 프로그램
- pdf -> jpg 완료 (pdf_to_jpg.py) 2023.02.14 완료
- pptx -> pdf
- 

# converter.py
- 다른 파일들에서 테스트를 하고 통합 버전은 여기에 만들 계획

# 2023.02.22
- openpyxl 기존 버전 강제로 지우고 3.1.0으로 재설치 (아래 명령어 참조)
    pip install --force-reinstall -v "openpyxl==3.1.0"

# 2023.02.23
- 추가할 부분
    . 함수 제거는 다시 생각해보아야 함(1,2번 시트)
    . 병합에 대한 부분도 확인
    . 3, 4번 시트 보기 좋게 변경