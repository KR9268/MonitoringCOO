# MonitoringCOO

## 개요
- COO 신청내역에 대해 각각의 정보를 대신 스크레핑하여
  - 1개 화면에서 일괄조회
  - 실적 다운로드
  - 등록한 서류번호에 대한 윈도우 TOAST 알림
- 작업용 자동로그인창 제공

## 필수사항
- 크롬 및 selenium 드라이버
  2024.02.20기준 드라이버 다운로드 사이트 https://chromedriver.chromium.org/downloads
- 모든 파일을 하나의 폴더에 두고,
  - MonitoringCOO_crawler.json 파일에 스크레핑을 위한 id 및 pw 저장
  - 스크레핑에 사용할 인증서는 사전 설치
  - db저장주소를 확인하여, db를 참조할 파일들의 주소를 수정한다

## 사용법
- main 사이트(현황확인 및 실적다운로드, Push알람 및 로그인정보 관리) : cmd 명령어 실행 → streamlit run .\MonitoringCOO\MonitoringCOO.py
  1) 실적다운로드 : View and alarm(첫번째 탭) 우측 하단 '월 발급완료건 다운로드' 버튼
     (연, 월, 저장경로를 수정하여 다른 월의 자료를 지정한 자료명으로 저장가능, 전월실적을 기본값으로 자동 로딩함)
  2) Push 알람 : : View and alarm(첫번째 탭) 좌측에 id(사용자이름)와 번호를 입력해두면,
     - '발급완료' 또는 '오류통보'인 경우 Push알림 후 알림리스트에서 삭제
     - '접수완료'는 대기
  3) 현황확인 :  View and alarm(첫번째 탭) 우측에서 표를 직접 보거나 번호로 검색하여 확인
  4) 로그인정보 관리 : Setting Login info 에서 수정이 필요한 경우, 수정 후 반영버튼 클릭
  5) 
- Crawler(스크레핑용) 파일은 같은 PC, 또는 다른 PC나 서버에서 실행
  1) MonitoringCOO_crawler.py 를 실행하여 수동 스크레핑
  2) Scheduler.py 를 실행해두어 주기적으로 스크레핑 수행
  
## 특이사항
- 특이사항 (db기준)
  1) 접수번호가 유일한 값이라 Primary Key로 지정함
  2) 각 컬럼 데이터형
     접수번호 varchar PRIMARY KEY , 
     증명서종류 varchar, 
     대표Invoice varchar(10), 
     접수일시 datetime, 
     처리상태 varchar, 
     Remark varchar
- 특이사항2 (로그인창에서 멈추는 경우)
  1) 로그아웃 해두는 경우 이미지 인식이 되지 않음. Safemode해제시 대응 가능하나 Risk 등을 우려하여 Safemode해제는 하지 않음

## 기능
- find_window_by_name : selenium이 아닌 윈도우 인식용
- def close_all_except_original_window : 불필요한 팝업 제거용
- login_gongdong_byimg : 이미지 인식
- open_korcham, open_korcham2 : 스크레핑 및 작업창 등 작업시작시에 사용
- coo_crawler : 각 창에서 정보 조회 후 서류 1건을 리스트형으로 반환
- manage_db : db에 추가(INSERT)하거나, 상태가 수정된 경우 반영(REPLACE)한다
- main_crawler : 위의 기능을 통합하여 main기능을 수행
