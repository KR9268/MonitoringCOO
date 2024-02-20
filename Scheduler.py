import time
import MonitoringCOO_crawler

while True:
    # 상공회의소 근무시간 내에만 작업 수행
    if time.localtime().tm_hour >= 9 and time.localtime().tm_hour <= 17: 
        MonitoringCOO_crawler.main_crawler()

        time.sleep(300) # 전체 작업간 주기 5분