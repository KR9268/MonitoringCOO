import win32com.client
import subprocess
import time
import psutil
from typing import TypeVar, Tuple, List, Optional, Dict
import pandas as pd
from urllib import parse
from bs4 import BeautifulSoup
import xlwings as xw
import pickle
import win11toast


import NERP_PI_LC

# NERP 준비
sessions = NERP_PI_LC.check_and_open_sap('SEP', 'kibok.park', 'qwer456456*',3)
session = sessions[2]

# 기준정보
refresh_seconds = 60
send_id = ['MsgReferenceID','DocumentNumber','DocumentDate']
receive_id = ['MsgReferenceID','MsgStatus','MsgStatusDetail','DocumentDate']
file_path = 'C:\\Users\\samsung\\Downloads\\'

option_dict = {'srlist':{'xmltype':'send', 'column_id':send_id, 'path':file_path, 'filename_main':'srlist_main_df.csv','filename_temp':'srlist_temp.csv','duplicate':False},
                'status':{'xmltype':'receive','column_id':receive_id, 'path':file_path, 'filename_main':'status_main_df.csv','filename_temp':'status_temp.csv','duplicate':True}
                }

global sr_qty_on_sap
sr_qty_on_sap = 0
global status_qty_on_sap
status_qty_on_sap = 0

def search_xml_on_sap(session, xmltype:str, companyid:str, searchid:str, date:list)->bool:
    '''
        sap에서 메뉴 조작하기위해 사용 
        xmltype은 send/receive 2가지 입력
        날짜는 list안에 입력 ['2023.01.01', '2023.01.31'] or ['2023.01.01']

        조회결과 있으면 (True, 결과메시지), 없으면 (False, 에러메시지) 반환
    '''
    session.findById("wnd[0]/usr/radR_ACT_D").select() # Transaction Base
    session.findById("wnd[0]/usr/radR_EXW_X").select() # EDI
    session.findById("wnd[0]/usr/radP_ACK_LN").select() # Summary
    session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = companyid # "C100"

    # send면 reveiver ID에 KORCHAMXML2 입력
    if xmltype == 'send':session.findById("wnd[0]/usr/txtS_RID-LOW").text = searchid # RECEIVER ID
    else: session.findById("wnd[0]/usr/txtS_SID-LOW").text = searchid # SENDER ID

    # 날짜입력
    if type(date) == str: date_start, date_end = date, date
    elif type(date) == list:
        if len(date) == 1: date_start, date_end = date[0], date[0]
        elif len(date) == 2: date_start, date_end = date[0], date[1]
    else: raise ValueError('날짜는 []안에 1개 또는 2개 입력필요')

    session.findById("wnd[0]/usr/ctxtSO_AEDAT-LOW").text = date_start
    session.findById("wnd[0]/usr/ctxtSO_AEDAT-HIGH").text = date_end
    session.findById("wnd[0]").sendVKey(8)

    # 조회결과 없으면 종료
    if  'Message' in session.findById("wnd[0]/sbar").Text:  # == 'Message=>Data not found':
        return (False, session.findById("wnd[0]/sbar").Text)
    elif 'limit is greater' in session.findById("wnd[0]/sbar").Text:
        return (False, session.findById("wnd[0]/sbar").Text)
    elif 'Invalid date' in session.findById("wnd[0]/sbar").Text:
        return (False, session.findById("wnd[0]/sbar").Text)
    
    # 조회 결과에서 NORMAL건 클릭하여 진입
    session.findById("wnd[0]/usr/shell/shellcont[1]/shell").currentCellColumn = "NORMAL"
    session.findById("wnd[0]/usr/shell/shellcont[1]/shell").selectedRows = "0"
    session.findById("wnd[0]/usr/shell/shellcont[1]/shell").doubleClickCurrentCell()

    # 기존 조회건수를 불러와서 같으면 조회하지않고 넘어가도록 함
    if xmltype == 'send': 
        global sr_qty_on_sap
        chk_qty = sr_qty_on_sap
    else: 
        global status_qty_on_sap
        chk_qty = status_qty_on_sap

    if session.findById("wnd[0]/usr/shell/shellcont[1]/shell").RowCount == chk_qty:
        return (False, f'{companyid} {date}추가건 없음')
    else: # 전역변수에 건수 저장
        if xmltype == 'send': 
            sr_qty_on_sap = session.findById("wnd[0]/usr/shell/shellcont[1]/shell").RowCount
        else: 
            status_qty_on_sap = session.findById("wnd[0]/usr/shell/shellcont[1]/shell").RowCount

        return (True, f'{companyid} {date}조회완료')
    
def chk_each_xml_status(session, df:pd.DataFrame, id_list:list, permit_duplicate:bool)->pd.DataFrame:

    for i in range(session.findById("wnd[0]/usr/shell/shellcont[1]/shell").RowCount):
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell").selectedRows = i
        session.findById("wnd[0]/tbar[1]/btn[5]").press()

        # 조회결과 없으면 종료
        if  'Message' in session.findById("wnd[0]/sbar").Text:  # == 'Message=>XML is not found
            print(f'{companyid} {date}(조회불가)', session.findById("wnd[0]/sbar").Text )
            return df
        
        # xml파일경로 확인 및 변환
        file_path = session.findById("wnd[0]/usr/cntlGUI_CONTAINER_X/shellcont/shell").BrowserHandle.LocationURL
        for (before, after) in (('file:///', ''), ('/', '\\'), ('\\', '\\\\')):
            file_path = file_path.replace(before, after)
        file_path = parse.unquote(file_path)

        #  xml 읽고 파싱
        with open(file_path, 'r', encoding='utf-8') as f:
            xml = f.read()
            soup = BeautifulSoup(xml, 'xml')

            temp_send_sr = {}
            for id in id_list:
                temp_send_sr[id] = soup.find(id).string

            # 중복허용시 데이터추가, 불허시 추가X (Send는 중복미허용, Receive는 중복허용 / MsgReferenceID[9101996116_COB_01])
            if permit_duplicate:
                df = pd.concat([df, pd.DataFrame([temp_send_sr])])
            else:
                if temp_send_sr['MsgReferenceID'] not in list(df['MsgReferenceID']):
                    df = pd.concat([df, pd.DataFrame([temp_send_sr])])
        
        session.findById("wnd[0]").sendVKey(3)
    
    df['DocumentDate'] = pd.to_numeric(df['DocumentDate']) # 날짜int로 변경(최신데이터 검증위한 로직)
    return df

# def read_excel_to_dataframe(file_path:str):
#     wb = xw.Book(file_path, read_only=True)
#     main_df = wb.sheets[0].used_range.options(pd.DataFrame).value
#     wb.close()

#     main_df['DocumentDate'] = main_df['DocumentDate'].astype('int64')
#     if 'DocumentNumber' in main_df.index:
#         main_df['DocumentNumber'] = main_df['DocumentNumber'].astype('int64').astype('str')

#     return main_df

def read_excel_to_dataframe(file_path:str, for_streamlit=False):
    with xw.App(visible=False) as app:
        wb = app.books.open(file_path)#, read_only=True)
        #wb = xw.Book(file_path, read_only=True,)
        main_df = wb.sheets[0].used_range.options(pd.DataFrame).value
        wb.close()

    for each_label in ['DocumentDate', 'DocumentNumber', 'DocumentDate(send)', 'DocumentDate(receive)']:
        if each_label in main_df.columns:
            main_df[each_label] = main_df[each_label].astype('int64')
            if each_label in ['DocumentNumber', 'DocumentDate(send)', 'DocumentDate(receive)']:
                main_df[each_label] = main_df[each_label].astype('str')

    return main_df

def view_chk_xmls(session, option_dict, companyid:str, date:list):
    main_df = read_excel_to_dataframe(option_dict['path'] + option_dict['filename_main'])

    session.StartTransaction('ZLLEI09020')

    xml_exist = search_xml_on_sap(session, xmltype=option_dict['xmltype'], companyid=companyid, searchid='KORCHAMXML2' , date=date)
    if xml_exist[0]:
        temp_df = pd.DataFrame(columns=option_dict['column_id'])
        temp_df = chk_each_xml_status(session, df=temp_df, id_list=option_dict['column_id'], permit_duplicate=option_dict['duplicate'])
        temp_df.to_csv(option_dict['path'] + option_dict['filename_temp'], encoding='euc-kr')
        #print(temp_df)
    else:
        print('(조회불가)', xml_exist[1])
        main_df.to_csv(option_dict['path'] + option_dict['filename_main'], encoding='euc-kr')
        return main_df

    main_df = pd.concat([main_df, temp_df])
    if option_dict['xmltype'] == 'send':
        main_df.drop_duplicates(subset=['MsgReferenceID'], inplace=True)
    else:
        main_df.drop_duplicates(inplace=True)

    main_df.to_csv(option_dict['path'] + option_dict['filename_main'], encoding='euc-kr')

    return main_df

def check_and_return_co_status(main_df:pd.DataFrame, sr_no:str):
    if len(main_df[main_df['DocumentNumber'] == sr_no]) == 0:
        # 대상없음(접수했는지 알람필요)
        status_dict = {'need_alert':True, 'sr_no':sr_no, 'msg_main':'5분 뒤 확인해주시거나, 정상접수되었는지 확인해주세요', 'msg_sub':'', 'final_alert':False}
        #msg = (sr_no, '정상접수되었는지 확인 필요합니다')
    else:
        first_row_of_df = main_df[main_df['DocumentNumber']==sr_no].sort_values(by='DocumentDate(receive)', ascending = False).iloc[0]
        if first_row_of_df['MsgStatus'] == 'ISSUED':
            pass # 발급완료(알림필요)
            status_dict = {'need_alert':True, 'sr_no':sr_no, 'msg_main':'발급이 완료되었습니다', 'msg_sub':'', 'final_alert':True}
            #msg = (sr_no, '발급이 완료되었습니다', '')
        elif first_row_of_df['MsgStatus'] == 'RECEIPT':
            pass # 접수완료(알림생략)
            status_dict = {'need_alert':False, 'sr_no':sr_no, 'msg_main':None, 'msg_sub':'', 'final_alert':False}
            #msg = (sr_no, None, '')
        elif first_row_of_df['MsgStatus'] == 'RETURN':
            pass # 발급거절(알림필요)
            status_dict = {'need_alert':True, 'sr_no':sr_no, 'msg_main':'반려되었으니 사유확인하세요', 'msg_sub':first_row_of_df['MsgStatusDetail'], 'final_alert':False}
            #msg = (sr_no, '반려되었으니 사유확인하세요', first_row_of_df['MsgStatusDetail'])
    
    return status_dict
        
def alert_with_win11toast(msg):
    if msg['final_alert'] is True:
        image_src = r'C:\python_source\noti_image.jpg'
    else:
        image_src = None
    if msg['need_alert'] is True:
        win11toast.notify(title=f"{msg['sr_no']} {msg['msg_main']}", body=msg['msg_sub'],
                    #icon=r'C:\Users\k\Pictures\BlueStacks\Screenshot_2021.06.28_00.14.22.296.png',
                    image=image_src,
                    )


## 코드 시작 ##
if __name__ == '__main__' or __name__ == 'MonitoringCOO':
    # 마지막 업데이트일부터 현재까지의 내역 업데이트
    df_srlist = read_excel_to_dataframe(option_dict['srlist']['path'] + option_dict['srlist']['filename_main'])
    df_status = read_excel_to_dataframe(option_dict['status']['path'] + option_dict['srlist']['filename_main'])

    date_value = [time.strftime('%Y.%m.%d', time.strptime(str(df_srlist['DocumentDate'].max()), '%Y%m%d')), time.strftime('%Y.%m.%d', time.localtime(time.time()))]
    main_srlist_df = view_chk_xmls(session, option_dict['srlist'], companyid='C100', date=date_value)
    main_srlist_df = view_chk_xmls(session, option_dict['srlist'], companyid='C1X0', date=date_value)
    date_value = [time.strftime('%Y.%m.%d', time.strptime(str(df_status['DocumentDate'].max()), '%Y%m%d')), time.strftime('%Y.%m.%d', time.localtime(time.time()))]
    main_status_df = view_chk_xmls(session, option_dict['status'], companyid='C100', date=date_value)
    main_status_df = view_chk_xmls(session, option_dict['status'], companyid='C1X0', date=date_value)

    # 오늘자 데이터 크롤링 시작
    date_value = time.strftime('%Y.%m.%d', time.localtime(time.time()))
    while True:
        # NERP 크롤링
        main_srlist_df = view_chk_xmls(session, option_dict['srlist'], companyid='C100', date=date_value)
        main_srlist_df = view_chk_xmls(session, option_dict['srlist'], companyid='C1X0', date=date_value)

        main_status_df = view_chk_xmls(session, option_dict['status'], companyid='C100', date=date_value)
        main_status_df = view_chk_xmls(session, option_dict['status'], companyid='C1X0', date=date_value)

        # 조회용 테이블 생성
        main_total_list_df = pd.merge(main_srlist_df, main_status_df, how='inner', on='MsgReferenceID', suffixes=['(send)', '(receive)'])
        main_total_list_df.to_csv(file_path+'test_total.csv', encoding='euc-kr')

        # 알람필요한 SR 확인 후 실행 (완료건은 리스트 삭제)
        with open(file_path + "list_sr.pickle","rb") as f:
            list_sr = pickle.load(f)

        for sr_no in list_sr:
            status_dict = check_and_return_co_status(main_total_list_df, sr_no)
            if status_dict['need_alert'] is True:
                alert_with_win11toast(status_dict)
            if status_dict['final_alert'] is True:
                list_sr.remove(sr_no)

        with open(file_path + "list_sr.pickle","wb") as f:
            pickle.dump(list_sr, f)
        
        time.sleep(refresh_seconds)