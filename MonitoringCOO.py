import streamlit as st
import pandas as pd
import pickle
import sqlite3
import json
from datetime import datetime
from datetime import timedelta

import MonitoringCOO_crawler


# 기준정보
#file_path_db = 'C:\\python_source\\MonitoringCOO\\'
file_path_db = r'\\23.20.135.83\파일 공유\MoniteringCOO' + '\\'
#file_path_pickle = '.\dict_for_push.json'
#file_path_pickle = 'D:\\파일 공유\\MoniteringCOO\\dict_for_push.json'
file_path_pickle = r'\\23.20.135.83\파일 공유\MoniteringCOO' + '\\dict_for_push.json'
file_path_down = 'C:\\python_source\\'

with open(file_path_db + 'MonitoringCOO_crawler.json', 'r', encoding='utf-8')as f:
        korcham_opt = json.load(f, strict=False)

def read_db_to_dataframe(path_of_db, sql_txt):
    conn_db = sqlite3.connect(path_of_db)
    db_cursor = conn_db.cursor()
    
    query = db_cursor.execute(sql_txt)
    cols = [column[0] for column in query.description]
    df = pd.DataFrame.from_records(data=query.fetchall(),columns=cols)
    conn_db.close()

    return df.sort_values(by='접수일시', ascending=False)

def write_load_pickle(wr_type, file_path, list_for_pickle=None):
    with open(file_path, wr_type) as f:
        if 'w' in wr_type:
            if list_for_pickle is not None:
                pickle.dump(list_for_pickle, f)
            else:
                raise Exception('Dump할 list를 함수에 넣어주세요')
        elif 'r' in wr_type:
            return pickle.load(f)
        else:
            raise Exception('w, wb, r, rb 중 하나로 입력하세요')

def write_load_json(wr_type, file_path, list_object=None):
    with open(file_path, wr_type, encoding='utf-8') as f:
        if 'w' in wr_type:
            if list_object is not None:
                json.dump(list_object, f, indent=2, ensure_ascii=False)
            else:
                raise Exception('Dump할 값을 함수에 넣어주세요')
        elif 'r' in wr_type:
            return json.load(f, strict=False)
        else:
            raise Exception('w, wb, r, rb 중 하나로 입력하세요')

st.set_page_config(layout="wide")

tab1, tab2, tab3 = st.tabs(["View and alarm", "Setting Login info(미사용)", 'Login manually(미사용)'])

with tab1:
    col1,col2 = st.columns([2,8])
    # 공간을 2:3 으로 분할하여 col1과 col2라는 이름을 가진 컬럼을 생성합니다.  

    with col1 : # 화면 좌측
        st.title('Input SR No for Push alarm')
        username = st.text_input(label="본인을 식별할 id를 입력해주세요", value='', max_chars=20, help='20자리 텍스트만 입력가능', autocomplete='on')
        sr_alarm = st.text_input(label="Push를 받을 SR번호를 입력하세요", value='', max_chars=10, help='10자리 숫자만 입력가능',)#, autocomplete='on')
        
        col1_1, col1_2, col1_3 = st.columns([3,3,4])
        with col1_1:
            if st.button('📌추가'): # sr번호를 입력하면 알림받을 리스트로 저장
                list_sr = write_load_json("r", file_path_pickle)
                list_sr[sr_alarm] = username
                write_load_json("w", file_path_pickle, list_sr)
        with col1_2:
            if st.button('🛒삭제'):
                list_sr = write_load_json("r", file_path_pickle)
                try: # 있으면 삭제하고 없으면 넘김
                    list_sr.pop(sr_alarm)
                except: pass
                write_load_json("w", file_path_pickle, list_sr)  
        with col1_3:
            if st.button('⏳새로고침',):
                pass

        list_sr = write_load_json("r", file_path_pickle)
        for sr_each in list_sr:
            st.markdown('* '+sr_each)
                 
    with col2 : # 화면 우측
        st.title('View status of COO with SR No')
        user_text_input = st.text_input(label="검색할 SR번호를 입력하세요", value='', max_chars=10, help='10자리 숫자만 입력 가능')#, autocomplete='on')
        search_df = read_db_to_dataframe(file_path_db + 'Korcham_status.db', '''SELECT * FROM 상공_SEC_NEGO1
                          UNION ALL
                          SELECT * FROM 상공_SEC_NEGO2
                          UNION ALL
                          SELECT * FROM 상공_SMC_NEGO1''')
        if user_text_input or user_text_input == '':
            search_df = search_df[search_df['대표Invoice'].str.contains(user_text_input)]
        st.dataframe(search_df, width=1500, hide_index=True)

        col2_col1, col2_col2, col2_col3, col2_col4 = st.columns([5,0.5,0.5,2])
        with col2_col1:
            with st.container():
                st.text(f"마지막 업데이트 : \n SEC_NEGO1 : {korcham_opt['last_update']['SEC_NEGO1']}\n SEC_NEGO2 : {korcham_opt['last_update']['SEC_NEGO2']}\n SMC_NEGO1 : {korcham_opt['last_update']['SMC_NEGO1']}")
        with col2_col2:
            year_to_update_report = st.text_input(label="연(YYYY) ", value=str((datetime.today()-timedelta(30)).year), max_chars=4, help='4자리 숫자만 입력 가능')
        with col2_col3:    
            month_to_update_report = st.text_input(label="월(MM) ", value=str((datetime.today()-timedelta(30)).month), max_chars=2, help='2자리 숫자만 입력 가능')
        with col2_col4:
            report_criteria = f'{year_to_update_report}-{int(month_to_update_report):02d}'
            report_save_to = st.text_input(label="저장경로", value=f'{file_path_down}\COO_실적_{report_criteria}.xlsx', help='4자리 숫자만 입력 가능')
            if st.button('월 발급완료건 다운로드(NEGO1,2)',):
                # read_db_to_dataframe(file_path_db + 'Korcham_status.db', f'''
                #                     SELECT *
                #                     FROM (
                #                     SELECT * FROM 상공_SEC_NEGO1
                #                     WHERE 
                #                     접수일시 LIKE '{report_criteria}%' 
                #                     AND 처리상태 LIKE '%발급완료%'
                #                     AND 증명서종류 = '일반(비특혜/Non-preferential) 원산지증명서'
                #                     UNION ALL
                #                     SELECT * FROM 상공_SEC_NEGO2
                #                     WHERE 
                #                     접수일시 LIKE '{report_criteria}%' 
                #                     AND 처리상태 LIKE '%발급완료%'
                #                     AND 증명서종류 = '일반(비특혜/Non-preferential) 원산지증명서'
                #                     ) AS CombinedResults;''').drop_duplicates('대표Invoice').to_excel(report_save_to)
                read_db_to_dataframe(file_path_db + 'Korcham_status.db', f'''
                                    SELECT *
                                    FROM (
                                        SELECT * FROM 상공_SEC_NEGO1
                                        WHERE 
                                        접수일시 LIKE '{report_criteria}%' 
                                        AND (처리상태 LIKE '%발급완료 (Accept)\n[ 신규 ]%'
                                            OR 처리상태 LIKE '%발급완료 (Accept)\n[ 정정 ]%')
                                        AND 증명서종류 = '일반(비특혜/Non-preferential) 원산지증명서'
                                        UNION ALL
                                        SELECT * FROM 상공_SEC_NEGO2
                                        WHERE 
                                        접수일시 LIKE '{report_criteria}%' 
                                        AND (처리상태 LIKE '%발급완료 (Accept)\n[ 신규 ]%'
                                            OR 처리상태 LIKE '%발급완료 (Accept)\n[ 정정 ]%')
                                        AND 증명서종류 = '일반(비특혜/Non-preferential) 원산지증명서'
                                    ) AS CombinedResults;''').drop_duplicates('대표Invoice').to_excel(report_save_to)



with tab2:
    st.title('상공회의소 로그인에 필요한 정보를 입력해주세요')
    with open(file_path_db + 'MonitoringCOO_crawler.json', 'r', encoding='utf-8')as f:
        korcham_opt = json.load(f, strict=False)

    for each_company in [i for i in list(korcham_opt.keys()) if i not in['loc_and_expiry','last_update']]:
        chg_taxid = st.text_input(label=f'{each_company} 사업자등록번호', value=korcham_opt[each_company][0])
        if chg_taxid:
            korcham_opt[each_company][0] = chg_taxid
        chg_id = st.text_input(label=f'{each_company} ID', value=korcham_opt[each_company][1])
        if chg_id:
            korcham_opt[each_company][1] = chg_id
        chg_pw = st.text_input(label=f'{each_company} PW', value=korcham_opt[each_company][2])
        if chg_pw:
            korcham_opt[each_company][2] = chg_pw
        if st.button(f'{each_company}반영'):

            with open(file_path_db + 'MonitoringCOO_crawler.json', 'w', encoding='UTF-8')as f:
                json.dump(korcham_opt, f, indent=2, ensure_ascii=False)

with tab3:
    
    # 로그인 기능버튼
    for each_company in [i for i in list(korcham_opt.keys()) if i not in['loc_and_expiry','last_update']]:
        if st.button(each_company):
            selenium_driver =  MonitoringCOO_crawler.open_korcham(korcham_opt, each_company)

# streamlit run MonitoringCOO.py
# streamlit run .\MonitoringCOO\MonitoringCOO.py