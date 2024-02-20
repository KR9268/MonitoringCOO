import pickle
import win11toast
import sqlite3
import pandas as pd
import time
from MonitoringCOO import write_load_json, read_db_to_dataframe


# file_path_db = 'C:\\python_source\\MonitoringCOO\\'
file_path_db =  r'\\23.20.135.83\파일 공유\MoniteringCOO' + '\\'
#file_path_db = 'D:\\파일 공유\\MoniteringCOO\\'

file_path_pickle = file_path_db + 'dict_for_push.json'
file_path_image = file_path_db + 'noti_image.jpg'

#file_path = 'D:\\파일 공유\\MoniteringCOO\\'


def check_and_return_co_status(main_df:pd.DataFrame, sr_no:str):
    if len(main_df[main_df['대표Invoice'] == sr_no]) == 0:
        # 대상없음(접수했는지 알람필요)
        status_dict = {'need_alert':True, 'sr_no':sr_no, 'msg_main':'5분 뒤 확인해주시거나, 정상접수되었는지 확인해주세요', 'msg_sub':'', 'final_alert':False, 'use_img':False}
    else:
        first_row_of_df = main_df[main_df['대표Invoice']==sr_no].sort_values(by='접수일시', ascending = False).iloc[0]
        if (first_row_of_df['처리상태'] == '발급완료 (Accept)\n[ 정정 ]') or (first_row_of_df['처리상태'] == '발급완료 (Accept)\n[ 신규 ]') or (first_row_of_df['처리상태'] == '발급완료 (Accept)'):
            status_dict = {'need_alert':True, 'sr_no':sr_no, 'msg_main':'발급이 완료되었습니다', 'msg_sub':'', 'final_alert':True, 'use_img':True}
        elif first_row_of_df['처리상태'] == '접수완료 (Application)' or (first_row_of_df['처리상태'] == '발급완료 (Accept)\n[ 진정등본/정정에 의한 취소 ]'):
            status_dict = {'need_alert':False, 'sr_no':sr_no, 'msg_main':None, 'msg_sub':'', 'final_alert':False, 'use_img':False}
        elif first_row_of_df['처리상태'] == '오류통보':
            status_dict = {'need_alert':True, 'sr_no':sr_no, 'msg_main':'반려되었으니 사유확인하세요', 'msg_sub':first_row_of_df['Remark'], 'final_alert':True, 'use_img':False}
        else:
            status_dict = {'need_alert':False, 'sr_no':sr_no, 'msg_main':'특이사항이 발생했습니다 확인해주세요', 'msg_sub':'발급완료 등 일반적인 처리상태가 아닙니다', 'final_alert':False, 'use_img':False}
            
    return status_dict
        
def alert_with_win11toast(msg:dict, username:str):
    if msg['use_img'] is True:
        image_src = file_path_db + 'noti_image.jpg'
    else:
        image_src = None
    if msg['need_alert'] is True:
        win11toast.notify(title=f"[{username}]{msg['sr_no']} {msg['msg_main']}", body=msg['msg_sub'],
                    #icon=r'C:\Users\k\Pictures\BlueStacks\Screenshot_2021.06.28_00.14.22.296.png',
                    #image=image_src,
                    )
        
def main_push():
    wait_to_pop = []
    while True:
        df_coo = read_db_to_dataframe(file_path_db + 'Korcham_status.db')

        list_sr = write_load_json("r", file_path_pickle)
        # with open(file_path_db + "list_sr.pickle","rb") as f:
        #     list_sr = pickle.load(f)

        for sr_no in list_sr:
            status_dict = check_and_return_co_status(df_coo, sr_no)
            if status_dict['need_alert'] is True:
                alert_with_win11toast(status_dict, list_sr[sr_no])
            if status_dict['final_alert'] is True:
                wait_to_pop.append(sr_no)
    
        for sr_to_delete in wait_to_pop:
            try:
                list_sr.pop(sr_to_delete)
            except: pass

        write_load_json("w", file_path_pickle, list_sr)  
        # with open(file_path_db + "list_sr.pickle","wb") as f:
        #     pickle.dump(list_sr, f)

        time.sleep(30)
        
if __name__ == '__main__':
    # 알람필요한 SR 확인 후 실행 (완료건은 리스트 삭제)
    main_push()