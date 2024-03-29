import streamlit as st
from streamlit_option_menu import option_menu
import streamlit.components.v1 as html
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, time
from dateutil.relativedelta import relativedelta
from openpyxl import load_workbook
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment
from dateutil.parser import parse #텍스트날짜를 날짜 형식으로 변경
import warnings
import re
import os

warnings.filterwarnings(action='ignore')  #경로 무시, 다시 적용시 default

st.title('스마트링크 청구내역서 관리')
st.header('월간청구내역서 생성')
st.write(os.path.dirname(__file__))



# 차량번호 cleansing 함수
def carnoclean(car):
    carno = re.sub('\([^)]*\)', '', car)    
    if '(삭제)_고객사변경' in str(carno):
        return carno.replace('(삭제)_고객사변경', '')
    elif '_고객사변경' in str(carno):
        return carno.replace('_고객사변경', '')
    elif '__고객사변경' in str(carno):
        return carno.replace('__고객사변경', '')
    elif '(삭제)' in str(carno):
        return carno.replace('(삭제)', '')
    elif '(반납)' in str(carno):
        return carno.replace('(반납)', '')
    elif '(교체)' in str(carno):
        return carno.replace('(교체)', '')
    elif '(진행불가)' in str(carno):
        return carno.replace('(진행불가)', '')
    elif '(임시)' in str(carno):
        return carno.replace('(임시)', '')
    elif '(회수)' in str(carno):
        return carno.replace('(회수)', '')
    elif '(기존)' in str(carno):
        return carno.replace('(기존)', '')
    elif '(ㅅㅈ)' in str(carno):
        return carno.replace('(ㅅㅈ)', '')
    elif '(테스트)' in str(carno):
        return carno.replace('(테스트)', '')
    elif '__' in str(carno):
        return carno.replace('_', '')
    elif '_' in str(carno):
        return carno.replace('_', '')
    
    return carno

# 서비스명 조정 함수
def service_name(service):  
    name = re.sub('\([^)]*\)', '', service[0])
    if name != '-':       
        return name.replace('-', '')
    else:
        if service[1] == 'keybox':
            return '카셰어링프리미엄'
        return '차량운행관리'

# 이용료 누락 차량에 이용료 추가
def no_fare(service, price1, price2):
    print(service[0], service[1])
    if service[1] == 0.0:
        print(service[1])
        if service[0] == '차량운행관리':
            print(price1)
            return price1
        else:
            return price2
    return service[1]

# 청구기준일 이후 신규장착 차량 서비스시작일 변경
def service_start(service_date, s_date):
    if (service_date > s_date):
        return service_date
    else:
        return s_date

@st.cache_data
def sims():
    sims_raw=pd.read_excel("billing_car.xlsx")
    sims_raw['서비스1'] = sims_raw['서비스1'].fillna('-')
    sims_raw['equipnam2'] = sims_raw['equipnam2'].fillna('-')
    with st.expander("청구대상 차량리스트"):
        st.dataframe(sims_raw)

    sims_raw['차량번호(clean)'] = sims_raw['차량번호'].apply(carnoclean)
    sims_raw['서비스명'] = sims_raw[['서비스1', 'equipnam2']].apply(service_name, axis=1)
    sims_raw['장착일'] = sims_raw['장착일'].apply(pd.to_datetime)
    return sims_raw


customer_file = st.file_uploader('#### 청구고객사 엑셀파일을 업로드하세요 ####')
if customer_file is not None:
    customer_raw=pd.read_excel(customer_file)
    with st.expander("청구대상 고객사"):
        st.dataframe(customer_raw)
    # 청구양식 사용 고객사 리스트 추출하기
    customer_raw = customer_raw.drop(['순번'],axis=1)
    customer_raw.loc[customer_raw['주유'].str.contains('Y', na=False), '카드'] = 'Y'
    customer_raw.loc[customer_raw['하이패스'].str.contains('Y', na=False), '카드'] = 'Y'
    customer_raw['카드'] = customer_raw['카드'].fillna('N')
    customer_list = customer_raw[customer_raw['계좌번호'].notnull() ]
    customer_list['CMS고객사명'] = customer_list['CMS고객사명'].str.replace('㈜','(주)')
    customer_list['CMS고객사명'] = customer_list['CMS고객사명'].str.replace(' ','')
    customer_count = customer_list['법인명'].count()

    st.write('직접청구 및 청구계좌번호 존재하는 고객사 대상')
    with st.expander("청구이력서 생성 고객사"):
        st.dataframe(customer_list)

    customer=pd.read_excel("customer_name.xlsx")
    st.write('고객사 수 ', customer_count, ' 개사')

    customer_name = st.selectbox(
        '고객사를 선택하세요',
        (customer))

    customer_bill = customer_list[customer_list['CMS고객사명'] == customer_name]
    customer_bill_name = customer_bill.iloc[0,2]  #청구서용 법인명 불러오기
    customer_code = customer_bill.iloc[0,9]    #사업자번호 불러오기
    customer_account = customer_bill.iloc[0,17] #계좌번호 불러오기
    card_use = customer_bill.iloc[0,28]   #하나카드 사용여부
    bill_month = customer_bill.iloc[0,26]
            # 고객사 서비스별 단가 매칭 
    company_data = list(dataframe_to_rows(customer_bill, index=False, header=False))
    price1 = company_data[0][19]  # 차량운행관리
    price2 = company_data[0][21]  # 카셰어링프리미엄
    print('이용료',company_data)

    st.write('청구고객사명:',customer_name, ', 사업자등록번호:',customer_code, ', 계좌번호:',customer_account, ', 청구기준:', bill_month, ', 하나카드사용:', card_use)
    # 청구기준일, 해당월 종료일, 해당월 날짜기간
    start_date = st.date_input('##### 청구기준일자 입력 #####', value=None)
    bill_date = st.date_input('##### 청구서 작성일 입력 #####', datetime.now())
    st.write('청구기준일:', start_date)
    if start_date is not None:
        end_date = (start_date + relativedelta(months=1)- timedelta(days=1)).strftime('%Y-%m-%d')
        # full_day
        full_day = (datetime.strptime(end_date, '%Y-%m-%d').date() - start_date).days

        #청구월, 연도, 연월 계산
        month = start_date.month
        year = start_date.year
        year_month = start_date.strftime("%Y%m")

    if st.button('청구대상 차량 불러오기'):
        sims_raw = sims()
        sims_customer = sims_raw.loc[(sims_raw['고객명'] == customer_name) & (sims_raw['장착일'] <= end_date)]
        columns = ['차량번호(clean)', '차종', '서비스명', '단가1', '장착일']
        customer_car = sims_customer[columns]
    # 청구 고객사 차량 리스트 추출
    # if st.button('고객사 차량 불러오기'):
    # sims_customer = sims_raw.loc[sims_raw['고객명'] == customer_name]
        st.write(customer_name, ' 차량리스트')
        st.dataframe(sims_customer)
        st.write('청구내역서 항목')
        st.dataframe(customer_car)


        #CMS 데이터 기준 서비스 미사용 차량 조회 (월별로 사전에 생성)
        cms_raw = pd.read_excel(f"cms_off_list_({year_month}).xlsx")
        car_off = cms_raw.loc[cms_raw['고객사'] == customer_name]
        car_off['종료일'] = car_off['종료일자'].dt.strftime("%Y-%m-%d")
        #청구 고객사 정보와 차량 매칭
        car_off_merge = pd.merge(car_off, customer_bill, left_on='고객사', right_on='CMS고객사명', how='left')
        columns = ['차량번호(clean)','모델','서비스명1', '단가1', '종료일']
        car_off_list = car_off_merge[columns]
        #종료 차량 조회
        st.write('##### 종료차량 조회 #####')
        st.dataframe(car_off_list)

        # 청구 데이터 생성
        customer_car['단말기상태'] = '이용'
        customer_car['계약기간'] = '-'


        s_date = pd.Timestamp(start_date)
        start_date_str = start_date.strftime('%Y-%m-%d')
        customer_car['이용시작'] = customer_car['장착일'].apply(service_start, s_date =s_date)
        customer_car['이용시작'] = customer_car['이용시작'].dt.strftime("%Y-%m-%d")
        # 말일 날짜 계산 (1개월 더해서 1일 빼기)
        customer_car['이용종료'] = end_date
        customer_car.reset_index(drop=True, inplace=True)
        # 이용료 누락 차량 단가 채우기
        customer_car['단가1'] = customer_car[['서비스명', '단가1']].apply(no_fare, args=(price1, price2), axis=1)
        # st.dataframe(customer_car)
        # 청구대상 차량대수
        number_of_list = customer_car['차량번호(clean)'].count()
        # 종료차량 청구대상에 추가하기
        append_data = list(dataframe_to_rows(car_off_list, index=False, header=False))
        for r_idx, row in enumerate(append_data):
            customer_car.loc[number_of_list + r_idx] = [row[0], row[1], row[2], row[3], row[4], '반납', '-', start_date_str, row[4]]

        #청구대상 기산 산정
        customer_car['start']=pd.to_datetime(customer_car['이용시작'])
        customer_car['end']=pd.to_datetime(customer_car['이용종료'])
        customer_car['gap']= customer_car['end'] - customer_car['start']
        customer_car['gap'] = customer_car['gap'].dt.days

        # 단가 일할 계산 (신규장착 - 장착일이 이용시작 이후 발생)
        customer_car['공급가액'] = round((customer_car['단가1'] / full_day) * customer_car['gap'],-1)
        #청구양식 만들기
        columns = ['차량번호(clean)', '차종','서비스명','단말기상태','계약기간','이용시작', '이용종료', '공급가액']
        car_list = customer_car[columns]
        car_list.columns = ['차량번호', '차종', '서비스구분', '단말기상태','계약기간','이용시작', '이용종료', '공급가액']
        car_sort = car_list.sort_values(by=['공급가액'])
        # 청구대상 대수 확정
        row = car_list['차량번호'].count()

        st.write(customer_name," 청구리스트")
        st.dataframe(car_sort) 
        # st.data_editor(car_sort, num_rows="dynamic")

        #청구내역서 엑셀 저장하기
        # if st.button('청구내역서 만들기'):
    #   st.write('청구내역서 만들기')
        path = os.path.dirname(__file__)
        wb = (load_workbook(f'{path}\기본청구양식.xlsx') if card_use != 'Y' else load_workbook(f'{path}\카드청구양식.xlsx'))
        # 청구서 표지 만들기
        #   st.write('청구표지 만들기')
        ws1 = wb['청구서']
        print(ws1['E15'].value)
        print(ws1['G15'].value)
        print(ws1['I23'].value)
        ws1['B4'].value = customer_bill_name  #고객사명
        ws1['B6'].value = f'{month}월 이용대금 청구서'
        ws1['O1'].value = bill_date  #청구서작성일자
        if card_use == 'N':
            ws1['I23'].value = customer_account   #계좌번호
        else:
            ws1['I31'].value = customer_account   #계좌번호
        # ws1['I23'].value = customer_account   #계좌번호
        #   st.write(customer_account)

        # 이용료 내역 만들기
        ws2 = wb['이용료']
        ws2['B2'].value = f'{year}년 {month}월' 
        ws2['B4'].value = customer_code      #사업자등록번호
        ws2['C4'].value = customer_bill_name  #고객사명
        # 카드상세 내역 만들기
        if card_use == 'Y':
            ws3 = wb['카드상세내역']
            if month != 1:
                ws3['B2'].value = f'{year}년 {month-1}월' 
            else:
                ws3['B2'].value = f'{year-1}년 {month+11}월' 

        # 스타일 지정
        border = openpyxl.styles.Border(left=openpyxl.styles.Side(style='thin'), 
                                        right=openpyxl.styles.Side(style='thin'), 
                                        top=openpyxl.styles.Side(style='thin'), 
                                        bottom=openpyxl.styles.Side(style='thin'))
        font = openpyxl.styles.Font(name='맑은고딕', size=9)

        def write_dataframe_to_excel(df, row, start_row, start_col):
            rows = list(dataframe_to_rows(df, index=False, header=False))
            for r_idx, row in enumerate(rows, start_row):
                for c_idx, value in enumerate(row, start_col):
                    cell = wb['이용료'].cell(row = r_idx, column = c_idx, value = value)
                    cell.border = border
                    cell.font = font
                    cell.alignment = Alignment(horizontal='center', vertical='center')
        
        write_dataframe_to_excel(car_sort, row-1, 6, 2)
        for i in range(6, 6+row):
            ws2['I'+str(i)].alignment = Alignment(horizontal='right', vertical='center')
            ws2['J'+str(i)] = "=round(I"+str(i)+"*0.1,0)"
            ws2['J'+str(i)].number_format = '#,##0'
            ws2['J'+str(i)].border = border
            ws2['J'+str(i)].font = font
            ws2['K'+str(i)] = "=I"+str(i)+"+J"+str(i)
            ws2['K'+str(i)].number_format = '#,##0'
            ws2['K'+str(i)].border = border
            ws2['K'+str(i)].font = font
            ws2['L'+str(i)].border = border
            ws2['M'+str(i)].border = border
            ws2['I'+str(i)].number_format = '#,##0'

        wb.save(f'{path}\{customer_name}_{month}월_스마트링크내역서.xlsx')
        st.write('청구내역서 생성완료')

else:
    st.warning('엑셀파일을 업로드 하세요')