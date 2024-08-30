# -*- coding: utf-8 -*-
import win32com.client
from pandas import Series, DataFrame
import datetime
import pandas as pd

instFutOptChart = win32com.client.Dispatch("CpSysDib.FutOptChart")  # 선물/옵션 데이터 연결
final_DF = DataFrame() 
start = '20240713'

while True:
    code = '10100'
    instFutOptChart.SetInputValue(0, code)   # 받고자 하는 종목 코드 입력
    instFutOptChart.SetInputValue(1, ord('1'))  # 1: 기간별 데이터, 2: 개수 데이터
    instFutOptChart.SetInputValue(2, int(start))  # 요청 종료일 / 기재한 날짜포함 출력됨
    instFutOptChart.SetInputValue(3, 20100101)  # 요청 시작일 / 기재한 날짜포함 출력됨
    instFutOptChart.SetInputValue(5, (0, 1, 2, 3, 4, 5, 8)) # 필드 또는 필드 배열, 0: 날짜, 1: 시간, 2~5: 시고저종, 8: 거래량
    instFutOptChart.SetInputValue(6, ord('m'))  # 차트 구분, ('m': 분), D(일), W(주), M(월), S(초), T(틱)
    instFutOptChart.SetInputValue(7, 1) # 주기 (1분, 1일, 1주 등등)
    instFutOptChart.SetInputValue(9, ord('0'))  # 수정주가, 0: 미사용, 1: 사용

    instFutOptChart.BlockRequest()

    numData = instFutOptChart.GetHeaderValue(3) # 3: 수신개수
    numField = instFutOptChart.GetHeaderValue(1)    # 1: 필드개수(위에서 SetInputValue의 5번 항목의 갯수)
    nameField = instFutOptChart.GetHeaderValue(2)   #  2: 필드명

    # 데이터 받아오기
    df = DataFrame()
    for i in range(numField):
        a = []
        for j in range(numData):
            a.append(instFutOptChart.GetDataValue(i, j))
        df[instFutOptChart.GetHeaderValue(2)[i]] = a

    
    # 날짜, 시간 type 변경
    Date = []
    for i in range(len(df[df.keys()[0]])):
        Date.append(str(df[df.keys()[0]][i])[:4] + str(df[df.keys()[0]][i])[4:6] + str(df[df.keys()[0]][i])[6:])

    Time = []
    for i in range(len(df[df.keys()[1]])):
        if len(str(df[df.keys()[1]][i])) == 4:
            Time.append(str(df[df.keys()[1]][i])[:-2] + str(df[df.keys()[1]][i])[-2:] + '00')
        else:
            Time.append('0'+str(df[df.keys()[1]][i])[:-2] + str(df[df.keys()[1]][i])[-2:] + '00')

    
    df[df.keys()[0]] = Date
    df[df.keys()[1]] = Time
    # print(df.head(1))
    # 소수점 둘째자리까지
    df = round(df, 2)
    if len(df) == 0:
        final_DF.to_csv(r'C:\Users\parkgong\Documents\1min_future_data\1m_future.csv',index=False)
        break
    
    # Dates of start reset
    dates = df[df.keys()[0]]
    dates = list(set(dates))
    dates.sort()
    start = dates[0]
    start = datetime.datetime.strptime(start, '%Y%m%d').date()
    start = start - datetime.timedelta(days=1)
    start = start.strftime('%Y%m%d')

    # Date Time 을 하나로 병합
    Date_Time = []
    for i in range(len(df[df.keys()[0]])):
        Date_Time.append(str(df[df.keys()[0]][i])+str(df[df.keys()[1]][i]))
    
     
       
    del df['날짜']
    del df['시간']
    
    df.insert(0,'date_time',Date_Time) 
   
                         
    

    # 각 열의 제목을 일괄 변경
    index = ['date_time', 'open', 'high', 'low', 'close', 'volume']
    for i in range(len(df.columns)):
        df.rename(columns = {df.keys()[i]: index[i]}, inplace = True)
    
    df = df[::-1]

    print(df.head(1))
   
    
    
    
    final_DF = pd.concat([df,final_DF]).reset_index(drop=True)

    
    # csv.to_csv(r'C:\Users\parkgong\Documents\1min_future_data\1m_future.csv' , sep=",")


    # for date in dates:
    #     try:

    #         csv = df.loc[date]
    #         csv = csv.reset_index(drop=True)
    #         csv.index = csv['time']
    #         del csv['time']
    #         csv.to_csv(r'C:\Users\parkgong\Documents\1min_future_data\1m_future.csv' , sep=",")

    #     except KeyError:
    #         print(date)
