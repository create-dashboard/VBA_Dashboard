import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt

# def  main():
#     #xlwings 통해 Workbook 호출
#     wb = xw.Book.caller()

#     #Sheet 설정
#     sheet = wb.sheets[0]

#     #텍스트 입력
#     sheet["A1"].value = "xlwings 테스트 코드 작성"
#     sheet["A2"].value = "파이썬 업무 자동화"

# if  __name__== "__main__":
#     path = r"C:\sjbang\STUDY\VBA_Dashboard\02_xlwings"
#     #매크로 파일 설정
#     xw.Book(path+"/"+"p1.xlsm").set_mock_caller()
#     #main 함수 호출
#     main()


def main():
    wb = xw.Book('py1.xlsm')
    sheet = wb.sheets['Sheet1']  # 원본 데이터가 있는 시트
    
    # 1. 데이터 읽기
    df = sheet.range('A1').options(pd.DataFrame, expand='table', index=False).value
    
    # 2. Pandas로 피벗 테이블 생성
    pivot_df = df.pivot_table(
        index='Region', 
        # columns='Sales', 
        values='Sales', 
        aggfunc='sum',
        fill_value=0
    ).reset_index()
    
    # plt.plot(pivot_df['Region'], pivot_df['Sales'])
    # plt.title('Sales by Region')
    # plt.xlabel('Region')
    # plt.ylabel('Sales')
    # plt.show()

    # 3. 결과 전송 (새 시트 생성 또는 특정 위치)
    if 'PivotResult' not in [s.name for s in wb.sheets]:
        wb.sheets.add('PivotResult')
    
    res_sheet = wb.sheets['PivotResult']
    res_sheet.clear()
    res_sheet.range('A1').value = pivot_df

    chart = res_sheet.charts.add(left=300, top=10) # 차트 위치 설정
    
    # 4. 차트에 데이터 연결
    # xlwings 차트의 데이터 소스는 DataFrame이 아닌 Range 객체여야 합니다.
    data_range = res_sheet.range('A1').expand()
    chart.set_source_data(data_range)
    
    # 5. 차트 종류 및 옵션 설정
    chart.chart_type = 'line'
    chart.name = "Sales by Region"
    
    # 제목 넣기 (Excel 객체 모델 접근)
    chart.api[1].HasTitle = True
    chart.api[1].ChartTitle.Text = "Sales by Region"

    # 6. 엑셀의 피벗테이블
    # 원본 데이터 시트
    data_sheet = wb.sheets['Sheet1']
    # 피벗 테이블이 들어갈 시트 (없으면 생성)
    if 'PivotSheet' not in [s.name for s in wb.sheets]:
        wb.sheets.add('PivotSheet')
    pivot_sheet = wb.sheets['PivotSheet']
    pivot_sheet.clear()

    # 1. 원본 데이터 범위 잡기 (A1부터 데이터가 있는 곳까지)
    data_range = data_sheet.range('A1').expand()
    
    # 2. 피벗 캐시 생성 및 피벗 테이블 삽입
    # 엑셀의 내부 API를 직접 호출합니다.
    pivot_cache = wb.api.PivotCaches().Create(SourceType=1, SourceData=data_range.api)
    
    # 피벗 테이블 생성 (PivotSheet의 A3 셀에 생성)
    pivot_table = pivot_cache.CreatePivotTable(
        TableDestination=pivot_sheet.range('A3').api,
        TableName="MyPivotTable"
    )

    # 3. 피벗 필드 설정 (행, 열, 값)
    # 엑셀 필드 이름과 동일해야 합니다.
    # 예: '도시'를 행으로, '상품'을 열로, '판매량'을 합계 값으로 설정
    pivot_table.PivotFields("Region").Orientation = 1  # 1 = xlRowField (행)
    # pivot_table.PivotFields("상품").Orientation = 2  # 2 = xlColumnField (열)
    
    data_field = pivot_table.PivotFields("Sales")
    data_field.Orientation = 4  # 4 = xlDataField (값)
    data_field.Function = -4157  # -4157 = xlSum (합계)
    data_field.NumberFormat = "#,##0"

    print("피벗 테이블 생성 완료!")