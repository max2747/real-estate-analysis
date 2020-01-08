import pandas as pd
import xlwings as xw

def KBpriceindex_preprocessing(path, sheet_name):
    # path : KB 데이터 엑셀 파일의 경로
    # sheet_name : ‘매매종합’, ‘매매APT’, ‘매매연립’, ‘매매단독’, ‘전세종합’, ‘전세APT’, ‘전세연립’, ‘전세단독’
    
    wb = xw.Book(path)
    sheet = wb.sheets[sheet_name]
    row_num = sheet.range((1,1)).end('down').end('down').end('down').row
    data_range = 'A2:GE' + str(row_num)
    raw_data = sheet[data_range].options(pd.DataFrame, index=False, header=True).value
    
    big_names = '서울 대구 부산 대전 광주 인천 울산 세종 경기 강원 충북 충남 전북 전남 경북 경남 제주도 6개광역시 5개광역시 수도권 기타지방 구분 전국'
    big_name_list = big_names.split(' ')
    
    big_col = list(raw_data.columns)

    small_col = list(raw_data.iloc[0])
    for idx, gu_data in enumerate(small_col):
        if gu_data == None:
            small_col[idx] = big_col[idx]
        check = idx
        while True:
            if big_col[check] in big_name_list:
                big_col[idx] = big_col[check]
                break
            else:
                check = check - 1
    
    # '광주' -> '경기'
    big_col[129] = '경기'
    big_col[130] = '경기'
    # '제주/\n서귀포' -> '서귀포'
    small_col[185] = '서귀포'
    
    raw_data.columns = [big_col, small_col]
    new_col_data = raw_data.drop([0, 1])
    
    index_list = list(new_col_data['구분']['구분'])
    new_index = []

    for idx, raw_index in enumerate(index_list):
        tmp = str(raw_index).split('.')
        if int(tmp[0]) > 12:
            if len(tmp[0]) == 2:
                # 19XX년 
                new_index.append('19' + tmp[0] + '.' + tmp[1])
            else:
                # 2000년 이후
                new_index.append(tmp[0] + '.' + tmp[1])
        else:
            # month만 있는 경우
            # 연도 값을 앞 index에서 가져오고, 뒤에 month를 붙임.
            new_index.append(new_index[idx-1].split('.')[0] + '.' + tmp[0])
    
    new_col_data.set_index(pd.to_datetime(new_index), inplace=True)
    cleaned_data = new_col_data.drop(('구분', '구분'), axis=1)
    
    return cleaned_data