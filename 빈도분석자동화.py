# 이 스크립트를 실행하게 되면 spss 분석이 끝난 빈도분석 xlsx파일을 한글에 바로 붙여넣기 할 수 있게 가공해준다.
# 가공이 완료된 파일은 같은 폴더내에 생성되고, 파일 이름앞에 "가공된"이라는 단어가 추가되서 생선된다.

import numpy as np
import pandas as pd
import copy


# 같은 폴더 내에 파일이 있어야합니다. 파일을 바꾸고싶으시면 파일경로 변수에 파일을 변경하시면 됩니다
파일경로 = "기술통계_빈도분석.xlsx"
df = pd.read_excel(파일경로, header=None)


# 데이터프레임 사이에 NaN행을 제거한 복사본 생성, index 초기화
df = df.dropna(how='all').reset_index(drop=True)
# round함수가 문자열이 포함되어있는 열에는 작동이 안돼서, 문자열은 그대로 숫자형만 소수점 반올림 적용 
df[3] = df[3].apply(lambda x: round(float(x), 1) if str(x).replace('.', '', 1).isdigit() else x)


# np.where()를 통해 value가 "전체"인 데이터의 행열을 리스트에 저장한다.
value_to_find = "전체"
전체rows, 전체cols = np.where(df.values == value_to_find)

# 마찬가지로 value가 "유효"인 데이터의 행열을 리스트에 저장한다.
# 우리가 찾고싶은건 유효가 아니라 그 두칸위의 제목이므로 row에 2씩 추가한다.
# 첫번째 유효는 데이터의 전체통계량에도 있어서 하나가 더생겼다. 첫번째 데이터는 삭제해야됨
value_to_find = "유효"
유효rows, 유효cols = np.where(df.values == value_to_find)

for i in range(len(유효rows)):
    유효rows[i] = 유효rows[i] - 2

# rows에서 첫번째 제거해주고, cols도 마찬가지로 1번째 데이터 제거 (첫번쩨로 나온 "유효"는 전체통계량을 나타내므로 필요없음)
제목rows = 유효rows[1:]
제목cols = 유효cols[1:]


# 결국 rows와 cols에는 그 데이터프레임의 제목의 위치가 담겨져있게 된다.
for r, c in zip(제목rows,제목cols):
    value = df.iloc[r,c]
    


# 분리된 데이터프레임이 저장될 새로운 리스트 생성
# 열은 0열부터 5열까지만 저장하면됨 
# 행은 성별이 있는 행부터 전체가 있는행까지 저장하면, 분리된 리스트 완성
df_list = []

for start, end in zip(제목rows,전체rows):
    # 슬라이싱 :  각 제목 행부터 전체 행까지, 0열부터 5열까지
    print(f"시작 행 : {start}, {df.iloc[start,0]} 마지막 행 : {end}, {df.iloc[end,1]}")
    sub_df = df.iloc[start:end+1, :6]
    df_list.append(sub_df)


fixed_list = []

columns_to_drop = df.columns[[3,4]]

for i, data in enumerate(df_list):
    data_copy = copy.deepcopy(df_list[i]).reset_index(drop=True)
    data_copy.iloc[2,0] = data_copy.iloc[0,0]
    #위에서부터 2개의 행 제거 
    data_copy = data_copy.iloc[2:]
    # 마지막 행 제거 
    data_copy = data_copy.iloc[:-1]
    # 빈도 수치에 백분율값을 붙이기위해, iloc을 통해 가져온 value와 기존 value를 합쳐준다.
    # 결합 결과를 2번째 열에 추가
    data_copy[3] = data_copy[3].round(1)
    data_copy[2] = data_copy[2].astype(str) + " (" + data_copy[3].astype(str) + ")"
    del data_copy[3]
    del data_copy[4]
    del data_copy[5]
    fixed_list.append(data_copy)







# 마지막으로 각 리스트에 저장된 df를 하나로 합쳐준다
merged_df = pd.concat(fixed_list, ignore_index=True)

# 열 개수와 행 개수가 결과값에 나오므로 한글에서는 그에 맞게 표 만들기를 누른다음 
# 데이터만 덮어쓰기를 하면 편하게 작성할 수 있다.
# 빈도표에 꼭 필요한 값만 작성했다.(빈도수, 백분율) 논문에 따라 상이한 내용은 추가적으로 입력할 수 밖에 없다. (현재로서는)
# (연습했던 논문에서는 나이에 추가적으로 표준편차값 입력이 필요했다)


# 이제 엑셀파일로 내보내면 된다. 후......

merged_df.to_excel(f'F_빈도_{파일경로}', index=False)
print("Excel로 저장완료")






