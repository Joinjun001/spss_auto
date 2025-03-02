

# 이 스크립트를 실행하고나면 상관분석을 논문의 한글 파일에 쉽게 붙여넣기 할 수 있게 가공해준다.
# 가공이 완료된 파일은 같은 폴더내에 생성되게 되고, 파일이름앞에 "가공된"이 추가된다.


import numpy as np
import pandas as pd
import warnings
import re

# 이 변수를 True로 만들면 유의확률을 생략하고 **로 표시하고, False로 만들면 유의확률로 표기하게 된다.
별표on = False


# 파일을 불러온다. 파일은 같은 폴더내에 있어야한다.
파일경로 = "상관분석.xlsx"
df = pd.read_excel(파일경로, header=None)



# 데이터프레임 사이에 NaN행을 제거한 복사본 생성, index 초기화
df = df.dropna(how='all')

# 숫자형 데이터를 소수점 3자리까지 반올림
df = df.applymap(lambda x: round(x,3) if isinstance(x, float) else x)


value_to_find = "상관관계"
# np.where를 통해서 기술통계량의 행과 열을 찾는다. 참고로 where로 찾은 데이터는 리스트형태로 저장된다.
상관관계row, 상관관계col = np.where(df.values == value_to_find)

# "상관관계"라는 단어가 분석할때 2개씩 나오는듯 하다, 각 상관분석마다 두번째 상관관계가 표에 있는 상관관계인거 같다.
# 즉 여기서는 326열, 371열에 있는 걸 기준으로 삼으면 되겠다.
# 하드코딩하지 않으려면, 홀수 인덱스만 꺼내서 쓰면 될거같다.
상관관계row = [상관관계row[i] for i in range(len(상관관계row)) if i % 2 == 1]
상관관계row = np.array(상관관계row)

# col은 현재는 쓰진 않지만, 혹시 모르니 같이 짝을 맞춰준다.
상관관계col = [상관관계col[i] for i in range(len(상관관계col)) if i % 2 == 1]
상관관계col = np.array(상관관계col)






# 표 끝에 있는 문장으로 표의 마지막 부분을 찾는다
# 만약 ** 표시를 해제한 파일이면, 여기서 못찾고 오류가 발생하게됨
value_to_find = "**. 상관관계가 0.01 수준에서 유의합니다(양측)."

끝row, 끝col = np.where(df.values == value_to_find)



# 표의 시작과 끝을 알았으니 이제 선택할수 있게 된다.
df = df.iloc[상관관계row[0]:끝row[0]]



# 현재 nan이 실제 NaN 자료형이 아니라 문자형으로 인식되어서 필요없는 열이 삭제가 안되었었다.
# df.isna().sum()과 df.dtypes로 nan값이 문자형인지 실제 nan값인지 확인해줘야함.
# nan값을 실제 np.nan값으로 먼저 변환해주는 작업이 필요함.
df = df.replace("nan",np.nan)

df = df.dropna(axis=1,how="all").reset_index(drop=True)




# 먼저 각행에 일일이 N값을 표현한건 제거해도 되니까, 그거부터 제거하자.
# N이라는 value가 있는 cell의 row를 np.where로 찾은다음, del 반복문으로 제거해주자.
value_to_find = "N"

nrows = np.where(df.values == value_to_find)[0]

상관df = df.drop(index=nrows).reset_index(drop=True)

value_to_find = "Pearson 상관"
상관rows, 상관cols = np.where(상관df.values == value_to_find)
value_to_find = "유의확률 (양측)"
유의rows, 유의cols = np.where(상관df.values == value_to_find)

# Pearson 상관값이 있는 위치에서 수행
for 상관row in 상관rows: 
    # 앞에 2칸은 제목이기 때문에 생략
    for i in range(2,len(상관df.columns)):
        # 여기서 .363**(.0)처럼 값을 합치는건데, 별표on=True면 합치면안됨)
        if (별표on == False):
             상관df.iloc[상관row,i] = str(상관df.iloc[상관row,i]) + "(" + str(상관df.iloc[상관row+1,i]) + ")"

# 이제 유의확률이 상관계수옆에 표기가 완료되었다.
# 유의확률 행은 필요없으니 제거 
상관df = 상관df.drop(index=유의rows).reset_index(drop=True)




value_to_find = "Pearson 상관"
상관rows, 상관cols = np.where(상관df.values == value_to_find)
상관rows



# 정규 표현식 기반 변환 함수
def modify_values(value):
    if isinstance(value, str):  # 값이 문자열일 경우
        value = re.sub(r"\(nan\)", "", value)  # "(NaN)"만 제거
    if isinstance(value,str) and "**" in value and 별표on == False: # 값이 문자열이고 **가 포함된 경우
        value = re.sub(r"\*\*", "", value)  # ** 삭제
        value = re.sub(r"\(\d*\.?\d*\)", "(<.001)", value)  # 괄호 안 숫자를 (<.001)로 변경
    if isinstance(value, str) and "*" in value and 별표on == False:
        value = re.sub(r"\*", "", value)  # * 삭제
    return value

상관df = 상관df.applymap(modify_values)


# 논문에 복붙하기 편하게, 위쪽에 중복되는 데이터는 nan으로 표시
for 상관row in 상관rows:
    for i in range(2,len(상관df.columns)):
         
        if(i > 상관row):
            상관df.iloc[상관row, i] = np.nan


# 1열은 필요없으니 삭제
del 상관df[상관df.columns[1]]


# 완료~~
상관df.to_excel(f"F_상관_{파일경로}", index=False)
print("Excel로 저장완료")





