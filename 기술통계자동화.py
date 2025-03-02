
# 스크립트가 실행되고 나면 기술통계표를 한글에 간편하게 데이터 붙여넣기로 옮기기 쉽게 만들어주게 된다.
# 준비물 : 같은 폴더내에 기술통계 분석한 파일 xlsx형태로 있어야함.
# 참고로 spss에서 기술통계량을 분석할때, 
# 보편적으로 필요한 (평균 표준편차 최소값 최대값 첨도 왜도) 에 체크하고 분석한 상태, 다르게 분석할 시 오류가 발생할 수 있음을 유의

import numpy as np
import pandas as pd



# 파일을 불러온다. 파일은 같은 폴더내에 있어야한다.
파일경로 = "주요요인_기술통계.xlsx"
df = pd.read_excel(파일경로, header=None)



# 데이터프레임 사이에 NaN행을 제거한 복사본 생성, index 초기화
df = df.dropna(how='all').reset_index(drop=True)

# 숫자형 데이터를 소수점 2자리까지 반올림
df = df.applymap(lambda x: round(x,2) if isinstance(x, float) else x)



# 0열에 "기술통계량"이 있으면 그때부터 기술통계 표가 나오게된다. 
value_to_find = "기술통계량"
# np.where를 통해서 기술통계량의 행과 열을 찾는다. 참고로 where로 찾은 데이터는 리스트형태로 저장된다.
기술통계량row, 기술통계량col = np.where(df.values == value_to_find)



# 통계 표의 마지막 행의 0열에는 "유효 N(목록별)"로 시작하게된다. 따라서 그부분을 찾으면 기술통계표만 따로 저장할수 있게됨.
value_to_find = "유효 N(목록별)"

유효Nrow, 유효Ncol = np.where(df.values == value_to_find)


# 현재 찾은 데이터들로 새로운 df를 만든다. iloc함수는 첫번째인자가 행, 두번째인자가 열이다.
df = df.iloc[기술통계량row[0]:유효Nrow[0], :10].reset_index(drop=True)





# 보편적으로 기술통계량에서 평균과 표준편차만 +-로 같이 묶어서 표현하는 것 같으므로, 그것을 수행할거다.
# 평균 통계량과 표준편차 통계량의 col값을 찾아준다.
value_to_find = "평균"

평균row, 평균col = np.where(df.values == value_to_find)

value_to_find = "표준편차"
표준편차row,표준편차col = np.where(df.values == value_to_find)

# 찾은 col값을 기반으로 두 데이터를 합쳐서 나타낸다.
# 논문마다 표현법이 상이 하므로, 만약 ± 가 아닌 다른 표현을 작성해야 된다면 아래의 코드에서 원하는 표현방식으로 바꿔주면된다.
df[평균col[0]] = df[평균col[0]].astype(str) + "±" + df[표준편차col[0]].astype(str)




# 사실 기술통계는 입력하는 시간이 오래걸리는 작업이 아니라서 더 건드릴것도 없는 것 같다.
# 가공이 완료된 데이터프레임은 가공된기술통계분석.xlsx 라는 이름으로 같은 폴더내에 저장된다.

df.to_excel(f"F_기술_{파일경로}", index=False)
print("Excel로 저장완료")


