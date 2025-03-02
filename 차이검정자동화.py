#%% 1. 초기 설정 및 데이터 로드
import numpy as np
import pandas as pd
import re

# 파일 경로 설정
파일경로 = "차이검정.xlsx"
df = pd.read_excel(파일경로, header=None)

# 숫자형 데이터를 소수점 2자리까지 변환, 0도 유지
df = df.map(lambda x: f"{x:.2f}" if isinstance(x, (float, int)) else x)

# 결측치(nan) 처리 (0은 건드리지 않도록 함)
df = df.replace({"nan": np.nan, "NaN": np.nan, np.nan: ""})

#%% 2. 테이블 추출
# 찾을 키워드 목록
keywords = ["집단통계량", "독립표본 검정", "기술통계", "ANOVA", "Scheffe"]  # Scheffe만 추가

# 결과 저장을 위한 딕셔너리
table_dict = {keyword: [] for keyword in keywords}

# 키워드별로 표를 자동으로 추출하여 저장
for keyword in keywords:
    start_indices = df[df.iloc[:, 0].astype(str).str.contains(keyword, na=False)].index.tolist()
    
    for start_idx in start_indices:
        end_idx = start_idx + 1 
        
        # Scheffe 결과 테이블의 끝 찾기
        if keyword == "Scheffe":
            for idx in range(start_idx + 1, len(df)):
                if "CTT 유의확률" in str(df.iloc[idx, 0]):
                    end_idx = idx + 1
                    break
        else:
            # 기존 방식대로 빈 행 찾기
            for idx in range(start_idx + 1, len(df)):
                if df.iloc[idx].isnull().all():
                    end_idx = idx
                    break

        table = df.iloc[start_idx:end_idx].dropna(how="all").reset_index(drop=True)
        table_dict[keyword].append(table)

# 사후검정 결과 저장을 위한 딕셔너리
사후검정_결과 = {}

# Scheffe 결과 처리
if "Scheffe" in table_dict:
    for table in table_dict["Scheffe"]:
        # 변수명 찾기 (테이블 위의 행)
        변수명 = None
        for idx in range(len(table)):
            if "유의수준 = 0.05에 대한 부분집합" in str(table.iloc[idx].values):
                변수명 = table.iloc[0, 0]  # 첫 번째 열의 변수명
                break
        
        if 변수명:
            사후검정_결과[변수명] = {
                'groups': [],  # 그룹 정보 저장
                'subsets': []  # 부분집합 정보 저장
            }
            
            # 그룹과 부분집합 정보 추출
            for idx in range(len(table)):
                row = table.iloc[idx]
                if pd.notna(row[0]) and "CTT" not in str(row[0]):
                    group = row[0]  # 그룹명
                    n = row[1]      # N
                    subset1 = row[2] if pd.notna(row[2]) else None  # 부분집합 1
                    subset2 = row[3] if pd.notna(row[3]) else None  # 부분집합 2
                    
                    사후검정_결과[변수명]['groups'].append({
                        'name': group,
                        'n': n,
                        'subset1': subset1,
                        'subset2': subset2
                    })

#%% 3. 변수 추출 및 저장
# 개별 변수 저장
집단통계량리스트 = table_dict["집단통계량"]
독립표본검정리스트 = table_dict["독립표본 검정"]
기술통계리스트 = table_dict["기술통계"]
ANOVA리스트 = table_dict["ANOVA"]

# 종속변수 추출
변수이름리스트 = []
for 집단통계량 in 집단통계량리스트:
    변수이름리스트.append(집단통계량[0][1])

# 집단통계량에서 종속변수 목록 추출
unique_values = 집단통계량리스트[0][0].dropna().unique()
unique_values = unique_values[unique_values != "집단통계량"]
종속변수리스트 = unique_values[unique_values != 변수이름리스트[0]]

#%% 4. 카테고리 추출
# 원본 파일에서 카테고리 순서 추출
카테고리_리스트 = []
for idx in range(len(df)):
    row = df.iloc[idx]
    # 집단통계량 또는 기술통계 테이블 내의 카테고리만 추출
    if pd.notna(row[1]):
        # 집단통계량 테이블에서 추출
        for 집단통계량 in 집단통계량리스트:
            if row[1] in 집단통계량[1].dropna().unique():
                if row[1] not in 카테고리_리스트 and row[1] != "전체":
                    카테고리_리스트.append(row[1])
                break
        
        # 기술통계 테이블에서 추출
        for 기술통계 in 기술통계리스트:
            if row[1] in 기술통계[1].dropna().unique():
                if row[1] not in 카테고리_리스트 and row[1] != "전체":
                    카테고리_리스트.append(row[1])
                break

#%% 5. n값 추출
# n값 저장을 위한 딕셔너리
n_values = {}

# 집단통계량에서 n값 추출
for 집단통계량 in 집단통계량리스트:
    current_categories = 집단통계량[1].dropna().unique().tolist()
    for i, category in enumerate(current_categories):
        if category in 카테고리_리스트:
            n_values[category] = 집단통계량.iloc[i, 2]  # n값은 보통 3번째 열에 있음

# 기술통계에서 n값 추출
for 기술통계 in 기술통계리스트:
    current_categories = [cat for cat in 기술통계[1].dropna().unique() if cat != "전체"]
    for i, category in enumerate(current_categories):
        if category in 카테고리_리스트:
            n_values[category] = 기술통계.iloc[i, 2]  # n값은 보통 3번째 열에 있음

#%% 6. 데이터 처리
def get_group_labels(category, variable):
    """카테고리에 그룹 라벨 추가"""
    if variable not in 사후검정_결과:
        return category
    
    # 해당 변수의 사후검정 결과 가져오기
    groups = 사후검정_결과[variable]['groups']
    
    # 현재 카테고리가 그룹에 있는지 확인
    for group in groups:
        if category in group['name']:
            # subset1과 subset2를 확인하여 그룹 라벨 결정
            label = None
            if group['subset1'] is not None:
                label = 'a'
            if group['subset2'] is not None:
                label = 'b' if label is None else f"{label}b"
            
            if label:
                return f"{category}({label})"
    
    return category

# 종속변수별 빈 리스트 생성 (각 카테고리마다 기본값 설정)
종속변수딕셔너리 = {변수: ["0.00±0.00"] * len(카테고리_리스트) for 변수 in 종속변수리스트}

# 집단통계량(독립검정)에서 값 추가
for index, 집단통계량 in enumerate(집단통계량리스트):
    current_categories = 집단통계량[1].dropna().unique().tolist()
    for key, value in 종속변수딕셔너리.items():
        row, col = np.where(집단통계량.values == key)
        if len(row) > 0:  # 해당 키가 존재할 경우에만
            for i in range(len(current_categories)):
                category = current_categories[i]
                if category in 카테고리_리스트:  # 카테고리가 리스트에 있는 경우에만
                    category_index = 카테고리_리스트.index(category)
                    mean_val = 집단통계량.iloc[row[0] + i, 3]
                    std_val = 집단통계량.iloc[row[0] + i, 4]
                    value[category_index] = f"{mean_val}±{std_val}"

# 검정통계량 추가를 위한 딕셔너리
검정통계량딕셔너리 = {}

# 독립표본 t검정 결과 추출
for 독립검정 in 독립표본검정리스트:
    for key in 종속변수딕셔너리.keys():
        try:
            # 변수명이 포함된 행 찾기
            for idx in range(len(독립검정)):
                if key in str(독립검정.iloc[idx].values):
                    # Levene의 등분산 검정 결과 확인
                    levene_sig = 독립검정.iloc[idx, 1]  # 유의확률
                    
                    if float(levene_sig) < 0.05:
                        # 등분산을 가정하지 않음
                        t_value = 독립검정.iloc[idx, 8]  # 등분산을 가정하지 않음의 t값
                        p_value = 독립검정.iloc[idx, 9]  # 등분산을 가정하지 않음의 유의확률(양측)
                    else:
                        # 등분산을 가정함
                        t_value = 독립검정.iloc[idx, 6]  # 등분산을 가정함의 t값
                        p_value = 독립검정.iloc[idx, 7]  # 등분산을 가정함의 유의확률(양측)
                    
                    if pd.notna(t_value) and pd.notna(p_value):
                        # p값에 따른 별표 추가
                        stars = ""
                        if float(p_value) < 0.001:
                            stars = "***"
                        elif float(p_value) < 0.01:
                            stars = "**"
                        elif float(p_value) < 0.05:
                            stars = "*"
                        검정통계량딕셔너리[key] = f"{float(t_value):.3f}({float(p_value):.3f}){stars}"
                    break
        except (ValueError, TypeError):
            continue

# 기술통계(ANOVA)에서 값 추가
for index, 기술통계 in enumerate(기술통계리스트):
    current_categories = [cat for cat in 기술통계[1].dropna().unique() if cat != "전체"]
    for key, value in 종속변수딕셔너리.items():
        row, col = np.where(기술통계.values == key)
        if len(row) > 0:  # 해당 키가 존재할 경우에만
            for i in range(len(current_categories)):
                category = current_categories[i]
                if category in 카테고리_리스트:  # 카테고리가 리스트에 있는 경우에만
                    category_index = 카테고리_리스트.index(category)
                    try:
                        mean_val = 기술통계.iloc[row[0] + i, 3]
                        std_val = 기술통계.iloc[row[0] + i, 4]
                        value[category_index] = f"{mean_val}±{std_val}"
                    except (IndexError, KeyError):
                        continue

# ANOVA 결과 추출 
for anova in ANOVA리스트:
    for key in 종속변수딕셔너리.keys():
        try:
            # 변수명이 포함된 행 찾기
            for idx in range(len(anova)):
                if key in str(anova.iloc[idx].values):
                    f_value = anova.iloc[idx, 2]  # F값
                    p_value = anova.iloc[idx, -1]  # p값
                    if pd.notna(f_value) and pd.notna(p_value):
                        # p값에 따른 별표 추가
                        stars = ""
                        if float(p_value) < 0.001:
                            stars = "***"
                        elif float(p_value) < 0.01:
                            stars = "**"
                        elif float(p_value) < 0.05:
                            stars = "*"
                        검정통계량딕셔너리[key] = f"{float(f_value):.3f}({float(p_value):.3f}){stars}"
                        break
        except (ValueError, TypeError):
            continue

# 사후검정 결과 추가
for key in 검정통계량딕셔너리.keys():
    if key in 사후검정_결과:
        # 사후검정 결과가 있는 경우 괄호 안에 추가
        post_hoc_pairs = []
        for group in 사후검정_결과[key]['groups']:
            if group['subset1'] is not None and group['subset2'] is not None:
                post_hoc_pairs.append(f"a<c")  # 실제 비교 결과에 따라 수정 필요
        if post_hoc_pairs:
            검정통계량딕셔너리[key] += f" ({','.join(post_hoc_pairs)})"

#%% 6. 데이터프레임 생성 및 포맷팅
# 새 데이터프레임 생성
rows = []
current_variable = None

# 변수명과 카테고리 헤더 행 추가
header_row = ["Variables", "Categories"]
for var in 종속변수리스트:
    header_row.extend([var, ""])  # M±SD와 t or F(p)를 위한 빈 열 추가
rows.append(header_row)

# 하위 헤더 행 추가
subheader_row = ["", ""]
for _ in 종속변수리스트:
    subheader_row.extend(["M±SD", "t or F(p)"])
rows.append(subheader_row)

# 데이터 행 추가
for category in 카테고리_리스트:
    row_data = ["", get_group_labels(category, None)]
    
    # 각 종속변수별 데이터 추가
    for var in 종속변수리스트:
        category_index = 카테고리_리스트.index(category)
        row_data.extend([
            종속변수딕셔너리[var][category_index],  # M±SD
            검정통계량딕셔너리.get(var, "")  # t or F(p)
        ])
    rows.append(row_data)

# 데이터프레임 생성
df_expanded = pd.DataFrame(rows)

# 첫 번째 행을 멀티인덱스 컬럼으로 설정
df_expanded.columns = pd.MultiIndex.from_arrays([
    df_expanded.iloc[0], 
    df_expanded.iloc[1]
])

# 헤더로 사용된 행 제거
df_expanded = df_expanded.iloc[2:].reset_index(drop=True)

#%% 7. 결측치 처리 및 저장
# "nan"을 "0.00"으로 변경
df_expanded = df_expanded.replace("nan", "0.00")
df_expanded = df_expanded.replace("nan±nan", "0.00±0.00")

# "±nan" 또는 "±NaN"이 포함된 값을 "±0.00"으로 변경
df_expanded = df_expanded.replace(to_replace=r"±\s*nan", value="±0.00", regex=True)
df_expanded = df_expanded.replace(to_replace=r"±\s*NaN", value="±0.00", regex=True)

# 엑셀 저장 (MultiIndex 컬럼 처리)
with pd.ExcelWriter(f'F_차이검정_{파일경로}', engine='openpyxl') as writer:
    df_expanded.to_excel(writer, index=True, index_label=None)  # index=True로 변경
    
    # 인덱스 열 숨기기
    worksheet = writer.sheets['Sheet1']
    worksheet.column_dimensions['A'].hidden = True

print("엑셀 저장 완료 ✅") 