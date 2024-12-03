import pandas as pd
from datetime import datetime
import os

# 현재 날짜 가져오기
current_date = datetime.now().strftime("%Y.%m.%d")

# 경로 설정
csv_folder_path = "J:/Lager/EXPORT Local/"
csv_file_path = f"{csv_folder_path}{current_date}.csv"
excel_file_base = f"{csv_folder_path}{current_date}.xlsx"

# Lagerbestand 관련 경로 설정
lagerbestand_folder_path = "J:/Lager/EXPORT Local/"
lagerbestand_file = f"{lagerbestand_folder_path}2024 Lagerbestand.xlsx"
output_file_path_final = f"{lagerbestand_folder_path}{current_date}_final.xlsx"

# 고유 파일 이름 생성 함수
def generate_unique_filename(base_path):
    if not os.path.exists(base_path):
        return base_path
    counter = 2
    while True:
        new_path = base_path.replace(".xlsx", f" ({counter}).xlsx")
        if not os.path.exists(new_path):
            return new_path
        counter += 1

# 1. CSV 파일 읽기 및 필터링
try:
    df = pd.read_csv(
        csv_file_path,
        encoding="Windows-1252",
        sep=";",
        quotechar='"',
        low_memory=False
    )
    print(f"CSV 데이터를 성공적으로 읽어왔습니다. 열 이름: {df.columns.tolist()}")
except Exception as e:
    print(f"CSV 파일 읽기 오류: {e}")
    exit()

# 'Artikel' 열에서 'V'로 시작하는 데이터 제외
if 'Artikel' in df.columns:
    filtered_df = df[~df['Artikel'].str.startswith('V', na=False)]
    print("필터링 성공! 결과 데이터프레임 준비 완료.")
else:
    print("'Artikel' 열이 존재하지 않습니다. CSV 파일의 열 이름을 확인하세요.")
    print(f"현재 열 이름: {df.columns.tolist()}")
    exit()

# KolliBestand가 null인 행 삭제
if "KolliBestand" in filtered_df.columns:
    filtered_df = filtered_df.dropna(subset=["KolliBestand"])
    print("KolliBestand가 null인 행이 삭제되었습니다.")
else:
    print("'KolliBestand' 열이 존재하지 않습니다. 계속 진행합니다.")

# 필요 없는 열 삭제
columns_to_remove = ["NettoVerfügbar", "NettoBestand", "NettoBestellt", "NettoEingeliefert", "NettoReserviert"]
filtered_df = filtered_df.drop(columns=columns_to_remove, errors="ignore")
print(f"열 {columns_to_remove}가 삭제되었습니다.")

# Lagerbestand 데이터 읽기
try:
    lager_df = pd.read_excel(lagerbestand_file, engine="openpyxl")
    print("Lagerbestand 파일의 열 이름:", lager_df.columns.tolist())
    lager_df.columns = lager_df.columns.str.strip()
    if "K_Name" not in lager_df.columns:
        raise KeyError("'K_Name' 열이 존재하지 않습니다.")
except Exception as e:
    print(f"파일 읽기 오류: {e}")
    exit()

# Artikel과 K_Name 병합 및 정렬
filtered_df["Artikel"] = filtered_df["Artikel"].astype(str)
lager_df["K_Name"] = lager_df["K_Name"].astype(str)

merged_df = filtered_df.merge(
    lager_df[["Nummer", "K_Name"]].rename(columns={"K_Name": "New_Column"}),
    how="left",
    left_on="Artikel",
    right_on="Nummer"
)

# A열과 B열 사이에 New_Column 삽입
col_order = ["Artikel", "New_Column"] + [col for col in merged_df.columns if col not in ["Artikel", "New_Column"]]
merged_df = merged_df[col_order]

# Lieferanten.Name별로 시트 생성 및 저장
try:
    output_file_path_final = generate_unique_filename(output_file_path_final)
    with pd.ExcelWriter(output_file_path_final, engine="openpyxl") as writer:
        unique_lieferanten = merged_df["Lieferanten.Name"].dropna().unique()
        for lieferant in unique_lieferanten:
            sheet_data = merged_df[merged_df["Lieferanten.Name"] == lieferant]
            sheet_data = sheet_data.sort_values(by="Artikel")  # Artikel 기준 오름차순 정렬
            sheet_name = str(lieferant)[:30]  # 시트 이름은 30자 제한
            sheet_data.to_excel(writer, index=False, sheet_name=sheet_name)
    print(f"최종 데이터가 {output_file_path_final}에 저장되었습니다.")
except Exception as e:
    print(f"최종 Excel 저장 중 오류 발생: {e}")
