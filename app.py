
import io
import re
import math
import datetime as dt
import pandas as pd
import streamlit as st

# ======================
# 설정값
# ======================
BOARD_W_MM = 2440
BOARD_H_MM = 1220
MARGIN_MM = 20
KERF_MM = 3.2

USABLE_W_MM = BOARD_W_MM - 2 * MARGIN_MM
USABLE_H_MM = BOARD_H_MM - 2 * MARGIN_MM
USABLE_AREA_M2 = (USABLE_W_MM * USABLE_H_MM) / 1_000_000


def pick_col(cols, candidates):
    for c in candidates:
        if c in cols:
            return c
    return None


_spec_pat = re.compile(r"^\s*(\d+(\.\d+)?)\s*\*\s*(\d+(\.\d+)?)\s*\*\s*(\d+(\.\d+)?)\s*$")


def parse_spec(s):
    if pd.isna(s):
        return None, None, None
    m = _spec_pat.match(str(s))
    if not m:
        return None, None, None
    return float(m.group(1)), float(m.group(3)), float(m.group(5))


def compute(df_raw):
    df = df_raw.copy()

    col_date = pick_col(df.columns, ["생산일"])
    col_spec = pick_col(df.columns, ["규격상세"])
    col_qty = pick_col(df.columns, ["생산량"])
    col_prod = pick_col(df.columns, ["제품코드"])
    col_color = pick_col(df.columns, ["색상"])
    col_part = pick_col(df.columns, ["부품명"])

    if not all([col_date, col_spec, col_qty, col_prod, col_color, col_part]):
        raise ValueError("필수 컬럼(생산일, 규격상세, 생산량, 제품코드, 색상, 부품명)을 확인하세요.")

    df["생산일"] = pd.to_datetime(df[col_date]).dt.date

    parsed = df[col_spec].apply(parse_spec)
    df["폭_mm"] = parsed.apply(lambda x: x[0])
    df["길이_mm"] = parsed.apply(lambda x: x[1])
    df["두께_mm"] = parsed.apply(lambda x: x[2])

    df["생산량"] = pd.to_numeric(df[col_qty])
    df["단품면적_m2"] = (df["폭_mm"] * df["길이_mm"]) / 1_000_000
    df["총면적_m2"] = df["단품면적_m2"] * df["생산량"]

    sum_df = (
        df.groupby("생산일", as_index=False)
        .agg(순부품면적_m2=("총면적_m2", "sum"))
        .sort_values("생산일")
    )

    sum_df["원장수_장"] = sum_df["순부품면적_m2"].apply(
        lambda x: math.ceil(x / USABLE_AREA_M2) if x > 0 else 0
    )
    sum_df["사용면적_m2"] = sum_df["원장수_장"] * USABLE_AREA_M2
    sum_df["기본로스_m2"] = sum_df["사용면적_m2"] - sum_df["순부품면적_m2"]
    sum_df["자투리사용_m2"] = 0.0
    sum_df["조정로스_m2"] = sum_df["기본로스_m2"]
    sum_df["로스율"] = sum_df["조정로스_m2"] / sum_df["사용면적_m2"]

    return df, sum_df


st.set_page_config(page_title="목재 재단 로스율 계산기", layout="wide")
st.title("목재 재단 로스율 자동 계산기")

uploaded = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

if uploaded:
    df_raw = pd.read_excel(uploaded)
    df_data, df_sum = compute(df_raw)

    st.subheader("생산일별 요약")
    edited = st.data_editor(df_sum, use_container_width=True)

    edited["조정로스_m2"] = (edited["기본로스_m2"] - edited["자투리사용_m2"]).clip(lower=0)
    edited["로스율"] = edited["조정로스_m2"] / edited["사용면적_m2"]

    st.dataframe(edited, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_data.to_excel(writer, index=False, sheet_name="Data")
        edited.to_excel(writer, index=False, sheet_name="요약")

    st.download_button(
        "결과 엑셀 다운로드",
        data=output.getvalue(),
        file_name=f"wood_loss_result_{dt.date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
