
import io
import re
from datetime import datetime
import pandas as pd
import streamlit as st

st.set_page_config(page_title="수율(로스율) 프로그램", layout="wide")

DEFAULT_SHEET_W = 1220.0
DEFAULT_SHEET_H = 2440.0
DEFAULT_MARGIN  = 10.0
DEFAULT_BLADE_MM = 3.2
DEFAULT_KERF_MM  = 20.0

def parse_spec(spec: str):
    if spec is None or (isinstance(spec, float) and pd.isna(spec)):
        return (None, None, None)
    s = str(spec).strip()
    parts = re.split(r'[*xX×]', s)
    nums = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        m = re.findall(r'\d+(?:\.\d+)?', p)
        if m:
            nums.append(float(m[0]))
    if len(nums) >= 2:
        return (nums[0], nums[1], nums[2] if len(nums) >= 3 else None)
    return (None, None, None)

def build_excel(sheets: dict):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=name[:31])
    out.seek(0)
    return out

def fmt4(x):
    try:
        return f"{float(x):,.4f}"
    except Exception:
        return ""

st.title("수율(로스율) 프로그램 · v3 (ASCII 안전버전)")
st.caption("Streamlit Cloud 오류(U+33A1 ㎡ 문자) 수정 버전")

with st.sidebar:
    st.header("원장 기준")
    sheet_w = st.number_input("원장 가로(W) [mm]", value=float(DEFAULT_SHEET_W))
    sheet_h = st.number_input("원장 세로(H) [mm]", value=float(DEFAULT_SHEET_H))
    margin  = st.number_input("Margin [mm]", value=float(DEFAULT_MARGIN))

    st.divider()
    st.header("엑셀 컬럼 매핑")
    col_date = st.text_input("생산일", value="생산일")
    col_spec = st.text_input("규격상세", value="규격상세")
    col_qty  = st.text_input("생산량", value="생산량")

uploaded = st.file_uploader("엑셀 업로드 (.xlsx)", type=["xlsx"])
if not uploaded:
    st.stop()

df = pd.read_excel(uploaded)

df["수량"] = pd.to_numeric(df[col_qty], errors="coerce").fillna(0)
parsed = df[col_spec].apply(parse_spec)
df[["W_raw","H_raw","T_raw"]] = pd.DataFrame(parsed.tolist(), index=df.index)

df["area_mm2"] = df["W_raw"] * df["H_raw"] * df["수량"]
df["area_m2"] = df["area_mm2"] / 1_000_000

df["_생산일"] = pd.to_datetime(df[col_date], errors="coerce").dt.date
dates = sorted(df["_생산일"].dropna().unique())

rows = ["생산량 면적(m2)"]
table = pd.DataFrame(index=rows)

for d in dates:
    sub = df[df["_생산일"] == d]
    total = sub["area_m2"].sum()
    table[d] = [total]

overall = df["area_m2"].sum()
table["합계"] = [overall]

disp = table.copy()
for c in disp.columns:
    disp.loc["생산량 면적(m2)", c] = fmt4(disp.loc["생산량 면적(m2)", c])

disp = disp.reset_index().rename(columns={"index": "구분"})

st.subheader("1. 첫페이지 (생산일별 + 합계)")
st.dataframe(disp, use_container_width=True, hide_index=True)

st.divider()
st.subheader("엑셀 다운로드")
result = build_excel({"1_첫페이지": table.reset_index().rename(columns={"index":"구분"})})
st.download_button(
    "엑셀 다운로드",
    data=result,
    file_name=f"수율_결과_v3_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
