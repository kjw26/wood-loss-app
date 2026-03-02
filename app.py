
import io
import re
import math
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

def shelf_pack_count(rects, W_eff, H_eff, gap=0.0, allow_rotate=True):
    if not rects:
        return 0
    rects_sorted = sorted(rects, key=lambda x: max(x[0], x[1]) * 100000 + min(x[0], x[1]), reverse=True)

    sheets = 1
    x = 0.0
    y = 0.0
    shelf_h = 0.0

    for (w0, h0) in rects_sorted:
        candidates = []
        ori = [(w0, h0)]
        if allow_rotate and w0 != h0:
            ori.append((h0, w0))

        for (w, h) in ori:
            wi = w + gap
            hi = h + gap
            if wi <= W_eff and hi <= H_eff:
                candidates.append((wi, hi))

        if not candidates:
            return None

        placed = False
        candidates.sort(key=lambda t: t[0])
        for (wi, hi) in candidates:
            if x + wi <= W_eff and y + max(shelf_h, hi) <= H_eff:
                x += wi
                shelf_h = max(shelf_h, hi)
                placed = True
                break
        if placed:
            continue

        x = 0.0
        y += shelf_h
        shelf_h = 0.0
        for (wi, hi) in candidates:
            if x + wi <= W_eff and y + hi <= H_eff:
                x += wi
                shelf_h = max(shelf_h, hi)
                placed = True
                break
        if placed:
            continue

        sheets += 1
        x = candidates[0][0]
        y = 0.0
        shelf_h = candidates[0][1]

    return sheets

def build_excel(sheets: dict):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=name[:31])
    out.seek(0)
    return out

def to_date_series(s):
    try:
        return pd.to_datetime(s, errors="coerce").dt.date
    except Exception:
        return pd.Series([None]*len(s))

def fmt4(x):
    try:
        return f"{float(x):,.4f}"
    except Exception:
        return ""

st.title("수율(로스율) 프로그램 · 생산일별 표 + 합계 (v2)")
st.caption("요청 반영: 1페이지에서 생산일별로 표기 + 합계 표시, 자투리 사용 시 로스율이 '감소'하도록 로직 수정")

with st.sidebar:
    st.header("원장/가공 기준")
    sheet_w = st.number_input("원장 가로(W) [mm]", value=float(DEFAULT_SHEET_W), step=1.0)
    sheet_h = st.number_input("원장 세로(H) [mm]", value=float(DEFAULT_SHEET_H), step=1.0)
    margin  = st.number_input("Margin(테두리 여유) [mm]", value=float(DEFAULT_MARGIN), step=1.0)

    st.divider()
    st.header("절단/간격")
    blade_mm = st.number_input("톱날 두께 [mm]", value=float(DEFAULT_BLADE_MM), step=0.1)
    kerf_mm  = st.number_input("Kerf(부품 간 여유) [mm]", value=float(DEFAULT_KERF_MM), step=1.0)

    st.divider()
    st.header("방향/회전")
    allow_rotate = st.checkbox("회전(90도) 허용", value=True)
    auto_long_to_h = st.checkbox("긴변을 세로(H)로 자동 정렬", value=True)

    st.divider()
    st.header("엑셀 컬럼 매핑")
    col_date = st.text_input("생산일 컬럼", value="생산일")
    col_prod = st.text_input("제품코드 컬럼", value="제품코드")
    col_color = st.text_input("색상 컬럼", value="색상")
    col_part = st.text_input("부품명 컬럼", value="부품명")
    col_spec = st.text_input("규격상세 컬럼", value="규격상세")
    col_thk  = st.text_input("자재두께 컬럼(없으면 규격상세 3번째 사용)", value="자재두께(mm)")
    col_qty  = st.text_input("생산량 컬럼", value="생산량")

uploaded = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx"])
if not uploaded:
    st.info("엑셀 파일을 업로드하면 결과가 표시됩니다.")
    st.stop()

try:
    df = pd.read_excel(uploaded)
except Exception as e:
    st.error(f"엑셀 읽기 오류: {e}")
    st.stop()

need_cols = [col_date, col_prod, col_color, col_part, col_spec, col_qty]
missing = [c for c in need_cols if c not in df.columns]
if missing:
    st.error(f"필수 컬럼이 없습니다: {missing}\n사이드바의 컬럼 매핑을 엑셀에 맞게 수정해 주세요.")
    st.stop()

df["_생산일"] = to_date_series(df[col_date])
dates = sorted([d for d in df["_생산일"].dropna().unique().tolist()])
if not dates:
    st.warning("생산일 컬럼을 날짜로 인식하지 못했습니다. (전체 합계만 계산합니다)")
    df["_생산일"] = "전체"
    dates = ["전체"]

df["수량"] = pd.to_numeric(df[col_qty], errors="coerce").fillna(0).astype(int)
parsed = df[col_spec].apply(parse_spec)
df[["W_raw","H_raw","T_raw"]] = pd.DataFrame(parsed.tolist(), index=df.index)

if col_thk in df.columns:
    df["자재두께(mm)"] = pd.to_numeric(df[col_thk], errors="coerce")
else:
    df["자재두께(mm)"] = df["T_raw"]

def norm_wh(row):
    w, h = row["W_raw"], row["H_raw"]
    if pd.isna(w) or pd.isna(h):
        return pd.Series({"가로(mm)": None, "세로(mm)": None})
    if auto_long_to_h and w > h:
        return pd.Series({"가로(mm)": h, "세로(mm)": w})
    return pd.Series({"가로(mm)": w, "세로(mm)": h})

df[["가로(mm)","세로(mm)"]] = df.apply(norm_wh, axis=1)

W_eff = float(sheet_w - 2*margin)
H_eff = float(sheet_h - 2*margin)
A_sheet = W_eff * H_eff
A_sheet_m2 = A_sheet / 1e6
gap = float(kerf_mm + blade_mm)

def can_fit(w, h):
    if w is None or h is None:
        return False
    if w <= W_eff and h <= H_eff:
        return True
    if allow_rotate and (h <= W_eff and w <= H_eff):
        return True
    return False

df_valid = df[(df["수량"] > 0) & df["가로(mm)"].notna() & df["세로(mm)"].notna()].copy()
df_invalid = df[~((df["수량"] > 0) & df["가로(mm)"].notna() & df["세로(mm)"].notna())].copy()
df_oversize = df_valid[~df_valid.apply(lambda r: can_fit(r["가로(mm)"], r["세로(mm)"]), axis=1)].copy()
df_ok = df_valid.drop(index=df_oversize.index, errors="ignore").copy()

df_ok["재단면적_mm2"] = df_ok["가로(mm)"] * df_ok["세로(mm)"] * df_ok["수량"]
df_ok["재단면적(㎡)"] = df_ok["재단면적_mm2"] / 1e6

with st.sidebar:
    st.divider()
    st.header("자투리 보관 사용(생산일별 입력)")
    st.caption("자투리 사용은 '신규 자재투입을 대체'한다고 가정하면 로스율이 줄어듭니다.")
    scrap_mode = st.radio(
        "자투리 반영 방식",
        options=["신규 자재투입을 대체(로스율↓)", "투입면적에 더함(참고용)"],
        index=0
    )
    scrap_df = pd.DataFrame({"생산일": dates, "자투리보관사용(㎡)": [0.0]*len(dates)})
    scrap_df = st.data_editor(scrap_df, num_rows="fixed", use_container_width=True, hide_index=True)
    scrap_map = {r["생산일"]: float(r["자투리보관사용(㎡)"] or 0.0) for _, r in scrap_df.iterrows()}

def compute_by_date(d):
    sub = df_ok[df_ok["_생산일"] == d].copy() if d != "전체" else df_ok.copy()
    cut_m2 = float(sub["재단면적(㎡)"].sum())
    rects = []
    for _, r in sub.iterrows():
        rects.extend([(float(r["가로(mm)"]), float(r["세로(mm)"]))] * int(r["수량"]))
    sheets = shelf_pack_count(rects, W_eff, H_eff, gap=gap, allow_rotate=allow_rotate) if rects else 0
    if sheets is None:
        sheets = 0
    input_m2 = float(sheets) * A_sheet_m2

    scrap_used = float(scrap_map.get(d, 0.0))
    if scrap_mode.startswith("신규"):
        input_eff = max(input_m2 - scrap_used, 0.0)
    else:
        input_eff = input_m2 + scrap_used

    loss_m2 = max(input_eff - cut_m2, 0.0)
    loss_pct = (loss_m2 / input_eff * 100) if input_eff > 0 else 0.0

    return {"생산량 면적": cut_m2, "자재투입 면적": input_eff, "자투리 보관 사용": scrap_used, "loss율": loss_pct, "_원장수": sheets, "_원장투입면적": input_m2}

rows = ["생산량 면적", "자재투입 면적", "자투리 보관 사용", "loss율"]
table = pd.DataFrame(index=rows)
for d in dates:
    r = compute_by_date(d)
    table[d] = [r["생산량 면적"], r["자재투입 면적"], r["자투리 보관 사용"], r["loss율"]]

# 합계(전체 물량 기준으로 재계산)
overall = compute_by_date("전체")
table["합계"] = [overall["생산량 면적"], overall["자재투입 면적"], sum(scrap_map.values()), overall["loss율"]]

disp = table.copy()
for c in disp.columns:
    disp.loc["생산량 면적", c] = fmt4(disp.loc["생산량 면적", c])
    disp.loc["자재투입 면적", c] = fmt4(disp.loc["자재투입 면적", c])
    disp.loc["자투리 보관 사용", c] = fmt4(disp.loc["자투리 보관 사용", c])
    disp.loc["loss율", c] = fmt4(disp.loc["loss율", c])
disp = disp.reset_index().rename(columns={"index": "구분"})

sel_date = st.selectbox("상세 조회 생산일 선택", options=dates, index=len(dates)-1)
sub_ok = df_ok[df_ok["_생산일"] == sel_date].copy() if sel_date != "전체" else df_ok.copy()
sel_calc = compute_by_date(sel_date)
cut_total = float(sub_ok["재단면적(㎡)"].sum())
loss_m2_sel = max(sel_calc["자재투입 면적"] - sel_calc["생산량 면적"], 0.0)

group_cols = [col_prod, col_color, col_part, "자재두께(mm)"]
detail = (sub_ok.groupby(group_cols, dropna=False, as_index=False)
          .agg(수량=("수량","sum"),
               재단면적_㎡=("재단면적(㎡)","sum")))
detail["재단 점유율(%)"] = (detail["재단면적_㎡"] / cut_total * 100) if cut_total > 0 else 0.0
detail["로스면적(㎡)"] = loss_m2_sel * (detail["재단면적_㎡"] / cut_total) if cut_total > 0 else 0.0
detail["로스율(%)"] = (detail["로스면적(㎡)"] / (detail["로스면적(㎡)"] + detail["재단면적_㎡"]) * 100).fillna(0.0)

detail = detail.rename(columns={
    col_prod: "제품코드",
    col_color: "색상",
    col_part: "부품명",
    "재단면적_㎡": "재단면적(㎡)"
}).sort_values("재단면적(㎡)", ascending=False)

an_part = (detail.groupby(["부품명"], as_index=False)
           .agg(수량=("수량","sum"),
                재단면적_㎡=("재단면적(㎡)","sum"),
                로스면적_㎡=("로스면적(㎡)","sum")))
an_part["로스 기여율(%)"] = (an_part["로스면적_㎡"] / loss_m2_sel * 100) if loss_m2_sel > 0 else 0.0
an_part = an_part.sort_values("로스면적_㎡", ascending=False)

tab1, tab2, tab3, tab4 = st.tabs(["1. 첫페이지", "2 페이지", "3 페이지", "오류/예외"])

with tab1:
    st.subheader("1. 첫페이지 (생산일별 + 합계)")
    st.dataframe(disp, use_container_width=True, hide_index=True)

with tab2:
    st.subheader(f"2 페이지 (생산일: {sel_date})")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("생산량 면적(㎡)", fmt4(sel_calc["생산량 면적"]))
    c2.metric("자재투입 면적(㎡)", fmt4(sel_calc["자재투입 면적"]))
    c3.metric("자투리 보관 사용(㎡)", fmt4(sel_calc["자투리 보관 사용"]))
    c4.metric("loss율(%)", fmt4(sel_calc["loss율"]))
    st.dataframe(detail, use_container_width=True, hide_index=True)

with tab3:
    st.subheader(f"3 페이지 (생산일: {sel_date})")
    st.dataframe(an_part, use_container_width=True, hide_index=True)
    st.bar_chart(an_part.set_index("부품명")["로스면적_㎡"].head(20))

with tab4:
    st.subheader("오류/예외")
    if not df_invalid.empty:
        st.warning("규격 파싱 실패/수량 0 등으로 제외된 행")
        st.dataframe(df_invalid[[col_date, col_prod, col_color, col_part, col_spec, col_qty]].copy(),
                     use_container_width=True, hide_index=True)
    if not df_oversize.empty:
        st.error("원장 유효치수 초과(배치 불가) 행")
        show = df_oversize[[col_date, col_prod, col_color, col_part, "가로(mm)", "세로(mm)", "수량", col_spec]].copy()
        st.dataframe(show, use_container_width=True, hide_index=True)
    if df_invalid.empty and df_oversize.empty:
        st.success("오류/예외가 없습니다.")

st.divider()
st.subheader("결과 다운로드(엑셀)")

firstpage_numeric = table.reset_index().rename(columns={"index": "구분"})
sheets = {
    "1_첫페이지(표시)": disp,
    "1_첫페이지(수치)": firstpage_numeric,
    "2_상세": detail,
    "3_부품별_로스기여": an_part,
}
result = build_excel(sheets)

st.download_button(
    "엑셀 다운로드",
    data=result,
    file_name=f"수율_로스율_결과_v2_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
