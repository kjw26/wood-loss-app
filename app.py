
import io
import re
import math
from datetime import datetime
import pandas as pd
import streamlit as st

st.set_page_config(page_title="수율(로스율) 프로그램", layout="wide")

# =============================
# 기본값
# =============================
DEFAULT_SHEET_W = 1220.0
DEFAULT_SHEET_H = 2440.0
DEFAULT_MARGIN  = 10.0
DEFAULT_BLADE_MM = 3.2
DEFAULT_KERF_MM  = 20.0

# =============================
# 유틸
# =============================
def parse_spec(spec: str):
    """'1086*394*18' -> (W, H, T). 구분자: *, x, X, ×"""
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
    """선반(shelf) 방식으로 원장 장수 근사"""
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
        candidates.sort(key=lambda t: t[0])  # 폭 작은 것 우선
        for (wi, hi) in candidates:
            if x + wi <= W_eff and y + max(shelf_h, hi) <= H_eff:
                x += wi
                shelf_h = max(shelf_h, hi)
                placed = True
                break
        if placed:
            continue

        # 새 선반
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

        # 새 원장
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
    # 엑셀 날짜/문자 혼용 안전 처리
    try:
        return pd.to_datetime(s, errors="coerce").dt.date
    except Exception:
        return pd.Series([None]*len(s))

# =============================
# UI: 사이드바
# =============================
st.title("수율(로스율) 프로그램 · 템플릿 형식")
st.caption("첨부하신 엑셀 형식(첫페이지 요약 / 2페이지 상세 / 3페이지 분석)에 맞춘 Streamlit 버전입니다.")

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

# =============================
# 데이터 로드/검증
# =============================
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

# 날짜 선택
df["_생산일"] = to_date_series(df[col_date])
dates = sorted([d for d in df["_생산일"].dropna().unique().tolist()])
sel_date = None
if dates:
    sel_date = st.selectbox("구분(생산일 기준)", options=dates, index=len(dates)-1)
    df = df[df["_생산일"] == sel_date].copy()
else:
    st.warning("생산일 컬럼을 날짜로 인식하지 못했습니다. (전체 데이터로 계산합니다)")

# 수량/규격 파싱
df["수량"] = pd.to_numeric(df[col_qty], errors="coerce").fillna(0).astype(int)
parsed = df[col_spec].apply(parse_spec)
df[["W_raw","H_raw","T_raw"]] = pd.DataFrame(parsed.tolist(), index=df.index)

# 두께 결정: (1) 두께 컬럼이 있으면 우선, (2) 규격상세 3번째
if col_thk in df.columns:
    df["두께(mm)"] = pd.to_numeric(df[col_thk], errors="coerce")
else:
    df["두께(mm)"] = df["T_raw"]

# 유효치수/정렬
def norm_wh(row):
    w, h = row["W_raw"], row["H_raw"]
    if pd.isna(w) or pd.isna(h):
        return pd.Series({"가로(mm)": None, "세로(mm)": None})
    if auto_long_to_h and w > h:
        return pd.Series({"가로(mm)": h, "세로(mm)": w})
    return pd.Series({"가로(mm)": w, "세로(mm)": h})

df[["가로(mm)","세로(mm)"]] = df.apply(norm_wh, axis=1)

# 원장 유효치수
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

# 유효 데이터
df_valid = df[(df["수량"] > 0) & df["가로(mm)"].notna() & df["세로(mm)"].notna()].copy()
df_invalid = df[~((df["수량"] > 0) & df["가로(mm)"].notna() & df["세로(mm)"].notna())].copy()
df_oversize = df_valid[~df_valid.apply(lambda r: can_fit(r["가로(mm)"], r["세로(mm)"]), axis=1)].copy()
df_ok = df_valid.drop(index=df_oversize.index, errors="ignore").copy()

# 면적
df_ok["재단면적_mm2"] = df_ok["가로(mm)"] * df_ok["세로(mm)"] * df_ok["수량"]
df_ok["재단면적(m2)"] = df_ok["재단면적_mm2"] / 1e6

total_cut_mm2 = float(df_ok["재단면적_mm2"].sum())
total_cut_m2  = total_cut_mm2 / 1e6

# 원장수(네스팅 추정)
rects = []
for _, r in df_ok.iterrows():
    rects.extend([(float(r["가로(mm)"]), float(r["세로(mm)"]))] * int(r["수량"]))
nest_sheets = shelf_pack_count(rects, W_eff, H_eff, gap=gap, allow_rotate=allow_rotate) if rects else 0
if nest_sheets is None:
    nest_sheets = 0

input_area_mm2 = float(nest_sheets) * A_sheet
input_area_m2  = input_area_mm2 / 1e6

# 자투리 사용(직접 입력)
st.sidebar.divider()
st.sidebar.header("자투리 보관 사용(직접 입력)")
scrap_used_m2 = st.sidebar.number_input("자투리 사용 면적(㎡)", value=0.0, step=0.1)
scrap_include = st.sidebar.checkbox("자투리 사용을 투입면적에 포함", value=True)
if scrap_include:
    input_area_m2_eff = input_area_m2 + scrap_used_m2
else:
    input_area_m2_eff = input_area_m2

loss_area_m2 = max(input_area_m2_eff - total_cut_m2, 0.0)
loss_rate = (loss_area_m2 / input_area_m2_eff * 100) if input_area_m2_eff > 0 else 0.0

# =============================
# 2페이지 상세(제품코드/색상/부품/두께)
# =============================
group_cols = [col_prod, col_color, col_part, "두께(mm)"]
detail = (df_ok.groupby(group_cols, dropna=False, as_index=False)
          .agg(수량=("수량","sum"),
               재단면적_m2=("재단면적(m2)","sum")))

detail["재단 점유율(%)"] = (detail["재단면적_m2"] / total_cut_m2 * 100) if total_cut_m2 > 0 else 0.0

# 로스면적/로스율: 전체 로스를 재단면적 비중으로 배분(추정)
detail["로스면적(m2)"] = loss_area_m2 * (detail["재단면적_m2"] / total_cut_m2) if total_cut_m2 > 0 else 0.0
detail["로스율(%)"] = (detail["로스면적(m2)"] / (detail["로스면적(m2)"] + detail["재단면적_m2"]) * 100).fillna(0.0)

# 컬럼명 정리(표시용)
detail = detail.rename(columns={
    col_prod: "제품코드",
    col_color: "색상",
    col_part: "부품명",
    "두께(mm)": "자재두께(mm)",
    "재단면적_m2": "재단면적(m2)"
}).sort_values("재단면적(m2)", ascending=False)

# =============================
# 3페이지 분석: 어떤 부품이 로스 영향이 큰지
# =============================
an_part = (detail.groupby(["부품명"], as_index=False)
           .agg(수량=("수량","sum"),
                재단면적_m2=("재단면적(m2)","sum"),
                로스면적_m2=("로스면적(m2)","sum")))
an_part["로스 기여율(%)"] = (an_part["로스면적_m2"] / loss_area_m2 * 100) if loss_area_m2 > 0 else 0.0
an_part = an_part.sort_values("로스면적_m2", ascending=False)

# =============================
# 화면 구성(첨부 템플릿 느낌)
# =============================
tab1, tab2, tab3, tab4 = st.tabs(["1. 첫페이지", "2 페이지", "3 페이지", "오류/예외"])

with tab1:
    st.subheader("1. 첫페이지")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("구분(생산일)", str(sel_date) if sel_date else "전체")
    c2.metric("생산량 면적(㎡)", f"{total_cut_m2:,.4f}")
    c3.metric("자재투입 면적(㎡)", f"{input_area_m2_eff:,.4f}")
    c4.metric("loss율(%)", f"{loss_rate:,.4f}")

    st.write("")
    st.markdown("**자투리 보관 사용** (일자별로 직접 입력 가능)")
    st.write(f"- 자투리 사용 면적(㎡): **{scrap_used_m2:,.4f}**")
    st.write(f"- 자투리 사용을 투입면적에 포함: **{'예' if scrap_include else '아니오'}**")

    st.divider()
    st.caption("※ 원장 유효면적 기준(마진 제외) + 네스팅(선반 방식) 근사로 원장수를 추정합니다.")

with tab2:
    st.subheader("2 페이지")
    st.caption("제품코드/색상/부품명/자재두께 기준으로 재단면적, 로스면적(추정), 로스율(추정), 점유율을 표시합니다.")
    st.dataframe(detail, use_container_width=True, hide_index=True)

with tab3:
    st.subheader("3 페이지")
    st.write("→ 어떤 부품이 로스율(로스면적) 기여가 많은지 분석")
    st.dataframe(an_part.rename(columns={
        "재단면적_m2": "재단면적(㎡)",
        "로스면적_m2": "로스면적(㎡)"
    }), use_container_width=True, hide_index=True)
    st.bar_chart(an_part.set_index("부품명")["로스면적_m2"].head(20))

with tab4:
    st.subheader("오류/예외")
    if not df_invalid.empty:
        st.warning("규격 파싱 실패/수량 0 등으로 제외된 행")
        st.dataframe(df_invalid[[col_date, col_prod, col_color, col_part, col_spec, col_qty]].copy(), use_container_width=True, hide_index=True)
    if not df_oversize.empty:
        st.error("원장 유효치수 초과(배치 불가) 행")
        show = df_oversize[[col_date, col_prod, col_color, col_part, "가로(mm)", "세로(mm)", "수량", col_spec]].copy()
        st.dataframe(show, use_container_width=True, hide_index=True)
    if df_invalid.empty and df_oversize.empty:
        st.success("오류/예외가 없습니다.")

# =============================
# 다운로드(템플릿 시트 구성)
# =============================
st.divider()
st.subheader("결과 다운로드(엑셀)")

firstpage = pd.DataFrame({
    "구분":[str(sel_date) if sel_date else "전체"],
    "생산량 면적(㎡)":[total_cut_m2],
    "자재투입 면적(㎡)":[input_area_m2_eff],
    "자투리 보관 사용(㎡)":[scrap_used_m2],
    "loss율(%)":[loss_rate],
    "원장수(추정)":[nest_sheets],
    "유효 원장면적(㎡/장)":[A_sheet_m2]
})

sheets = {
    "1_첫페이지": firstpage,
    "2_상세": detail,
    "3_부품별_로스기여": an_part.rename(columns={"재단면적_m2":"재단면적(㎡)","로스면적_m2":"로스면적(㎡)"}),
}
result = build_excel(sheets)

st.download_button(
    "엑셀 다운로드",
    data=result,
    file_name=f"수율_로스율_결과_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
