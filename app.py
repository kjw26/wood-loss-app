import io
import re
import math
import datetime as dt
import pandas as pd
import streamlit as st

# ======================
# 유틸
# ======================
def pick_col(cols, candidates):
    cols_set = set(cols)
    for c in candidates:
        if c in cols_set:
            return c
    return None

_spec_pat = re.compile(r"^\s*(\d+(\.\d+)?)\s*\*\s*(\d+(\.\d+)?)\s*\*\s*(\d+(\.\d+)?)\s*$")

def parse_spec(s):
    """규격상세: 가로*세로*두께"""
    if pd.isna(s):
        return None, None, None
    m = _spec_pat.match(str(s))
    if not m:
        return None, None, None
    return float(m.group(1)), float(m.group(3)), float(m.group(5))

def compute(df_raw: pd.DataFrame, cfg: dict):
    """Data / 요약(기본) 계산"""
    df = df_raw.copy()

    # 필수 컬럼(기본: 사용자가 준 ERP 예시 기반)
    col_date = pick_col(df.columns, ["생산일", "생산일자", "일자", "Date"])
    col_spec = pick_col(df.columns, ["규격상세", "규격", "사이즈", "SIZE"])
    col_qty  = pick_col(df.columns, ["생산량", "수량", "Qty", "QTY"])
    col_prod = pick_col(df.columns, ["제품코드", "품번", "품목코드", "Product", "제품"])
    col_color= pick_col(df.columns, ["색상", "컬러", "Color", "색"])
    col_part = pick_col(df.columns, ["부품명", "품명", "Part", "부품"])

    missing = [name for name, col in [
        ("생산일", col_date),
        ("규격상세", col_spec),
        ("생산량", col_qty),
        ("제품코드", col_prod),
        ("색상", col_color),
        ("부품명", col_part),
    ] if col is None]
    if missing:
        raise ValueError(f"필수 컬럼을 찾지 못했습니다: {', '.join(missing)}")

    # 날짜
    df["생산일"] = pd.to_datetime(df[col_date], errors="coerce").dt.date
    if df["생산일"].isna().any():
        raise ValueError("생산일을 날짜로 변환할 수 없는 행이 있습니다. 생산일 컬럼 값을 확인하세요.")

    # 규격상세 파싱
    parsed = df[col_spec].apply(parse_spec)
    df["폭_mm"] = parsed.apply(lambda x: x[0])
    df["길이_mm"] = parsed.apply(lambda x: x[1])
    df["두께_mm"] = parsed.apply(lambda x: x[2])
    if df[["폭_mm", "길이_mm", "두께_mm"]].isna().any().any():
        raise ValueError("규격상세 파싱 실패 행이 있습니다. 규격상세는 '가로*세로*두께' (예: 300*500*18) 형식이어야 합니다.")

    # 수량
    df["생산량"] = pd.to_numeric(df[col_qty], errors="coerce").fillna(0)

    # 그룹키
    df["제품코드"] = df[col_prod].astype(str)
    df["색상"] = df[col_color].astype(str)
    df["부품명"] = df[col_part].astype(str)

    # 유효면적
    usable_w = cfg["board_w_mm"] - 2 * cfg["margin_mm"]
    usable_h = cfg["board_h_mm"] - 2 * cfg["margin_mm"]
    usable_area_m2 = (usable_w * usable_h) / 1_000_000

    # 면적
    df["단품면적_m2"] = (df["폭_mm"] * df["길이_mm"]) / 1_000_000
    df["총면적_m2"] = df["단품면적_m2"] * df["생산량"]

    # 요약(생산일별)
    sum_df = (
        df.groupby("생산일", as_index=False)
        .agg(순부품면적_m2=("총면적_m2", "sum"))
        .sort_values("생산일")
    )
    sum_df["원장수_장"] = sum_df["순부품면적_m2"].apply(lambda x: 0 if x <= 0 else math.ceil(x / usable_area_m2))
    sum_df["사용면적_m2"] = sum_df["원장수_장"] * usable_area_m2
    sum_df["기본로스_m2"] = sum_df["사용면적_m2"] - sum_df["순부품면적_m2"]
    sum_df["자투리사용_m2"] = 0.0
    sum_df["조정로스_m2"] = (sum_df["기본로스_m2"] - sum_df["자투리사용_m2"]).clip(lower=0)
    sum_df["로스율"] = sum_df.apply(lambda r: (r["조정로스_m2"] / r["사용면적_m2"]) if r["사용면적_m2"] > 0 else 0.0, axis=1)

    return df, sum_df, usable_area_m2

def recompute_with_scrap(df_data: pd.DataFrame, sum_df_edited: pd.DataFrame):
    """자투리 입력 반영 후 요약/분석 재계산"""
    sum_df = sum_df_edited.copy()
    sum_df["조정로스_m2"] = (sum_df["기본로스_m2"] - sum_df["자투리사용_m2"]).clip(lower=0)
    sum_df["로스율"] = sum_df.apply(lambda r: (r["조정로스_m2"] / r["사용면적_m2"]) if r["사용면적_m2"] > 0 else 0.0, axis=1)

    gcols = ["제품코드", "색상", "부품명", "두께_mm"]

    # 날짜-그룹 순면적
    daily_group = (
        df_data.groupby(["생산일"] + gcols, as_index=False)
        .agg(그룹순면적_m2=("총면적_m2", "sum"))
    )
    daily_total = (
        df_data.groupby("생산일", as_index=False)
        .agg(일자순면적_m2=("총면적_m2", "sum"))
    )
    daily = daily_group.merge(daily_total, on="생산일", how="left")
    daily["일자내비중"] = daily.apply(lambda r: (r["그룹순면적_m2"] / r["일자순면적_m2"]) if r["일자순면적_m2"] > 0 else 0.0, axis=1)

    sum_map = sum_df.set_index("생산일")[["사용면적_m2", "자투리사용_m2"]]
    daily["사용면적_m2"] = daily["생산일"].map(sum_map["사용면적_m2"])
    daily["자투리사용_m2"] = daily["생산일"].map(sum_map["자투리사용_m2"])

    daily["할당사용면적_m2"] = daily["사용면적_m2"] * daily["일자내비중"]
    daily["자투리할당_m2"] = daily["자투리사용_m2"] * daily["일자내비중"]

    # 그룹별 합산
    ana_df = (
        daily.groupby(gcols, as_index=False)
        .agg(
            순면적_m2=("그룹순면적_m2", "sum"),
            할당사용면적_m2=("할당사용면적_m2", "sum"),
            자투리할당_m2=("자투리할당_m2", "sum"),
        )
        .sort_values(["제품코드", "색상", "부품명", "두께_mm"])
    )

    ana_df["조정로스율"] = ana_df.apply(
        lambda r: (max(0.0, r["할당사용면적_m2"] - r["순면적_m2"] - r["자투리할당_m2"]) / r["할당사용면적_m2"])
        if r["할당사용면적_m2"] > 0 else 0.0,
        axis=1,
    )

    total_net = ana_df["순면적_m2"].sum()
    ana_df["재단점유율"] = ana_df["순면적_m2"].apply(lambda x: (x / total_net) if total_net > 0 else 0.0)

    return sum_df, ana_df

def to_excel_bytes(df_data, df_sum, df_ana, cfg: dict, usable_area_m2: float):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        cfg_df = pd.DataFrame({
            "항목": [
                "원장_가로_mm", "원장_세로_mm", "테두리여유_mm", "날물두께_mm(커프)",
                "유효원장면적_m2"
            ],
            "값": [
                cfg["board_w_mm"], cfg["board_h_mm"], cfg["margin_mm"], cfg["kerf_mm"],
                usable_area_m2
            ]
        })
        cfg_df.to_excel(writer, index=False, sheet_name="설정")

        df_data.to_excel(writer, index=False, sheet_name="Data")

        sum_out = df_sum.copy()
        sum_out["생산일"] = pd.to_datetime(sum_out["생산일"])
        sum_out.to_excel(writer, index=False, sheet_name="요약")

        df_ana.to_excel(writer, index=False, sheet_name="분석")

    output.seek(0)
    return output.getvalue()


# ======================
# Streamlit UI
# ======================
st.set_page_config(page_title="목재 재단 로스율 자동 계산", layout="wide")
st.title("목재 재단 로스율 자동 계산 (Excel 업로드)")

# 사이드바: 옵션값 보정(요청 반영)
with st.sidebar:
    st.header("옵션(보정 가능)")
    board_w = st.number_input("원장 가로(mm)", min_value=1, value=2440, step=10)
    board_h = st.number_input("원장 세로(mm)", min_value=1, value=1220, step=10)
    margin  = st.number_input("테두리 여유(mm)", min_value=0, value=20, step=1)
    kerf    = st.number_input("날물 두께(mm)", min_value=0.0, value=3.2, step=0.1, format="%.1f")

cfg = {"board_w_mm": int(board_w), "board_h_mm": int(board_h), "margin_mm": int(margin), "kerf_mm": float(kerf)}

usable_w = cfg["board_w_mm"] - 2 * cfg["margin_mm"]
usable_h = cfg["board_h_mm"] - 2 * cfg["margin_mm"]
usable_area_m2 = (usable_w * usable_h) / 1_000_000

st.caption(f"유효 원장 면적: {usable_w}×{usable_h}mm = {usable_area_m2:.4f} m²")

uploaded = st.file_uploader("ERP 엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("엑셀을 업로드하면 ①요약 / ②분석 탭이 생성됩니다.")
    st.stop()

try:
    df_raw = pd.read_excel(uploaded, sheet_name=0)
    df_data, df_sum, usable_area_m2 = compute(df_raw, cfg)
except Exception as e:
    st.error(str(e))
    with st.expander("업로드 데이터 미리보기(상위 50행)"):
        st.dataframe(df_raw.head(50), use_container_width=True)
    st.stop()

tab1, tab2 = st.tabs(["① 요약(생산일별)", "② 분석(제품/색상/부품/두께)"])

# ======================
# 탭1: 요약
# ======================
with tab1:
    st.subheader("생산일별 요약 (자투리 입력 가능)")
    st.write("`자투리사용_m2`만 입력/수정하면 **조정로스 및 로스율(%)**이 자동 반영됩니다.")

    editable = df_sum.copy()
    show_edit = editable[[
        "생산일", "순부품면적_m2", "원장수_장", "사용면적_m2", "기본로스_m2", "자투리사용_m2"
    ]].copy()

    edited = st.data_editor(
        show_edit,
        use_container_width=True,
        hide_index=True,
        column_config={
            "생산일": st.column_config.DateColumn("생산일", disabled=True),
            "순부품면적_m2": st.column_config.NumberColumn("순부품면적(m²)", disabled=True, format="%.3f"),
            "원장수_장": st.column_config.NumberColumn("원장수(장)", disabled=True, format="%d"),
            "사용면적_m2": st.column_config.NumberColumn("사용면적(m²)", disabled=True, format="%.3f"),
            "기본로스_m2": st.column_config.NumberColumn("기본로스(m²)", disabled=True, format="%.3f"),
            "자투리사용_m2": st.column_config.NumberColumn("자투리사용(m²)", min_value=0.0, step=0.1, format="%.3f"),
        },
        key="scrap_editor",
    )

    sum_edited = editable.copy()
    sum_edited["자투리사용_m2"] = pd.to_numeric(edited["자투리사용_m2"], errors="coerce").fillna(0.0)

    df_sum2, df_ana2 = recompute_with_scrap(df_data, sum_edited)

    show_sum = df_sum2.copy()
    show_sum["로스율(%)"] = show_sum["로스율"] * 100

    st.markdown("#### 생산일별 결과")
    st.dataframe(
        show_sum[[
            "생산일", "순부품면적_m2", "원장수_장", "사용면적_m2",
            "기본로스_m2", "자투리사용_m2", "조정로스_m2", "로스율(%)"
        ]],
        use_container_width=True,
        hide_index=True,
        column_config={
            "순부품면적_m2": st.column_config.NumberColumn("순부품면적(m²)", format="%.3f"),
            "사용면적_m2": st.column_config.NumberColumn("사용면적(m²)", format="%.3f"),
            "기본로스_m2": st.column_config.NumberColumn("기본로스(m²)", format="%.3f"),
            "자투리사용_m2": st.column_config.NumberColumn("자투리사용(m²)", format="%.3f"),
            "조정로스_m2": st.column_config.NumberColumn("조정로스(m²)", format="%.3f"),
            "로스율(%)": st.column_config.NumberColumn("로스율(%)", format="%.2f"),
        },
    )

    # (1) 생산일별 합계 + (2) 전체 합계(요청 반영)
    st.markdown("#### 전체 합계")
    total_used = float(show_sum["사용면적_m2"].sum())
    total_adj_loss = float(show_sum["조정로스_m2"].sum())
    total_rate = (total_adj_loss / total_used) * 100 if total_used > 0 else 0.0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("순부품면적 합계(m²)", f"{show_sum['순부품면적_m2'].sum():.3f}")
    c2.metric("사용면적 합계(m²)", f"{total_used:.3f}")
    c3.metric("조정로스 합계(m²)", f"{total_adj_loss:.3f}")
    c4.metric("총 로스율(%)", f"{total_rate:.2f}")

# ======================
# 탭2: 분석
# ======================
with tab2:
    st.subheader("제품코드/색상/부품명/두께별 분석")
    st.write("생산일별 사용면적·자투리사용을 **해당 일자 내 순면적 비중으로 배분**하여 그룹별 로스율을 계산합니다.")

    ana_show = df_ana2.copy()
    ana_show["조정로스율(%)"] = ana_show["조정로스율"] * 100
    ana_show["재단점유율(%)"] = ana_show["재단점유율"] * 100

    st.dataframe(
        ana_show[[
            "제품코드", "색상", "부품명", "두께_mm",
            "순면적_m2", "할당사용면적_m2", "자투리할당_m2",
            "조정로스율(%)", "재단점유율(%)"
        ]],
        use_container_width=True,
        hide_index=True,
        column_config={
            "두께_mm": st.column_config.NumberColumn("두께(mm)", format="%.1f"),
            "순면적_m2": st.column_config.NumberColumn("순면적(m²)", format="%.3f"),
            "할당사용면적_m2": st.column_config.NumberColumn("할당사용면적(m²)", format="%.3f"),
            "자투리할당_m2": st.column_config.NumberColumn("자투리할당(m²)", format="%.3f"),
            "조정로스율(%)": st.column_config.NumberColumn("조정로스율(%)", format="%.2f"),
            "재단점유율(%)": st.column_config.NumberColumn("재단점유율(%)", format="%.2f"),
        },
    )

    with st.expander("원본(Data) 미리보기"):
        st.dataframe(df_data.head(200), use_container_width=True)

st.divider()

# 엑셀 다운로드(설정/요약/분석/Data 포함)
excel_bytes = to_excel_bytes(df_data, df_sum2, df_ana2, cfg, usable_area_m2)
st.download_button(
    "결과 엑셀 다운로드 (설정/요약/분석/Data)",
    data=excel_bytes,
    file_name=f"목재재단_로스율결과_{dt.date.today().isoformat()}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.caption("※ 현재 로직은 면적 기준(원장수 = 순부품면적/유효면적 올림)입니다. 실제 네스팅/절단선(커프) 손실까지 반영하려면 재단 플랜 데이터가 추가로 필요합니다.")


# 규격상세 파싱: 문자열 안에서 W*H*T 패턴을 "검색"해서 추출 (텍스트가 앞/뒤에 붙어도 인식)
# 예) "(ra 2ea) 808*1100*22(PB)" -> 808, 1100, 22
_spec_pat = re.compile(
    r"(\d+(?:\.\d+)?)\s*(?:\*|x|X|\u00D7)\s*(\d+(?:\.\d+)?)\s*(?:\*|x|X|\u00D7)\s*(\d+(?:\.\d+)?)"
)

def parse_spec(s):
    if pd.isna(s):
        return None, None, None
    m = _spec_pat.search(str(s))  # <-- 핵심: match()가 아니라 search()
    if not m:
        return None, None, None
    return float(m.group(1)), float(m.group(2)), float(m.group(3))
