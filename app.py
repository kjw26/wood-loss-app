import io
import re
import math
import datetime as dt
import pandas as pd
import streamlit as st

# ======================
# Helpers
# ======================
def pick_col(cols, candidates):
    cols_set = set(cols)
    for c in candidates:
        if c in cols_set:
            return c
    return None

def normalize_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def safe_to_datetime(series: pd.Series) -> pd.Series:
    """Robust datetime parsing for mixed formats, including Excel serial dates."""
    s = series.copy()

    # common junk markers (sum rows, titles, etc.)
    junk = {"합계", "TOTAL", "Total", "total", "SUM", "sum", "전체", "TOTALS", "-"}
    s_clean = s.apply(lambda v: None if normalize_str(v) in junk else v)

    # try normal parsing first
    dt1 = pd.to_datetime(s_clean, errors="coerce", infer_datetime_format=True)

    # excel serial fallback (numbers) where dt1 is NaT
    mask = dt1.isna() & s_clean.apply(lambda v: isinstance(v, (int, float)) and pd.notna(v))
    if mask.any():
        dt1.loc[mask] = pd.to_datetime(s_clean[mask], unit="D", origin="1899-12-30", errors="coerce")

    return dt1

# 규격상세 파싱: 300*500*18 / 300x500x18 / 300X500X18 / 300×500×18
_spec_pat = re.compile(
    r"^\s*(\d+(?:\.\d+)?)\s*(?:\*|x|X|\u00D7)\s*(\d+(?:\.\d+)?)\s*(?:\*|x|X|\u00D7)\s*(\d+(?:\.\d+)?)\s*$"
)

def parse_spec(s):
    if pd.isna(s):
        return None, None, None
    m = _spec_pat.match(str(s).strip())
    if not m:
        return None, None, None
    return float(m.group(1)), float(m.group(2)), float(m.group(3))

def build_invalid_report(df_raw: pd.DataFrame, colmap: dict):
    """Return invalid_rows dataframe with reasons."""
    reasons = []

    # create a working copy with original index for traceability
    w = df_raw.copy()
    w["_row"] = w.index + 2  # +2 assuming header row at 1 in Excel

    # Date check
    dt_series = safe_to_datetime(w[colmap["date"]]) if colmap.get("date") else pd.Series([pd.NaT]*len(w))
    bad_date = dt_series.isna()

    # Spec check
    spec_series = w[colmap["spec"]] if colmap.get("spec") else pd.Series([None]*len(w))
    parsed = spec_series.apply(parse_spec)
    bad_spec = parsed.apply(lambda x: any(v is None for v in x))

    # Qty check
    qty_series = pd.to_numeric(w[colmap["qty"]], errors="coerce") if colmap.get("qty") else pd.Series([pd.NA]*len(w))
    bad_qty = qty_series.isna()

    # Identify empty rows (all NaN)
    bad_empty = w.drop(columns=["_row"], errors="ignore").isna().all(axis=1)

    for i in range(len(w)):
        rs = []
        if bad_empty.iloc[i]:
            rs.append("빈행")
        if bad_date.iloc[i]:
            rs.append("생산일 변환 실패")
        if bad_spec.iloc[i]:
            rs.append("규격상세 파싱 실패")
        if bad_qty.iloc[i]:
            rs.append("생산량 숫자 변환 실패")
        if rs:
            reasons.append(", ".join(rs))
        else:
            reasons.append("")

    w["_reason"] = reasons
    invalid = w[w["_reason"] != ""].copy()
    # Keep only helpful columns (and any mapped columns)
    keep = ["_row", "_reason"]
    for k in ["date","spec","qty","prod","color","part"]:
        c = colmap.get(k)
        if c and c in invalid.columns and c not in keep:
            keep.append(c)
    # plus a few more columns if exist (optional)
    return invalid[keep]

def clean_and_compute(df_raw: pd.DataFrame, colmap: dict, cfg: dict):
    """Create cleaned Data, base summary, usable_area."""
    df = df_raw.copy()

    # Drop fully empty rows early
    df = df[~df.isna().all(axis=1)].copy()

    # Production date
    dts = safe_to_datetime(df[colmap["date"]])
    df["생산일_dt"] = dts
    df = df[df["생산일_dt"].notna()].copy()
    df["생산일"] = df["생산일_dt"].dt.date

    # Spec parse
    parsed = df[colmap["spec"]].apply(parse_spec)
    df["폭_mm"] = parsed.apply(lambda x: x[0])
    df["길이_mm"] = parsed.apply(lambda x: x[1])
    df["두께_mm"] = parsed.apply(lambda x: x[2])
    df = df[df[["폭_mm","길이_mm","두께_mm"]].notna().all(axis=1)].copy()

    # Qty
    df["생산량"] = pd.to_numeric(df[colmap["qty"]], errors="coerce")
    df = df[df["생산량"].notna()].copy()

    # Keys
    df["제품코드"] = df[colmap["prod"]].astype(str)
    df["색상"] = df[colmap["color"]].astype(str)
    df["부품명"] = df[colmap["part"]].astype(str)

    # Usable area
    usable_w = cfg["board_w_mm"] - 2 * cfg["margin_mm"]
    usable_h = cfg["board_h_mm"] - 2 * cfg["margin_mm"]
    usable_area_m2 = (usable_w * usable_h) / 1_000_000

    # Area
    df["단품면적_m2"] = (df["폭_mm"] * df["길이_mm"]) / 1_000_000
    df["총면적_m2"] = df["단품면적_m2"] * df["생산량"]

    # Summary per day
    sum_df = (
        df.groupby("생산일", as_index=False)
          .agg(순부품면적_m2=("총면적_m2","sum"))
          .sort_values("생산일")
    )
    sum_df["기본원장수_장"] = sum_df["순부품면적_m2"].apply(lambda x: 0 if x <= 0 else math.ceil(x / usable_area_m2))

    # Input columns
    sum_df["실원장수_장"] = pd.NA
    sum_df["자투리입고_m2"] = 0.0
    sum_df["자투리사용_m2"] = 0.0

    return df, sum_df, usable_area_m2

def apply_overrides_and_inventory(sum_df_in: pd.DataFrame, usable_area_m2: float, opening_scrap_m2: float):
    sum_df = sum_df_in.copy().sort_values("생산일").reset_index(drop=True)

    sum_df["실원장수_장"] = pd.to_numeric(sum_df["실원장수_장"], errors="coerce")
    sum_df["자투리입고_m2"] = pd.to_numeric(sum_df["자투리입고_m2"], errors="coerce").fillna(0.0)
    sum_df["자투리사용_m2"] = pd.to_numeric(sum_df["자투리사용_m2"], errors="coerce").fillna(0.0)

    def effective_boards(row):
        v = row["실원장수_장"]
        if pd.notna(v) and v > 0:
            return int(v)
        return int(row["기본원장수_장"])

    sum_df["적용원장수_장"] = sum_df.apply(effective_boards, axis=1)
    sum_df["사용면적_m2"] = sum_df["적용원장수_장"] * usable_area_m2
    sum_df["기본로스_m2"] = sum_df["사용면적_m2"] - sum_df["순부품면적_m2"]

    opening = float(opening_scrap_m2 or 0.0)
    begin_list, end_list, used_eff_list = [], [], []
    current = opening
    for _, r in sum_df.iterrows():
        begin = current
        inbound = float(r["자투리입고_m2"])
        want_use = float(r["자투리사용_m2"])
        available = max(0.0, begin + inbound)

        used_eff = min(max(0.0, want_use), available)
        end = available - used_eff

        begin_list.append(begin)
        used_eff_list.append(used_eff)
        end_list.append(end)

        current = end

    sum_df["자투리기초_m2"] = begin_list
    sum_df["자투리사용_적용_m2"] = used_eff_list
    sum_df["자투리기말_m2"] = end_list

    sum_df["조정로스_m2"] = (sum_df["기본로스_m2"] - sum_df["자투리사용_적용_m2"]).clip(lower=0)
    sum_df["로스율"] = sum_df.apply(lambda r: (r["조정로스_m2"] / r["사용면적_m2"]) if r["사용면적_m2"] > 0 else 0.0, axis=1)
    sum_df["자투리사용_초과여부"] = sum_df["자투리사용_m2"] > sum_df["자투리사용_적용_m2"] + 1e-9
    return sum_df

def build_analysis(df_data: pd.DataFrame, sum_df_final: pd.DataFrame):
    gcols = ["제품코드", "색상", "부품명", "두께_mm"]

    daily_group = (
        df_data.groupby(["생산일"] + gcols, as_index=False)
               .agg(그룹순면적_m2=("총면적_m2","sum"))
    )
    daily_total = (
        df_data.groupby("생산일", as_index=False)
               .agg(일자순면적_m2=("총면적_m2","sum"))
    )
    daily = daily_group.merge(daily_total, on="생산일", how="left")
    daily["일자내비중"] = daily.apply(lambda r: (r["그룹순면적_m2"]/r["일자순면적_m2"]) if r["일자순면적_m2"]>0 else 0.0, axis=1)

    sum_map = sum_df_final.set_index("생산일")[["사용면적_m2","자투리사용_적용_m2"]]
    daily["사용면적_m2"] = daily["생산일"].map(sum_map["사용면적_m2"])
    daily["자투리사용_적용_m2"] = daily["생산일"].map(sum_map["자투리사용_적용_m2"])

    daily["할당사용면적_m2"] = daily["사용면적_m2"] * daily["일자내비중"]
    daily["자투리할당_m2"] = daily["자투리사용_적용_m2"] * daily["일자내비중"]

    ana = (
        daily.groupby(gcols, as_index=False)
             .agg(
                 순면적_m2=("그룹순면적_m2","sum"),
                 할당사용면적_m2=("할당사용면적_m2","sum"),
                 자투리할당_m2=("자투리할당_m2","sum"),
             )
             .sort_values(["제품코드","색상","부품명","두께_mm"])
    )

    ana["조정로스율"] = ana.apply(
        lambda r: (max(0.0, r["할당사용면적_m2"] - r["순면적_m2"] - r["자투리할당_m2"]) / r["할당사용면적_m2"])
        if r["할당사용면적_m2"] > 0 else 0.0,
        axis=1
    )
    total_net = ana["순면적_m2"].sum()
    ana["재단점유율"] = ana["순면적_m2"].apply(lambda x: (x/total_net) if total_net>0 else 0.0)
    return ana

def to_excel_bytes(df_data, df_sum_final, df_ana, invalid_df, cfg: dict, usable_area_m2: float):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        cfg_df = pd.DataFrame({
            "항목": ["원장_가로_mm","원장_세로_mm","테두리여유_mm","날물두께_mm(커프)","유효원장면적_m2"],
            "값": [cfg["board_w_mm"],cfg["board_h_mm"],cfg["margin_mm"],cfg["kerf_mm"],usable_area_m2]
        })
        cfg_df.to_excel(writer, index=False, sheet_name="설정")
        df_data.to_excel(writer, index=False, sheet_name="Data")
        out_sum = df_sum_final.copy()
        out_sum["생산일"] = pd.to_datetime(out_sum["생산일"])
        out_sum.to_excel(writer, index=False, sheet_name="요약")
        df_ana.to_excel(writer, index=False, sheet_name="분석")
        if invalid_df is not None and len(invalid_df) > 0:
            invalid_df.to_excel(writer, index=False, sheet_name="오류행")
    output.seek(0)
    return output.getvalue()

# ======================
# UI
# ======================
st.set_page_config(page_title="목재 재단 로스율 자동 계산", layout="wide")
st.title("목재 재단 로스율 자동 계산 (Excel 업로드)")

with st.sidebar:
    st.header("옵션(보정 가능)")
    board_w = st.number_input("원장 가로(mm)", min_value=1, value=2440, step=10)
    board_h = st.number_input("원장 세로(mm)", min_value=1, value=1220, step=10)
    margin  = st.number_input("테두리 여유(mm)", min_value=0, value=20, step=1)
    kerf    = st.number_input("날물 두께(mm)", min_value=0.0, value=3.2, step=0.1, format="%.1f")
    st.divider()
    opening_scrap = st.number_input("기초 자투리 재고(m²)", min_value=0.0, value=0.0, step=0.1, format="%.3f")

cfg = {"board_w_mm": int(board_w), "board_h_mm": int(board_h), "margin_mm": int(margin), "kerf_mm": float(kerf)}
usable_w = cfg["board_w_mm"] - 2*cfg["margin_mm"]
usable_h = cfg["board_h_mm"] - 2*cfg["margin_mm"]
usable_area_m2 = (usable_w * usable_h) / 1_000_000
st.caption(f"유효 원장 면적: {usable_w}×{usable_h}mm = {usable_area_m2:.4f} m²")

uploaded = st.file_uploader("ERP 엑셀 파일 업로드 (.xlsx)", type=["xlsx"])
if uploaded is None:
    st.info("엑셀을 업로드하면 ①요약 / ②분석 탭이 생성됩니다.")
    st.stop()

df_raw = pd.read_excel(uploaded, sheet_name=0)

# --- Auto column mapping + fallback UI
cols = list(df_raw.columns)
auto_map = {
    "date":  pick_col(cols, ["생산일","생산일자","일자","Date"]),
    "spec":  pick_col(cols, ["규격상세","규격","사이즈","SIZE"]),
    "qty":   pick_col(cols, ["생산량","수량","Qty","QTY"]),
    "prod":  pick_col(cols, ["제품코드","품번","품목코드","Product","제품"]),
    "color": pick_col(cols, ["색상","컬러","Color","색"]),
    "part":  pick_col(cols, ["부품명","품명","Part","부품"]),
}

with st.expander("컬럼 자동매핑(필요 시 수정)"):
    st.caption("자동으로 인식된 컬럼이 맞지 않으면 아래에서 수정하세요.")
    c1, c2, c3 = st.columns(3)
    with c1:
        date_col = st.selectbox("생산일 컬럼", options=cols, index=cols.index(auto_map["date"]) if auto_map["date"] in cols else 0)
        qty_col  = st.selectbox("생산량 컬럼", options=cols, index=cols.index(auto_map["qty"]) if auto_map["qty"] in cols else 0)
    with c2:
        spec_col = st.selectbox("규격상세 컬럼", options=cols, index=cols.index(auto_map["spec"]) if auto_map["spec"] in cols else 0)
        prod_col = st.selectbox("제품코드 컬럼", options=cols, index=cols.index(auto_map["prod"]) if auto_map["prod"] in cols else 0)
    with c3:
        color_col= st.selectbox("색상 컬럼", options=cols, index=cols.index(auto_map["color"]) if auto_map["color"] in cols else 0)
        part_col = st.selectbox("부품명 컬럼", options=cols, index=cols.index(auto_map["part"]) if auto_map["part"] in cols else 0)

colmap = {"date":date_col,"spec":spec_col,"qty":qty_col,"prod":prod_col,"color":color_col,"part":part_col}

# --- Data quality report
invalid_df = build_invalid_report(df_raw, colmap)
invalid_count = len(invalid_df)

if invalid_count > 0:
    st.warning(f"데이터 오류/불량 행 {invalid_count}건을 감지했습니다. 오류행은 자동 제외하고 계산합니다.")
    with st.expander("오류행 상세 보기 / 다운로드"):
        st.dataframe(invalid_df, use_container_width=True, hide_index=True)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            invalid_df.to_excel(w, index=False, sheet_name="오류행")
        buf.seek(0)
        st.download_button(
            "오류행 엑셀 다운로드",
            data=buf.getvalue(),
            file_name=f"오류행_{dt.date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# --- Compute with cleaning
try:
    df_data, df_sum_base, usable_area_m2 = clean_and_compute(df_raw, colmap, cfg)
except Exception as e:
    st.error(f"계산 중 오류: {e}")
    with st.expander("업로드 데이터 미리보기(상위 100행)"):
        st.dataframe(df_raw.head(100), use_container_width=True)
    st.stop()

tab1, tab2 = st.tabs(["① 요약(생산일별)", "② 분석(제품/색상/부품/두께)"])

with tab1:
    st.subheader("생산일별 요약")
    st.write("""
- **실원장수(장)** 입력 시 실제 원장 사용수 기준으로 사용면적/로스가 재계산됩니다.
- **자투리입고(m²)** / **자투리사용(m²)** 입력 시 자투리 재고가 날짜순으로 이월됩니다.
""")

    edit_df = df_sum_base[[
        "생산일","순부품면적_m2","기본원장수_장","실원장수_장","자투리입고_m2","자투리사용_m2"
    ]].copy()

    edited = st.data_editor(
        edit_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "생산일": st.column_config.DateColumn("생산일", disabled=True),
            "순부품면적_m2": st.column_config.NumberColumn("순부품면적(m²)", disabled=True, format="%.3f"),
            "기본원장수_장": st.column_config.NumberColumn("기본원장수(장)", disabled=True, format="%d"),
            "실원장수_장": st.column_config.NumberColumn("실원장수(장)", min_value=0, step=1),
            "자투리입고_m2": st.column_config.NumberColumn("자투리입고(m²)", min_value=0.0, step=0.1, format="%.3f"),
            "자투리사용_m2": st.column_config.NumberColumn("자투리사용(m²)", min_value=0.0, step=0.1, format="%.3f"),
        },
        key="summary_editor",
    )

    sum_in = df_sum_base.copy()
    sum_in["실원장수_장"] = edited["실원장수_장"]
    sum_in["자투리입고_m2"] = edited["자투리입고_m2"]
    sum_in["자투리사용_m2"] = edited["자투리사용_m2"]

    df_sum_final = apply_overrides_and_inventory(sum_in, usable_area_m2, opening_scrap)
    df_ana = build_analysis(df_data, df_sum_final)

    show_sum = df_sum_final.copy()
    show_sum["로스율(%)"] = show_sum["로스율"] * 100

    st.markdown("#### 생산일별 결과")
    st.dataframe(
        show_sum[[
            "생산일","순부품면적_m2","적용원장수_장","사용면적_m2",
            "기본로스_m2","자투리기초_m2","자투리입고_m2","자투리사용_적용_m2","자투리기말_m2",
            "조정로스_m2","로스율(%)"
        ]],
        use_container_width=True,
        hide_index=True,
        column_config={
            "순부품면적_m2": st.column_config.NumberColumn("순부품면적(m²)", format="%.3f"),
            "사용면적_m2": st.column_config.NumberColumn("사용면적(m²)", format="%.3f"),
            "기본로스_m2": st.column_config.NumberColumn("기본로스(m²)", format="%.3f"),
            "자투리기초_m2": st.column_config.NumberColumn("자투리기초(m²)", format="%.3f"),
            "자투리입고_m2": st.column_config.NumberColumn("자투리입고(m²)", format="%.3f"),
            "자투리사용_적용_m2": st.column_config.NumberColumn("자투리사용(적용,m²)", format="%.3f"),
            "자투리기말_m2": st.column_config.NumberColumn("자투리기말(m²)", format="%.3f"),
            "조정로스_m2": st.column_config.NumberColumn("조정로스(m²)", format="%.3f"),
            "로스율(%)": st.column_config.NumberColumn("로스율(%)", format="%.2f"),
        },
    )

    if show_sum["자투리사용_초과여부"].any():
        st.warning("일부 날짜에서 자투리사용(m²)이 가용 재고를 초과하여, 가능한 만큼만 자동 적용했습니다.")

    st.markdown("#### 전체 합계")
    total_used = float(show_sum["사용면적_m2"].sum())
    total_adj_loss = float(show_sum["조정로스_m2"].sum())
    total_rate = (total_adj_loss / total_used) * 100 if total_used > 0 else 0.0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("순부품면적 합계(m²)", f"{show_sum['순부품면적_m2'].sum():.3f}")
    c2.metric("사용면적 합계(m²)", f"{total_used:.3f}")
    c3.metric("조정로스 합계(m²)", f"{total_adj_loss:.3f}")
    c4.metric("총 로스율(%)", f"{total_rate:.2f}")

with tab2:
    st.subheader("제품코드/색상/부품명/두께별 분석")
    st.write("생산일별 사용면적·자투리사용(적용)을 해당 일자 내 순면적 비중으로 배분하여 그룹별 로스율/점유율을 계산합니다.")

    ana_show = df_ana.copy()
    ana_show["조정로스율(%)"] = ana_show["조정로스율"] * 100
    ana_show["재단점유율(%)"] = ana_show["재단점유율"] * 100

    st.dataframe(
        ana_show[[
            "제품코드","색상","부품명","두께_mm",
            "순면적_m2","할당사용면적_m2","자투리할당_m2",
            "조정로스율(%)","재단점유율(%)"
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
        }
    )

    with st.expander("정제된 Data 미리보기(상위 200행)"):
        st.dataframe(df_data.head(200), use_container_width=True)

st.divider()

excel_bytes = to_excel_bytes(df_data, df_sum_final, df_ana, invalid_df, cfg, usable_area_m2)
st.download_button(
    "결과 엑셀 다운로드 (설정/요약/분석/Data/오류행)",
    data=excel_bytes,
    file_name=f"목재재단_로스율결과_{dt.date.today().isoformat()}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.caption("※ 계산은 면적기반(원장수=순부품면적/유효면적 올림) + 실원장수 우선 적용. 커프(날물) 기반 절단 손실은 재단 플랜 데이터가 있어야 정밀 반영 가능합니다.")
