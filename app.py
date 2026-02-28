
import io
import re
import math
from datetime import datetime
import pandas as pd
import streamlit as st

st.set_page_config(page_title="목재 재단 로스율 계산/분석", layout="wide")

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

def abc_grade(cum_pct: float) -> str:
    if cum_pct <= 80:
        return "A"
    if cum_pct <= 95:
        return "B"
    return "C"

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

def build_result_excel(sheets: dict):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=name[:31])
    out.seek(0)
    return out

def rename_for_display(df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    return df.rename(columns=mapping)

st.title("목재 재단 로스율 자동 계산 & 부품명 분석 (색상 구분 v4 · 한글 컬럼)")
st.caption("엑셀 업로드 → 색상별/부품명별 로스율(면적 하한 + 네스팅 추정) + ABC 분석 + 결과 엑셀 다운로드")

with st.sidebar:
    st.header("기준 설정")
    sheet_w = st.number_input("원장 가로(W) [mm]", value=float(DEFAULT_SHEET_W), step=1.0)
    sheet_h = st.number_input("원장 세로(H) [mm]", value=float(DEFAULT_SHEET_H), step=1.0)
    margin  = st.number_input("Margin(테두리 여유) [mm]", value=float(DEFAULT_MARGIN), step=1.0)

    st.divider()
    st.header("절단/간격 설정")
    blade_mm = st.number_input("톱날 두께 [mm]", value=float(DEFAULT_BLADE_MM), step=0.1)
    kerf_mm  = st.number_input("Kerf(부품 간 여유) [mm]", value=float(DEFAULT_KERF_MM), step=1.0)
    st.caption("※ 네스팅 추정 시 (kerf + 톱날) 만큼 부품을 '팽창'시켜 간격을 보수적으로 반영합니다.")

    st.divider()
    st.header("방향/회전 설정")
    allow_rotate = st.checkbox("회전(90도) 허용", value=True)
    auto_long_to_h = st.checkbox("긴변을 세로(H)로 자동 정렬", value=True)

    st.divider()
    st.header("엑셀 컬럼 매핑 (고정: 색상)")
    col_spec = st.text_input("규격상세 컬럼", value="규격상세")
    col_qty  = st.text_input("수량(생산량) 컬럼", value="생산량")
    col_part = st.text_input("부품명 컬럼", value="부품명")
    col_color = "색상"
    st.text_input("색상 컬럼(고정)", value=col_color, disabled=True)

uploaded = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx"])
if not uploaded:
    st.info("엑셀 파일을 업로드하면 자동 계산/분석이 표시됩니다.")
    st.stop()

try:
    df = pd.read_excel(uploaded)
except Exception as e:
    st.error(f"엑셀 읽기 오류: {e}")
    st.stop()

missing_cols = [c for c in [col_spec, col_qty, col_part, col_color] if c not in df.columns]
if missing_cols:
    st.error(f"필수 컬럼이 없습니다: {missing_cols}\n엑셀 컬럼명(특히 '색상')을 확인해 주세요.")
    st.stop()

all_colors = sorted(df[col_color].dropna().astype(str).unique().tolist())
sel_colors = st.sidebar.multiselect("분석할 색상 선택(미선택=전체)", options=all_colors, default=[])
if sel_colors:
    df = df[df[col_color].astype(str).isin(sel_colors)].copy()

df["__is_summary__"] = df[col_part].isna() & df[col_spec].isna()

parsed = df[col_spec].apply(parse_spec)
df[["W_raw","H_raw","T_mm"]] = pd.DataFrame(parsed.tolist(), index=df.index)
df["Qty"] = pd.to_numeric(df[col_qty], errors="coerce").fillna(0).astype(int)

valid = df[~df["__is_summary__"]].copy()
valid["__parse_ok__"] = valid["W_raw"].notna() & valid["H_raw"].notna()
invalid = valid[~valid["__parse_ok__"]].copy()
valid_ok = valid[valid["__parse_ok__"] & (valid["Qty"] > 0)].copy()

def norm_wh(row):
    w, h = row["W_raw"], row["H_raw"]
    if auto_long_to_h and (w is not None) and (h is not None):
        if w > h:
            return pd.Series({"W_mm": h, "H_mm": w})
    return pd.Series({"W_mm": w, "H_mm": h})

valid_ok[["W_mm","H_mm"]] = valid_ok.apply(norm_wh, axis=1)

W_eff = float(sheet_w - 2 * margin)
H_eff = float(sheet_h - 2 * margin)
A_sheet = W_eff * H_eff
A_sheet_m2 = A_sheet / 1e6

def can_fit(w, h):
    if w <= W_eff and h <= H_eff:
        return True
    if allow_rotate and (h <= W_eff and w <= H_eff):
        return True
    return False

valid_ok["part_area_mm2"]  = valid_ok["W_mm"] * valid_ok["H_mm"]
valid_ok["total_area_mm2"] = valid_ok["part_area_mm2"] * valid_ok["Qty"]
valid_ok["total_area_m2"]  = valid_ok["total_area_mm2"] / 1e6

oversize = valid_ok[~valid_ok.apply(lambda r: can_fit(r["W_mm"], r["H_mm"]), axis=1)].copy()

total_pieces = int(valid_ok["Qty"].sum())
total_area_mm2 = float(valid_ok["total_area_mm2"].sum())
total_area_m2 = total_area_mm2 / 1e6

N_min_all = int(math.ceil(total_area_mm2 / A_sheet)) if A_sheet > 0 else 0
loss_min_all = ((N_min_all*A_sheet - total_area_mm2) / (N_min_all*A_sheet) * 100) if N_min_all > 0 else 0.0

gap = float(kerf_mm + blade_mm)
rects_all = []
nest_src_all = valid_ok[~valid_ok.index.isin(oversize.index)].copy()
for _, r in nest_src_all.iterrows():
    rects_all.extend([(float(r["W_mm"]), float(r["H_mm"]))] * int(r["Qty"]))

N_nest_all = shelf_pack_count(rects_all, W_eff, H_eff, gap=gap, allow_rotate=allow_rotate) if rects_all else 0
if N_nest_all is None:
    N_nest_all = 0
loss_nest_all = ((N_nest_all*A_sheet - total_area_mm2) / (N_nest_all*A_sheet) * 100) if N_nest_all > 0 else 0.0

summary_all = pd.DataFrame({
    "항목":[
        "원장 규격(mm)",
        "유효 규격(mm)",
        "Margin(mm)",
        "톱날(mm)",
        "Kerf(mm)",
        "회전 허용",
        "긴변 세로 정렬",
        "총 부재수량(합)",
        "총 부재면적(㎡)",
        "이론 최소 로스율(%)",
        "네스팅 추정 로스율(%)",
        "치수초과(배치불가) 행 수"
    ],
    "값":[
        f"{int(sheet_w)}x{int(sheet_h)}",
        f"{int(W_eff)}x{int(H_eff)}",
        float(margin),
        float(blade_mm),
        float(kerf_mm),
        "Y" if allow_rotate else "N",
        "Y" if auto_long_to_h else "N",
        total_pieces,
        total_area_m2,
        float(loss_min_all),
        float(loss_nest_all),
        int(oversize.shape[0])
    ]
})

def calc_loss_for_df(sub: pd.DataFrame):
    pieces = int(sub["Qty"].sum())
    area_mm2 = float(sub["total_area_mm2"].sum())
    area_m2 = area_mm2 / 1e6
    n_min = int(math.ceil(area_mm2 / A_sheet)) if A_sheet > 0 else 0
    loss_min = ((n_min*A_sheet - area_mm2) / (n_min*A_sheet) * 100) if n_min > 0 else 0.0
    osz = sub[~sub.apply(lambda r: can_fit(r["W_mm"], r["H_mm"]), axis=1)]
    nest_sub = sub.drop(index=osz.index, errors="ignore")
    rects = []
    for _, r in nest_sub.iterrows():
        rects.extend([(float(r["W_mm"]), float(r["H_mm"]))] * int(r["Qty"]))
    n_nest = shelf_pack_count(rects, W_eff, H_eff, gap=gap, allow_rotate=allow_rotate) if rects else 0
    if n_nest is None:
        n_nest = 0
    loss_nest = ((n_nest*A_sheet - area_mm2) / (n_nest*A_sheet) * 100) if n_nest > 0 else 0.0
    waste_m2_nest = (n_nest * A_sheet_m2 - area_m2) if n_nest > 0 else 0.0
    return pieces, area_m2, n_min, loss_min, n_nest, loss_nest, waste_m2_nest, int(osz.shape[0])

color_rows = []
for c, sub in valid_ok.groupby(col_color):
    pieces, area_m2, n_min, loss_min, n_nest, loss_nest, waste_m2, osz_n = calc_loss_for_df(sub)
    color_rows.append([str(c), pieces, area_m2, n_min, loss_min, n_nest, loss_nest, waste_m2, osz_n])

color_summary = pd.DataFrame(color_rows, columns=[
    "색상", "Qty", "Area_m2",
    "MinSheets(면적)", "LossMin_%",
    "NestSheets(추정)", "LossNest_%",
    "Waste_m2(추정)", "OversizeRows"
]).sort_values("Area_m2", ascending=False)

part_color = (valid_ok.groupby([col_color, col_part], as_index=False)
              .agg(Qty=("Qty","sum"), Area_mm2=("total_area_mm2","sum")))
part_color["Area_m2"] = part_color["Area_mm2"] / 1e6

color_area = color_summary[["색상","Area_m2","LossNest_%","Waste_m2(추정)"]].rename(columns={
    "Area_m2":"ColorArea_m2",
    "LossNest_%":"ColorLossNest_%",
    "Waste_m2(추정)":"ColorWaste_m2"
})
part_color = part_color.merge(color_area, on="색상", how="left")
part_color["AreaShareInColor_%"] = part_color["Area_m2"] / part_color["ColorArea_m2"] * 100
part_color["EstWaste_m2_inColor"] = part_color["ColorWaste_m2"] * (part_color["Area_m2"] / part_color["ColorArea_m2"])
part_color = part_color.sort_values(["색상","Area_m2"], ascending=[True, False])
part_color["RankInColor"] = part_color.groupby("색상")["Area_m2"].rank(method="first", ascending=False).astype(int)

part_group = (valid_ok.groupby([col_part], as_index=False)
              .agg(Qty=("Qty","sum"),
                   Lines=("Qty","size"),
                   Area_mm2=("total_area_mm2","sum")))
part_group["Area_m2"] = part_group["Area_mm2"] / 1e6
part_group["Area_share_%"] = part_group["Area_mm2"] / total_area_mm2 * 100 if total_area_mm2 else 0
part_group = part_group.sort_values("Area_mm2", ascending=False)
part_group["Cum_share_%"] = part_group["Area_share_%"].cumsum()
part_group["ABC"] = part_group["Cum_share_%"].apply(abc_grade)

spec_group = (valid_ok.groupby(["W_mm","H_mm"], as_index=False)
              .agg(Qty=("Qty","sum"),
                   Lines=("Qty","size"),
                   Area_mm2=("total_area_mm2","sum")))
spec_group["Area_m2"] = spec_group["Area_mm2"] / 1e6
spec_group["Area_share_%"] = spec_group["Area_mm2"] / total_area_mm2 * 100 if total_area_mm2 else 0
spec_group = spec_group.sort_values("Area_mm2", ascending=False)
spec_group["Cum_share_%"] = spec_group["Area_share_%"].cumsum()
spec_group["ABC"] = spec_group["Cum_share_%"].apply(abc_grade)

MAP_COLOR_SUMMARY = {
    "색상": "색상",
    "Qty": "수량",
    "Area_m2": "면적(㎡)",
    "MinSheets(면적)": "원장수(면적하한)",
    "LossMin_%": "로스율(면적하한, %)",
    "NestSheets(추정)": "원장수(네스팅추정)",
    "LossNest_%": "로스율(네스팅추정, %)",
    "Waste_m2(추정)": "로스면적(㎡, 추정)",
    "OversizeRows": "치수초과 행수"
}
MAP_PART_COLOR = {
    "색상": "색상",
    col_part: "부품명",
    "Qty": "수량",
    "Area_m2": "면적(㎡)",
    "AreaShareInColor_%": "색상 내 면적비중(%)",
    "ColorLossNest_%": "색상 로스율(네스팅추정, %)",
    "EstWaste_m2_inColor": "부품 로스기여면적(㎡, 추정)",
    "RankInColor": "색상 내 면적순위"
}
MAP_PART_ABC = {
    col_part: "부품명",
    "Qty": "수량",
    "Lines": "데이터 행수",
    "Area_mm2": "총면적(mm²)",
    "Area_m2": "총면적(㎡)",
    "Area_share_%": "전체 면적비중(%)",
    "Cum_share_%": "누적 면적비중(%)",
    "ABC": "ABC등급"
}
MAP_SPEC_ABC = {
    "W_mm": "가로(mm)",
    "H_mm": "세로(mm)",
    "Qty": "수량",
    "Lines": "데이터 행수",
    "Area_mm2": "총면적(mm²)",
    "Area_m2": "총면적(㎡)",
    "Area_share_%": "전체 면적비중(%)",
    "Cum_share_%": "누적 면적비중(%)",
    "ABC": "ABC등급"
}

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "요약", "색상별 로스율", "부품명×색상(로스 기여)", "부품명/규격 분석", "오류/예외"
])

with tab1:
    st.subheader("전체 요약")
    st.dataframe(summary_all, use_container_width=True, hide_index=True)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("총 부재수량", f"{total_pieces:,} 개")
    c2.metric("총 부재면적", f"{total_area_m2:,.3f} ㎡")
    c3.metric("이론 최소 로스율", f"{loss_min_all:,.2f} %")
    c4.metric("네스팅 추정 로스율", f"{loss_nest_all:,.2f} %")
    if not oversize.empty:
        st.warning("원장 유효치수에 들어가지 않는 부재가 있습니다. '오류/예외' 탭에서 확인하세요.")

with tab2:
    st.subheader("색상별 로스율(%)")
    st.caption("면적 하한(낙관적)과 네스팅 추정(근사, 실무치에 더 가까움)을 함께 표시합니다.")
    st.dataframe(rename_for_display(color_summary, MAP_COLOR_SUMMARY), use_container_width=True, hide_index=True)
    st.bar_chart(color_summary.set_index("색상")["LossNest_%"])

with tab3:
    st.subheader("부품명별 - 해당 색상의 로스율과 로스 기여(추정)")
    st.caption("‘색상 로스율’은 색상 전체 네스팅추정 로스율(%)이며, ‘부품 로스기여면적’은 색상 로스면적을 부품 면적비중으로 분배한 추정치입니다.")
    show_cols = ["색상", col_part, "Qty", "Area_m2", "AreaShareInColor_%", "ColorLossNest_%", "EstWaste_m2_inColor", "RankInColor"]
    part_color_view = part_color[show_cols].copy()
    st.dataframe(rename_for_display(part_color_view, MAP_PART_COLOR), use_container_width=True, hide_index=True)
    sel = st.selectbox("색상 선택", options=sorted(part_color["색상"].unique().tolist()))
    sub = part_color[part_color["색상"] == sel].copy().sort_values("Area_m2", ascending=False).head(20)
    st.bar_chart(sub.set_index(col_part)["Area_m2"])

with tab4:
    st.subheader("부품명 기준 집계 + ABC(전체)")
    st.caption("ABC: 누적 면적비중 기준 A(<=80%), B(<=95%), C(>95%)")
    st.dataframe(rename_for_display(part_group, MAP_PART_ABC), use_container_width=True, hide_index=True)
    st.bar_chart(part_group.head(20).set_index(col_part)["Area_m2"])

    st.divider()
    st.subheader("규격(W×H) 기준 집계 + ABC(전체)")
    st.dataframe(rename_for_display(spec_group, MAP_SPEC_ABC), use_container_width=True, hide_index=True)
    top = spec_group.head(20).copy()
    top["Spec"] = top["W_mm"].astype(int).astype(str) + "×" + top["H_mm"].astype(int).astype(str)
    st.bar_chart(top.set_index("Spec")["Area_m2"])

with tab5:
    st.subheader("규격 파싱 실패")
    if invalid.empty:
        st.success("파싱 실패 행이 없습니다.")
        inv = pd.DataFrame(columns=["색상","부품명","규격상세","생산량"])
    else:
        inv = invalid[[col_color, col_part, col_spec, col_qty]].copy()
        inv = inv.rename(columns={col_color:"색상", col_part:"부품명", col_spec:"규격상세", col_qty:"생산량"})
        st.dataframe(inv, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("원장 유효치수 초과(배치 불가)")
    if oversize.empty:
        st.success("치수 초과 부재가 없습니다.")
        osz = pd.DataFrame(columns=["색상","부품명","가로(mm)","세로(mm)","수량","규격상세"])
    else:
        osz = oversize[[col_color, col_part, "W_mm", "H_mm", "Qty", col_spec]].copy()
        osz = osz.rename(columns={col_color:"색상", col_part:"부품명", "W_mm":"가로(mm)", "H_mm":"세로(mm)", "Qty":"수량", col_spec:"규격상세"})
        st.dataframe(osz, use_container_width=True, hide_index=True)

st.divider()
st.subheader("다운로드")

sheets = {
    "요약": summary_all,
    "색상별_로스율": rename_for_display(color_summary, MAP_COLOR_SUMMARY),
    "부품명x색상_로스기여": rename_for_display(part_color_view, MAP_PART_COLOR),
    "부품명별_집계_ABC": rename_for_display(part_group, MAP_PART_ABC),
    "규격별_집계_ABC": rename_for_display(spec_group, MAP_SPEC_ABC),
    "오류_파싱실패": inv,
    "예외_치수초과": osz,
}
result_xlsx = build_result_excel(sheets)

st.download_button(
    "결과 엑셀 다운로드(한글 컬럼)",
    data=result_xlsx,
    file_name=f"loss_analysis_color_v4_kr_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
