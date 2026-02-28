
import io
import re
import math
from datetime import datetime
import pandas as pd
import streamlit as st

st.set_page_config(page_title="목재 재단 로스율 계산/분석", layout="wide")

DEFAULT_SHEET_W = 1220
DEFAULT_SHEET_H = 2440
DEFAULT_MARGIN = 10.0
DEFAULT_BLADE = 3.2
DEFAULT_KERF_GAP = 20.0

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
        return (nums[0], nums[1], nums[2] if len(nums)>=3 else None)
    return (None, None, None)

def build_result_excel(line_df, spec_group_df, part_group_df, meta_summary_df, invalid_df, oversize_df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        meta_summary_df.to_excel(writer, index=False, sheet_name="요약")
        line_df.to_excel(writer, index=False, sheet_name="라인별_계산")
        spec_group_df.to_excel(writer, index=False, sheet_name="규격별_집계")
        part_group_df.to_excel(writer, index=False, sheet_name="부품명별_집계")
        invalid_df.to_excel(writer, index=False, sheet_name="오류_파싱실패")
        oversize_df.to_excel(writer, index=False, sheet_name="예외_치수초과")
    out.seek(0)
    return out

st.title("목재 재단 로스율 자동 계산 & 부품명 분석 (Streamlit Cloud)")
st.caption("엑셀(.xlsx) 업로드 → 규격상세(W×H×T) 파싱 → 면적 기준 로스율(이론 하한) + 부품명/규격 분석 + 결과 엑셀 다운로드")

with st.sidebar:
    st.header("기준 설정")
    sheet_w = st.number_input("원장 가로(W) [mm]", value=float(DEFAULT_SHEET_W), step=1.0)
    sheet_h = st.number_input("원장 세로(H) [mm]", value=float(DEFAULT_SHEET_H), step=1.0)
    margin = st.number_input("Margin(테두리 여유) [mm]", value=float(DEFAULT_MARGIN), step=1.0)
    blade = st.number_input("톱날 두께(표시용) [mm]", value=float(DEFAULT_BLADE), step=0.1)
    kerf_gap = st.number_input("Kerf(부품 간 여유) [mm]", value=float(DEFAULT_KERF_GAP), step=1.0)

    st.divider()
    st.header("엑셀 컬럼 매핑")
    col_spec = st.text_input("규격상세 컬럼", value="규격상세")
    col_qty  = st.text_input("수량(생산량) 컬럼", value="생산량")
    col_part = st.text_input("부품명 컬럼", value="부품명")

uploaded = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

if not uploaded:
    st.info("엑셀 파일을 업로드하면 자동 계산/분석이 표시됩니다.")
    st.stop()

try:
    df = pd.read_excel(uploaded)
except Exception as e:
    st.error(f"엑셀 읽기 오류: {e}")
    st.stop()

missing_cols = [c for c in [col_spec, col_qty, col_part] if c not in df.columns]
if missing_cols:
    st.error(f"필수 컬럼이 없습니다: {missing_cols}\n사이드바의 컬럼 매핑을 수정해 주세요.")
    st.stop()

df["__is_summary__"] = df[col_part].isna() & df[col_spec].isna()

parsed = df[col_spec].apply(parse_spec)
df[["W_mm","H_mm","T_mm"]] = pd.DataFrame(parsed.tolist(), index=df.index)

df["Qty"] = pd.to_numeric(df[col_qty], errors="coerce").fillna(0).astype(int)

valid = df[~df["__is_summary__"]].copy()
valid["__parse_ok__"] = valid["W_mm"].notna() & valid["H_mm"].notna()

valid_ok = valid[valid["__parse_ok__"] & (valid["Qty"]>0)].copy()
invalid = valid[~valid["__parse_ok__"]].copy()

valid_ok["part_area_mm2"] = valid_ok["W_mm"] * valid_ok["H_mm"]
valid_ok["total_area_mm2"] = valid_ok["part_area_mm2"] * valid_ok["Qty"]
valid_ok["part_area_m2"] = valid_ok["part_area_mm2"] / 1e6
valid_ok["total_area_m2"] = valid_ok["total_area_mm2"] / 1e6

W_eff = sheet_w - 2*margin
H_eff = sheet_h - 2*margin
A_sheet = W_eff * H_eff

total_pieces = int(valid_ok["Qty"].sum())
total_area = float(valid_ok["total_area_mm2"].sum())

N_min = int(math.ceil(total_area / A_sheet)) if A_sheet > 0 else 0
loss_min_pct = ((N_min*A_sheet - total_area) / (N_min*A_sheet) * 100) if N_min > 0 else 0.0

oversize = valid_ok[(valid_ok["W_mm"]>W_eff) | (valid_ok["H_mm"]>H_eff)].copy()

spec_group = (valid_ok
              .groupby(["W_mm","H_mm"], as_index=False)
              .agg(Qty=("Qty","sum"),
                   Lines=("Qty","size"),
                   Area_mm2=("total_area_mm2","sum")))
spec_group["Area_m2"] = spec_group["Area_mm2"] / 1e6
spec_group["Area_share_%"] = spec_group["Area_mm2"] / total_area * 100 if total_area else 0
spec_group = spec_group.sort_values("Area_mm2", ascending=False)

part_group = (valid_ok
              .groupby([col_part], as_index=False)
              .agg(Qty=("Qty","sum"),
                   Lines=("Qty","size"),
                   Area_mm2=("total_area_mm2","sum")))
part_group["Area_m2"] = part_group["Area_mm2"] / 1e6
part_group["Area_share_%"] = part_group["Area_mm2"] / total_area * 100 if total_area else 0
part_group = part_group.sort_values("Area_mm2", ascending=False)

meta_summary = pd.DataFrame({
    "항목":[
        "원장 규격(mm)",
        "유효 규격(mm) (margin 적용)",
        "Margin(mm)",
        "톱날 두께(mm) (표시용)",
        "Kerf(mm) (부품 간 여유)",
        "총 라인수(유효)",
        "총 부재수량(합)",
        "총 부재면적(mm²)",
        "총 부재면적(m²)",
        "이론 최소 원장수(면적 하한)",
        "이론 최소 로스율(%)"
    ],
    "값":[
        f"{int(sheet_w)}x{int(sheet_h)}",
        f"{int(W_eff)}x{int(H_eff)}",
        float(margin),
        float(blade),
        float(kerf_gap),
        int(valid_ok.shape[0]),
        total_pieces,
        total_area,
        total_area/1e6,
        N_min,
        float(loss_min_pct)
    ]
})

tab1, tab2, tab3, tab4 = st.tabs(["요약", "부품명 분석", "규격 분석", "오류/예외"])

with tab1:
    st.subheader("로스율 요약(면적 기준)")
    st.dataframe(meta_summary, use_container_width=True, hide_index=True)
    c1, c2, c3 = st.columns(3)
    c1.metric("총 부재수량", f"{total_pieces:,} 개")
    c2.metric("총 부재면적", f"{total_area/1e6:,.3f} m²")
    c3.metric("이론 최소 로스율", f"{loss_min_pct:,.2f} %")
    if not oversize.empty:
        st.warning("유효 원장(W_eff×H_eff)에 들어가지 않는 부재가 있습니다. '오류/예외' 탭에서 확인하세요.")

with tab2:
    st.subheader("부품명 기준 집계")
    st.dataframe(part_group, use_container_width=True, hide_index=True)
    st.bar_chart(part_group.head(20).set_index(col_part)["Area_m2"])

with tab3:
    st.subheader("규격(W×H) 기준 집계")
    st.dataframe(spec_group, use_container_width=True, hide_index=True)
    top = spec_group.head(20).copy()
    top["Spec"] = top["W_mm"].astype(int).astype(str) + "×" + top["H_mm"].astype(int).astype(str)
    st.bar_chart(top.set_index("Spec")["Area_m2"])

with tab4:
    st.subheader("규격 파싱 실패")
    if invalid.empty:
        st.success("파싱 실패 행이 없습니다.")
    else:
        st.dataframe(invalid[[col_part, col_spec, col_qty]].copy(), use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("원장 유효치수 초과(배치 불가)")
    if oversize.empty:
        st.success("치수 초과 부재가 없습니다.")
    else:
        st.dataframe(oversize[[col_part, "W_mm", "H_mm", "Qty", col_spec]].copy(), use_container_width=True, hide_index=True)

st.divider()
st.subheader("다운로드")

invalid_out = invalid[[col_part, col_spec, col_qty]].copy()
oversize_out = oversize[[col_part, "W_mm", "H_mm", "Qty", col_spec]].copy()

result_xlsx = build_result_excel(
    line_df=valid_ok,
    spec_group_df=spec_group,
    part_group_df=part_group,
    meta_summary_df=meta_summary,
    invalid_df=invalid_out,
    oversize_df=oversize_out
)

st.download_button(
    "결과 엑셀 다운로드 (요약/라인별/규격별/부품명별/오류)",
    data=result_xlsx,
    file_name=f"loss_analysis_result_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
