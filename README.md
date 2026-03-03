# 목재 재단 로스율 자동 계산기 (Streamlit)

## 이번 업데이트(요청 반영)
- 생산일 날짜 파싱 강화(혼합 포맷, 엑셀 시리얼, 합계/빈행 등 자동 처리)
- 데이터 정합성 리포트
  - 오류행(생산일/규격상세/생산량 변환 실패, 빈행) 자동 감지
  - 오류행 목록 화면 표시 + **오류행만 별도 엑셀 다운로드**
  - 계산 시 오류행은 자동 제외
- 컬럼 자동매핑 + 필요 시 UI에서 수정 가능(드롭다운 선택)

## 실행
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud
- Main file path: `app.py`
