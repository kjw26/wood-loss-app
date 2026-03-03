
# 목재 재단 로스율 자동 계산기 (Streamlit)

## 기능
- 엑셀 업로드
- 생산일별 사용면적 / 로스면적 / 로스율 계산
- 자투리 사용 입력 시 로스율 자동 감소
- 결과 엑셀 다운로드

## 필수 컬럼
- 생산일
- 규격상세 (예: 300*500*18)
- 생산량
- 제품코드
- 색상
- 부품명

## 실행 방법

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud 배포 방법
1. GitHub에 업로드
2. Streamlit Cloud에서 저장소 연결
3. main file path: app.py
