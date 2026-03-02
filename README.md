# 수율(로스율) 프로그램 · 템플릿 형식

첨부하신 '수율프로그램 만들기.xlsx'의 구조를 따라:
- **1. 첫페이지**: 생산일(구분) 기준 요약(생산량면적 / 자재투입면적 / 자투리사용 / loss율)
- **2 페이지**: 제품코드/색상/부품명/자재두께 기준 상세(재단면적, 로스면적/로스율 추정, 점유율)
- **3 페이지**: 로스 기여가 큰 부품 분석

## 실행(로컬)
```bash
pip install -r requirements.txt
streamlit run app.py
```

## 배포(Streamlit Cloud)
- GitHub 저장소 루트에 `app.py`, `requirements.txt`, `README.md` 업로드 후 Deploy
