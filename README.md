# 목재 재단 로스율 계산 & 부품명 분석 (Streamlit Cloud 배포용)

## 기본값(요청 반영)
- 원장: 1220×2440 (mm)
- 톱날 두께: 3.2mm (표시용)
- Kerf: 20mm (부품 간 여유 값으로 설정)
- 집계 기준: 부품명

---

## 1) GitHub에 올리기
1. GitHub에서 새 저장소(repository) 생성
2. 이 폴더의 파일 3개(`app.py`, `requirements.txt`, `README.md`)를 저장소 루트에 업로드(커밋)

---

## 2) Streamlit Community Cloud 배포
1. Streamlit Community Cloud에서 로그인 (GitHub 연동)
2. **New app** → GitHub 저장소 선택
3. 설정:
   - Repository: (방금 만든 repo)
   - Branch: `main`
   - Main file path: `app.py`
4. Deploy 클릭 → 배포 완료 후 `https://xxxx.streamlit.app` 주소가 생성됩니다.

---

## 3) 사용 방법
- 웹 주소로 접속 → 엑셀(.xlsx) 업로드
- 탭에서 요약/부품명/규격/오류 확인
- 하단에서 결과 엑셀 다운로드

---

## 4) 컬럼명이 다를 때
사이드바에서 다음 매핑을 바꾸세요:
- 규격상세 컬럼
- 수량(생산량) 컬럼
- 부품명 컬럼
