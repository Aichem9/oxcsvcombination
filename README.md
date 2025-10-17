# 엑셀 합치기 Streamlit 앱

여러 개의 엑셀 파일을 업로드하면, **7행에 있는 헤더**를 기준으로 데이터를 읽고 각 파일의 마지막 유효 행까지 잘라 하나의 표로 합친 뒤 다운로드할 수 있습니다.

## 로컬 실행
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Community Cloud 배포
1. 이 폴더를 GitHub 새 저장소로 푸시합니다.
2. [streamlit.io](https://streamlit.io) → **Community Cloud** → **New app**.
3. 방금 만든 저장소와 브랜치 선택, **app.py** 지정 후 **Deploy**.
4. 배포 후 URL이 생성됩니다. 공유하면 누구나 웹에서 사용 가능합니다.

## 옵션
- 헤더 행(사람 기준): 기본 7행 (사이드바에서 변경 가능)
- 원본 파일명 열 추가: 기본 ON
