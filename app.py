import io
import os
import pandas as pd
import streamlit as st

st.set_page_config(page_title="엑셀 합치기 · Row7 시작", page_icon="🧩", layout="wide")

st.title("🧩 여러 엑셀 파일을 한 번에 이어붙이기")
"""
같은 양식의 엑셀 파일 여러 개를 업로드하면, **7행(헤더)**부터 데이터를 읽고
각 파일의 **마지막 유효 행**까지만 잘라서 **하나의 표로 합친 후 다운로드**할 수 있어요.

- 헤더(열 이름) 위치: 기본 7행(사람 기준) → 0‑based 인덱스 6
- 데이터 시작: 8행부터(0‑based 인덱스 7)
- 파일마다 끝나는 행은 자동 감지(아래쪽 완전 빈 행 전까지)
- 완전히 빈 열/행은 제거
- 필요 시, 원본 파일명을 첫 열로 추가
"""

# --- 사이드바 옵션 ---
st.sidebar.header("옵션")
start_row_human = st.sidebar.number_input(
    "헤더(열 이름)가 위치한 행 번호", min_value=1, value=7, step=1,
    help="사람 기준 행 번호입니다. 기본값 7행."
)
add_filename_col = st.sidebar.checkbox("원본 파일명 열 추가", value=True)

st.sidebar.markdown("---")
st.sidebar.caption("파일은 세션 내 임시 처리되며 서버에 저장되지 않습니다.")

# --- 유틸 함수 ---
def read_and_trim_excel(file, start_row_header_idx: int) -> pd.DataFrame:
    """주어진 업로드 파일을 읽어 7행 헤더/8행 데이터 가정으로 잘라 반환.
    - start_row_header_idx: 0‑based (예: 7행 → 6)
    - 마지막 유효 행까지만 포함, 완전 빈 행/열 제거
    - 필요 시 파일명 열 추가
    """
    df_all = pd.read_excel(file, header=None)
    if df_all.empty:
        return pd.DataFrame()

    # 헤더 행 확보(병합 셀 대비 ffill)
    header_series = df_all.iloc[start_row_header_idx].astype(object)
    header_series = header_series.where(~header_series.isna(), None).ffill()
    header = header_series.tolist()

    # 데이터 본문: 헤더 다음 행부터
    body = df_all.iloc[start_row_header_idx + 1 : ].copy()
    if body.dropna(how="all").empty:
        trimmed = pd.DataFrame(columns=header)
    else:
        last_idx = body.dropna(how="all").index[-1]
        trimmed = body.loc[:last_idx].copy()
        trimmed.columns = header
        trimmed = trimmed.dropna(how="all")
        trimmed = trimmed.dropna(axis=1, how="all")  # 완전 빈 열 제거

        # 중복된 열 이름 정규화(Col, Col → Col, Col.1 ...)
        counts = {}
        new_cols = []
        for c in trimmed.columns:
            c = str(c)
            if c in counts:
                counts[c] += 1
                new_cols.append(f"{c}.{counts[c]}")
            else:
                counts[c] = 0
                new_cols.append(c)
        trimmed.columns = new_cols

    return trimmed

# --- 메인 UI ---
uploaded_files = st.file_uploader(
    "같은 양식의 엑셀 파일을 선택하세요 (여러 개 가능)",
    type=["xlsx", "xls"], accept_multiple_files=True
)

if uploaded_files:
    st.write(f"선택한 파일 수: **{len(uploaded_files)}**")

    if st.button("파일 합치기 실행", type="primary"):
        dataframes = []
        errors = []
        for f in uploaded_files:
            try:
                df = read_and_trim_excel(f, start_row_header_idx=int(start_row_human - 1))
                if add_filename_col:
                    fname = getattr(f, "name", "uploaded.xlsx")
                    df.insert(0, "source_file", os.path.basename(fname))
                dataframes.append(df)
            except Exception as e:
                errors.append((getattr(f, "name", "(이름 없음)"), str(e)))

        if errors:
            with st.expander("읽기 오류가 있었습니다 (펼치기)"):
                for name, msg in errors:
                    st.error(f"{name}: {msg}")

        if dataframes:
            merged = pd.concat(dataframes, ignore_index=True)
            st.success(f"병합 완료! 행 {merged.shape[0]} · 열 {merged.shape[1]}")
            st.dataframe(merged.head(200), use_container_width=True)

            # 다운로드용 파일 만들기 (Excel & CSV)
            excel_buf = io.BytesIO()
            with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
                merged.to_excel(writer, index=False, sheet_name="merged")
            excel_buf.seek(0)

            csv_buf = io.BytesIO()
            csv_buf.write(merged.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"))
            csv_buf.seek(0)

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    "엑셀 파일로 다운로드 (.xlsx)",
                    data=excel_buf,
                    file_name="merged_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with col2:
                st.download_button(
                    "CSV로 다운로드 (.csv)",
                    data=csv_buf,
                    file_name="merged_output.csv",
                    mime="text/csv",
                    use_container_width=True,
                )
        else:
            st.warning("유효한 데이터가 없어요. 헤더 행 번호가 맞는지 확인해주세요.")
else:
    st.info("상단에서 엑셀 파일을 올려주세요.")
