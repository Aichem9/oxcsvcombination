import io
import os
import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook

st.set_page_config(page_title='엑셀 합치기 · Row7 시작', page_icon='🧩', layout='wide')

st.title('🧩 여러 엑셀 파일을 한 번에 이어붙이기')
st.markdown("""
요구사항 반영:
1) **첫 번째 파일의 1~6행(양식)**을 그대로 유지해서 결과 엑셀의 맨 위에 배치합니다.
2) **첫 번째 파일에서 데이터가 끝난 다음 행(지표 설명)**을 병합 결과의 **맨 마지막 행**으로 보냅니다.

- 기본 가정: 헤더는 7행(사람 기준), 데이터는 8행부터 시작.
- 각 파일의 마지막 유효 행은 자동 감지.
- 완전히 빈 열/행 제거, 중복 열 이름 정규화.
""")

# --- 사이드바 옵션 ---
st.sidebar.header('옵션')
start_row_human = st.sidebar.number_input(
    '헤더(열 이름)가 위치한 행 번호', min_value=1, value=7, step=1,
    help='사람 기준 행 번호입니다. 기본값 7행.'
)
add_filename_col = st.sidebar.checkbox('원본 파일명 열 추가', value=True)

st.sidebar.markdown('---')
st.sidebar.caption('파일은 세션 내 임시 처리되며 서버에 저장되지 않습니다.')

# --- 유틸 함수(표 병합용) ---
def read_and_trim_excel(file, start_row_header_idx: int) -> pd.DataFrame:
    df_all = pd.read_excel(file, header=None)
    if df_all.empty:
        return pd.DataFrame(), df_all

    header_series = df_all.iloc[start_row_header_idx].astype(object)
    header_series = header_series.where(~header_series.isna(), None).ffill()
    header = header_series.tolist()

    body = df_all.iloc[start_row_header_idx + 1 : ].copy()
    if body.dropna(how='all').empty:
        trimmed = pd.DataFrame(columns=header)
    else:
        last_idx = body.dropna(how='all').index[-1]
        trimmed = body.loc[:last_idx].copy()
        trimmed.columns = header
        trimmed = trimmed.dropna(how='all')
        trimmed = trimmed.dropna(axis=1, how='all')
        counts = {}
        new_cols = []
        for c in trimmed.columns:
            c = str(c)
            if c in counts:
                counts[c] += 1
                new_cols.append(f'{c}.{counts[c]}')
            else:
                counts[c] = 0
                new_cols.append(c)
        trimmed.columns = new_cols
    return trimmed, df_all

def detect_description_row_index(df_all: pd.DataFrame, start_row_header_idx: int):
    body = df_all.iloc[start_row_header_idx + 1 : ]
    if body.dropna(how='all').empty:
        return None
    last_idx = body.dropna(how='all').index[-1]
    desc_idx = last_idx + 1
    if desc_idx < len(df_all):
        return desc_idx
    return None

def copy_first_six_rows(src_wb, dst_wb):
    src_ws = src_wb.active
    dst_ws = dst_wb.active

    max_col = src_ws.max_column
    # 값 복사 (1~6행)
    for r in range(1, 7):
        for c in range(1, max_col + 1):
            dst_ws.cell(row=r, column=c, value=src_ws.cell(row=r, column=c).value)

    # 병합 복사 (1~6행에 걸친 범위만)
    for mr in src_ws.merged_cells.ranges:
        min_row, min_col, max_row, max_col = mr.min_row, mr.min_col, mr.max_row, mr.max_col
        if max_row <= 6:
            dst_ws.merge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)

    # 열 너비 복사
    for col_letter, dim in src_ws.column_dimensions.items():
        if dim.width:
            dst_ws.column_dimensions[col_letter].width = dim.width

# --- 메인 UI ---
uploaded_files = st.file_uploader(
    '같은 양식의 엑셀 파일을 선택하세요 (여러 개 가능)',
    type=['xlsx', 'xls'], accept_multiple_files=True
)

if uploaded_files:
    st.write(f'선택한 파일 수: **{len(uploaded_files)}**')

    if st.button('파일 합치기 실행', type='primary'):
        dataframes = []
        errors = []
        first_file_stream_copy = None
        desc_row_values = None

        for idx, f in enumerate(uploaded_files):
            try:
                raw_bytes = f.read()
                f.seek(0)
                df, df_all = read_and_trim_excel(f, start_row_header_idx=int(start_row_human - 1))

                if add_filename_col:
                    fname = getattr(f, 'name', 'uploaded.xlsx')
                    df.insert(0, 'source_file', os.path.basename(fname))
                dataframes.append(df)

                if idx == 0:
                    first_file_stream_copy = io.BytesIO(raw_bytes)
                    desc_idx = detect_description_row_index(df_all, int(start_row_human - 1))
                    if desc_idx is not None:
                        desc_row_values = df_all.iloc[desc_idx].tolist()
            except Exception as e:
                errors.append((getattr(f, 'name', '(이름 없음)'), str(e)))

        if errors:
            with st.expander('읽기 오류가 있었습니다 (펼치기)'):
                for name, msg in errors:
                    st.error(f'{name}: {msg}')

        if dataframes:
            merged = pd.concat(dataframes, ignore_index=True)
            st.success(f'병합 완료! 행 {merged.shape[0]} · 열 {merged.shape[1]}')
            st.dataframe(merged.head(200), use_container_width=True)

            # 엑셀 내보내기(첫 파일 1~6행 보존 + 헤더 7행 + 데이터 + 마지막에 설명 행)
            out_buf = io.BytesIO()
            wb_out = Workbook()
            ws_out = wb_out.active

            if first_file_stream_copy is not None:
                try:
                    wb_src = load_workbook(first_file_stream_copy, data_only=True)
                    copy_first_six_rows(wb_src, wb_out)
                except Exception:
                    pass

            start_row = 7
            for j, col_name in enumerate(merged.columns, start=1):
                ws_out.cell(row=start_row, column=j, value=str(col_name))
            for i in range(len(merged)):
                for j, col_name in enumerate(merged.columns, start=1):
                    ws_out.cell(row=start_row + i + 1, column=j, value=merged.iat[i, j-1])

            if desc_row_values is not None:
                append_row = start_row + 1 + len(merged)
                for j, val in enumerate(desc_row_values, start=1):
                    ws_out.cell(row=append_row, column=j, value=val)

            wb_out.save(out_buf)
            out_buf.seek(0)

            csv_buf = io.BytesIO()
            csv_buf.write(merged.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig'))
            csv_buf.seek(0)

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    '엑셀 파일로 다운로드 (.xlsx) — 템플릿 상단 유지 & 설명 행 하단 이동',
                    data=out_buf,
                    file_name='merged_with_template_and_desc.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True,
                )
            with col2:
                st.download_button(
                    'CSV로 다운로드 (.csv)',
                    data=csv_buf,
                    file_name='merged_output.csv',
                    mime='text/csv',
                    use_container_width=True,
                )
        else:
            st.warning('유효한 데이터가 없어요. 헤더 행 번호가 맞는지 확인해주세요.')
else:
    st.info('상단에서 엑셀 파일을 올려주세요.')
